// Поиск и (при необходимости) создание договора "Без договора" в 1С.
// Перенесено из server.js без изменений логики.

// Пытается найти номер договора прямо в тексте назначения платежа —
// банки/бухгалтеры обычно пишут "по дог. №...", "договор №...", "дог. N..."
function extractContractNumber(purpose) {
  if (!purpose) return null;
  const match = purpose.match(/до?г(?:овор)?\.?\s*(?:№|N|no\.?)?\s*([\w\-\/]{2,})/i);
  return match ? match[1] : null;
}

const CONTRACT_CATALOG_CANDIDATES = ['Catalog_ДоговорыКонтрагентов', 'Catalog_Договоры'];
const NO_CONTRACT_NAME = 'Без договора';

// Имя поля-владельца в справочнике договоров тоже отличается между
// конфигурациями (Владелец_Key vs Контрагент_Key). Раньше оба поля
// проверялись ОДНИМ запросом через "or" — но если хотя бы одного из двух
// полей нет в конкретном справочнике, OData 1С отклоняет весь запрос
// целиком (ошибка парсинга фильтра), и мы тихо теряли договор, который на
// самом деле есть. Теперь, как и с БИН/ИИН, пробуем поля ПО ОДНОМУ и
// запоминаем рабочую комбинацию "справочник + поле".
const CONTRACT_OWNER_FIELD_CANDIDATES = ['Владелец_Key', 'Контрагент_Key'];
let workingContractCatalog = null;
let workingContractOwnerField = null;

async function fetchContractsByOwner(counterpartyKey, settings, extraFilter) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  async function tryCombo(catalog, field) {
    let filter = `${field} eq guid'${counterpartyKey}'`;
    if (extraFilter) filter += ` and ${extraFilter}`;
    const url = `${base}/${catalog}?$format=json&$filter=${encodeURIComponent(filter)}&$select=Ref_Key,Description,Number&$top=20`;
    console.log(`[findContractForCounterparty] запрос: ${catalog}.${field}${extraFilter ? ' + ' + extraFilter : ''}`);
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) return null; // это поле/справочник не подошли — пробуем другую комбинацию
    const data = await response.json().catch(() => null);
    return data ? (data.value || []) : null;
  }

  if (workingContractCatalog && workingContractOwnerField) {
    const list = await tryCombo(workingContractCatalog, workingContractOwnerField);
    if (list !== null) return { catalog: workingContractCatalog, list };
    workingContractCatalog = null;
    workingContractOwnerField = null; // рабочая комбинация перестала работать — определим заново
  }
  for (const catalog of CONTRACT_CATALOG_CANDIDATES) {
    for (const field of CONTRACT_OWNER_FIELD_CANDIDATES) {
      const list = await tryCombo(catalog, field);
      if (list !== null) {
        workingContractCatalog = catalog;
        workingContractOwnerField = field;
        return { catalog, list };
      }
    }
  }
  console.log('[findContractForCounterparty] ни одна комбинация справочник+поле не сработала — проверьте название справочника договоров в вашей 1С');
  return null;
}

// Ищет договор(ы) контрагента в 1С. Если найден ровно один — используем
// его автоматически. Если несколько — не гадаем, а помечаем операцию как
// требующую ручного выбора договора. Если в тексте назначения нашёлся
// номер договора — в первую очередь пытаемся сопоставить именно по нему.
async function findContractForCounterparty(counterpartyKey, purposeText, settings) {
  if (!counterpartyKey) return { status: 'none' };
  const result = await fetchContractsByOwner(counterpartyKey, settings);
  if (!result || result.list.length === 0) return { status: 'none' };
  const { list } = result;

  if (list.length === 1) {
    console.log(`[findContractForCounterparty] найден единственный договор: ${list[0].Description || list[0].Number}`);
    return { status: 'matched', key: list[0].Ref_Key, name: list[0].Description || list[0].Number };
  }

  // Несколько договоров — пробуем сопоставить по номеру, найденному в назначении платежа
  const hint = extractContractNumber(purposeText);
  if (hint) {
    const byNumber = list.find(d =>
      (d.Number && d.Number.includes(hint)) || (d.Description && d.Description.includes(hint))
    );
    if (byNumber) {
      return { status: 'matched', key: byNumber.Ref_Key, name: byNumber.Description || byNumber.Number };
    }
  }
  console.log(`[findContractForCounterparty] неоднозначно: найдено ${list.length} договоров`);
  return { status: 'ambiguous', options: list.map(d => ({ key: d.Ref_Key, name: d.Description || d.Number })) };
}

// Ищет договор конкретного контрагента с названием "Без договора" —
// отдельным точным запросом (не полагаемся только на findContractForCounterparty,
// т.к. если у контрагента больше одного договора, "Без договора" может
// потеряться среди неоднозначных вариантов).
async function findNoContractRecord(counterpartyKey, settings) {
  const result = await fetchContractsByOwner(counterpartyKey, settings, `Description eq '${NO_CONTRACT_NAME}'`);
  const found = result && result.list[0];
  if (found) {
    console.log(`[findNoContractRecord] найден договор "Без договора": ${found.Ref_Key}`);
    return { catalog: result.catalog, key: found.Ref_Key, name: found.Description };
  }
  console.log('[findNoContractRecord] договор "Без договора" не найден для этого контрагента');
  return null;
}

// Полная логика подбора договора для «Оплата поставщику»:
//   1) если у контрагента найден ровно один договор (или однозначно по
//      номеру в назначении платежа) — используем его;
//   2) если договоров несколько и не определить — статус "ambiguous",
//      выбор за бухгалтером;
//   3) если договоров нет вообще — отдельно ищем именно "Без договора";
//   4) если и его нет — сообщаем статус "need_create_no_contract" (создание —
//      отдельным шагом, см. createNoContractInOnec).
async function resolveSupplierContract(counterpartyKey, purposeText, settings) {
  if (!counterpartyKey) return { status: 'none' };
  const primary = await findContractForCounterparty(counterpartyKey, purposeText, settings);
  if (primary.status === 'matched' || primary.status === 'ambiguous') return primary;

  const noContract = await findNoContractRecord(counterpartyKey, settings);
  if (noContract) {
    return { status: 'matched', key: noContract.key, name: noContract.name };
  }
  return { status: 'need_create_no_contract' };
}

// Создаёт в справочнике договоров запись "Без договора" для контрагента.
async function createNoContractInOnec(counterpartyKey, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  async function tryPost(catalog, field) {
    const payload = { Description: NO_CONTRACT_NAME, [field]: counterpartyKey };
    console.log(`[createNoContractInOnec] создаю "${NO_CONTRACT_NAME}" в ${catalog}.${field}`);
    const response = await fetch(`${base}/${catalog}?$format=json`, {
      method: 'POST',
      headers: { Authorization: `Basic ${auth}`, 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!response.ok) return null;
    return response.json().catch(() => ({}));
  }

  // Если уже знаем рабочую комбинацию справочник+поле — используем её сразу.
  if (workingContractCatalog && workingContractOwnerField) {
    const created = await tryPost(workingContractCatalog, workingContractOwnerField);
    if (created) return created;
  }
  for (const catalog of CONTRACT_CATALOG_CANDIDATES) {
    for (const field of CONTRACT_OWNER_FIELD_CANDIDATES) {
      const created = await tryPost(catalog, field);
      if (created) {
        workingContractCatalog = catalog;
        workingContractOwnerField = field;
        return created;
      }
    }
  }
  throw new Error('Не удалось создать договор "Без договора" — проверьте название справочника договоров и поля владельца в вашей конфигурации 1С');
}

module.exports = {
  NO_CONTRACT_NAME,
  extractContractNumber,
  findContractForCounterparty,
  findNoContractRecord,
  resolveSupplierContract,
  createNoContractInOnec,
};
