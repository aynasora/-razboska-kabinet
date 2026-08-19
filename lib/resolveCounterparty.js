// Поиск и (по явному действию бухгалтера) создание контрагента в
// справочнике 1С. Перенесено из server.js без изменений логики.
const { normalizeCounterpartyName } = require('./shared');

// Название поля с БИН/ИИН отличается между конфигурациями 1С. Вместо
// одного запроса со всеми вариантами сразу (он падает целиком, если хотя
// бы одного поля нет — «Сегмент пути БИН не найден»), пробуем варианты
// по очереди и запоминаем то, которое сработало.
const BIN_FIELD_CANDIDATES = ['ИИН', 'БИН', 'ИНН', 'БИН_ИИН', 'ИИН_БИН', 'ИННЮЛ', 'ИННФЛ', 'КодПоОКПО', 'РегистрационныйНомер'];
let workingBinField = null; // определяется один раз за время работы сервера

async function findCounterpartyByBin(bin, settings, name) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  // БИН/ИИН из выписки иногда приходит с пробелами, апострофами или
  // текстовым форматированием Excel ("123 456 789 012", "'123456789012") —
  // а в справочнике 1С хранится как чистая строка цифр. Сравнение "как есть"
  // в таких случаях просто не находит существующего контрагента.
  const normalizedBin = String(bin || '').replace(/\D/g, '');
  console.log(`[findCounterpartyByBin] ищу контрагента: БИН/ИИН="${bin}" -> "${normalizedBin}", название="${name || ''}"`);

  async function tryField(field) {
    const filter = encodeURIComponent(`${field} eq '${normalizedBin}'`);
    const url = `${base}/Catalog_Контрагенты?$format=json&$filter=${filter}&$top=1`;
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) return { fieldExists: false };
    const data = await response.json().catch(() => null);
    if (!data) return { fieldExists: false };
    const found = (data.value || [])[0] || null;
    console.log(`[findCounterpartyByBin] поле ${field}: ${found ? 'найден ' + found.Description : 'не найден'}`);
    return { fieldExists: true, found };
  }

  // Запасной поиск по названию — на случай, если БИН в выписке не
  // считался или не совпал по формату, а контрагент в справочнике на
  // самом деле есть. ВАЖНО: раньше здесь брался первый попавшийся
  // результат ($top=1) — если в справочнике несколько похожих записей,
  // легко было тихо выбрать не того контрагента. Теперь: сначала пробуем
  // точное совпадение по названию; если его нет — смотрим частичные
  // совпадения, и если их больше одного, ЧЕСТНО сообщаем о неоднозначности,
  // а не гадаем.
  async function fetchByFilter(filter, top) {
    const url = `${base}/Catalog_Контрагенты?$format=json&$filter=${encodeURIComponent(filter)}&$top=${top}`;
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) return null;
    const data = await response.json().catch(() => null);
    return (data && data.value) || [];
  }

  async function tryByName() {
    if (!name || name.trim().length < 3) return null;
    const cleanName = name.trim().replace(/'/g, "''");

    // 1. Точное совпадение по исходному названию — самый надёжный вариант
    const exactList = await fetchByFilter(`Description eq '${cleanName}'`, 2);
    if (exactList && exactList.length === 1) return { match: exactList[0] };
    if (exactList && exactList.length > 1) return { ambiguous: true, options: exactList };

    // 2. Частичное совпадение по исходному названию
    const partialList = await fetchByFilter(`substringof('${cleanName}', Description)`, 10);
    let candidates = partialList || [];

    // 3. Если ничего не нашли — та же попытка, но по "ядру" названия без
    // организационно-правовой формы (ТОО/ИП/АО/…) и кавычек: банковская
    // выписка и справочник 1С часто пишут форму по-разному.
    if (candidates.length === 0) {
      const normalizedTarget = normalizeCounterpartyName(name);
      if (normalizedTarget && normalizedTarget.length >= 3) {
        const wide = await fetchByFilter(`substringof('${normalizedTarget.replace(/'/g, "''")}', Description)`, 15);
        candidates = (wide || []).filter(o => normalizeCounterpartyName(o.Description) === normalizedTarget);
        // Если точных совпадений "ядра" не набралось — берём как есть, лучше
        // честно показать неоднозначность, чем не найти существующего контрагента.
        if (candidates.length === 0) candidates = wide || [];
      }
    }

    if (candidates.length === 0) return null;
    if (candidates.length === 1) return { match: candidates[0] };
    return { ambiguous: true, options: candidates };
  }

  async function resolveByName() {
    const r = await tryByName();
    if (!r) return null;
    if (r.ambiguous) {
      const err = new Error(
        `По имени "${name}" в справочнике нашлось ${r.options.length} разных контрагентов ` +
        `(${r.options.map(o => o.Description).join(', ')}) — не могу понять, какой из них правильный. ` +
        `Сопоставьте контрагента вручную для этой операции.`
      );
      err.ambiguousCounterparty = r.options;
      throw err;
    }
    return r.match;
  }

  // Если рабочее поле уже определено — сразу используем его.
  if (workingBinField && normalizedBin) {
    const r = await tryField(workingBinField);
    if (r.fieldExists) {
      if (r.found) return r.found;
      return await resolveByName(); // БИН не совпал — пробуем по названию
    }
    workingBinField = null; // поле перестало работать — определим заново
  }

  if (!normalizedBin) return await resolveByName(); // БИН в выписке пуст — сразу по названию

  for (const field of BIN_FIELD_CANDIDATES) {
    const r = await tryField(field);
    if (r.fieldExists) {
      workingBinField = field; // запомнили — дальше будет быстро
      if (r.found) return r.found;
      return await resolveByName(); // поле рабочее, но по БИН не нашли — пробуем по названию
    }
  }
  // Ни одно поле с БИН не сработало вообще — прежде чем сдаться, всё
  // равно пробуем по названию.
  return await resolveByName();
}

// Свободный ручной поиск контрагента по части названия — используется, когда
// бухгалтер сам проверяет "а нет ли он уже в 1С", прежде чем соглашаться на
// создание нового. В отличие от tryByName() внутри findCounterpartyByBin (та
// возвращает ошибку при неоднозначности), здесь мы, наоборот, ХОТИМ увидеть
// все подходящие варианты списком, чтобы человек выбрал нужный сам.
async function searchCounterpartyByText(query, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const cleanQuery = query.replace(/'/g, "''");
  const filter = encodeURIComponent(`substringof('${cleanQuery}', Description)`);
  const url = `${base}/Catalog_Контрагенты?$format=json&$filter=${filter}&$select=Ref_Key,Description&$top=20`;
  const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`1С ответила ${response.status}: ${text.slice(0, 200)}`);
  }
  const data = await response.json().catch(() => null);
  return (data && data.value) || [];
}

// Создаёт нового контрагента в справочнике 1С. Вызывается ТОЛЬКО по явному
// клику бухгалтера на кнопку "Создать нового" — эндпоинт перед вызовом этой
// функции всегда сначала перепроверяет по нормализованному БИН/ИИН и
// названию, что контрагента точно ещё нет.
async function createCounterpartyInOnec(op, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const normalizedBin = String(op.bin || '').replace(/\D/g, '');
  const description = String(op.counterparty || normalizedBin || 'Без названия').trim();

  // Имя поля с БИН/ИИН берём то, которое реально работает в вашей базе
  // (определяется при поиске контрагента) — если ещё не определено,
  // используем ИИН как самый частый вариант для казахстанских конфигураций.
  const binField = workingBinField || 'ИИН';
  const payload = { Description: description };
  if (normalizedBin) payload[binField] = normalizedBin;
  console.log(`[createCounterpartyInOnec] создаю контрагента: "${description}", ${binField}="${normalizedBin}"`);

  const response = await fetch(`${base}/Catalog_Контрагенты?$format=json`, {
    method: 'POST',
    headers: { Authorization: `Basic ${auth}`, 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`1С ответила ${response.status}: ${text.slice(0, 300)}`);
  }
  return response.json().catch(() => ({}));
}

module.exports = {
  findCounterpartyByBin,
  searchCounterpartyByText,
  createCounterpartyInOnec,
};
