// Поиск GUID статьи ДДС и счёта учёта в 1С, плюс "самообучение" по истории
// уже созданных документов контрагента. Перенесено из server.js без
// изменений логики.

// Ищет статью ДДС в справочнике 1С по точному названию (например,
// «Аренда» или «Расчёты с поставщиками и подрядчиками») и возвращает её
// GUID. Нужно, когда категория пришла из текстового правила или ручной
// правки, а не из уже готового документа в 1С.
async function findCategoryKeyByName(name, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const candidates = ['Catalog_СтатьиДвиженияДенежныхСредств', 'Catalog_СтатьиДДС'];
  const cleanName = String(name).trim().replace(/'/g, "''");
  console.log(`[findCategoryKeyByName] ищу статью ДДС: "${name}"`);

  async function tryUrl(url) {
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) return null;
    const data = await response.json().catch(() => null);
    return (data && data.value) || [];
  }

  for (const catalog of candidates) {
    // 1. Точное совпадение по названию (самый надёжный вариант)
    const exactUrl = `${base}/${catalog}?$format=json&$filter=${encodeURIComponent(`Description eq '${cleanName}'`)}&$select=Ref_Key,Description&$top=1`;
    let list = await tryUrl(exactUrl);
    if (list === null) continue; // этого справочника нет под таким именем — пробуем следующий
    if (list.length) {
      console.log(`[findCategoryKeyByName] найдено точным совпадением в ${catalog}: ${list[0].Description}`);
      return list[0].Ref_Key;
    }

    // 2. Точное совпадение иногда не срабатывает из-за разного написания
    // "е"/"ё" ("Расчеты"/"Расчёты") или лишних пробелов — пробуем частичное
    // совпадение по буквам без "ё", это покрывает оба варианта написания.
    const looseName = cleanName.replace(/ё/gi, m => (m === 'ё' ? 'е' : 'Е'));
    const looseUrl = `${base}/${catalog}?$format=json&$filter=${encodeURIComponent(`substringof('${looseName}', Description)`)}&$select=Ref_Key,Description&$top=3`;
    list = await tryUrl(looseUrl);
    if (list && list.length) {
      console.log(`[findCategoryKeyByName] найдено частичным совпадением в ${catalog}: ${list[0].Description}`);
      return list[0].Ref_Key;
    }
  }
  console.log(`[findCategoryKeyByName] статья ДДС "${name}" не найдена ни в одном справочнике-кандидате`);
  return null;
}

// "Счёт учёта (БУ)" в табличной части документа — это ссылка на план счетов
// (Chart of Accounts), а не просто текст "1251": нужно найти GUID счёта по
// его коду. Название самого объекта плана счетов в OData отличается между
// конфигурациями 1С — пробуем частые варианты и запоминаем рабочий, как и с
// полем БИН/ИИН. Счёт может храниться у нас как пара "1251/3510"
// (счёт расчётов/счёт авансов) — для табличной части берём основную часть
// до "/".
const CHART_OF_ACCOUNTS_CANDIDATES = ['ChartOfAccounts_Хозрасчетный', 'ChartOfAccounts_Хозрасчетный2', 'ChartOfAccounts_Управленческий'];
let workingChartOfAccounts = null;

async function findAccountKeyByCode(code, settings) {
  if (!code) return null;
  const primaryCode = String(code).split('/')[0].trim();
  if (!primaryCode) return null;
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  async function tryChart(chart) {
    const filter = encodeURIComponent(`Code eq '${primaryCode}'`);
    const url = `${base}/${chart}?$format=json&$filter=${filter}&$select=Ref_Key&$top=1`;
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) return { exists: false };
    const data = await response.json().catch(() => null);
    if (!data) return { exists: false };
    return { exists: true, found: (data.value || [])[0] || null };
  }

  if (workingChartOfAccounts) {
    const r = await tryChart(workingChartOfAccounts);
    if (r.exists) return r.found ? r.found.Ref_Key : null;
    workingChartOfAccounts = null;
  }
  for (const chart of CHART_OF_ACCOUNTS_CANDIDATES) {
    const r = await tryChart(chart);
    if (r.exists) {
      workingChartOfAccounts = chart;
      return r.found ? r.found.Ref_Key : null;
    }
  }
  return null; // ни один вариант плана счетов не подошёл — не страшно, просто не заполним это поле
}

async function findHistoricalCategory(counterpartyKey, amount, settings, currentPurpose) {
  if (!counterpartyKey) return null;
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const docType = amount >= 0 ? 'Document_ПлатежноеПоручениеВходящее' : 'Document_ПлатежноеПоручениеИсходящее';

  // Статья ДДС в реальности лежит ВНУТРИ табличной части "РасшифровкаПлатежа",
  // а не на уровне самого документа — поэтому обязательно разворачиваем её
  // через $expand, иначе эти поля просто не приедут в ответе.
  //
  // Смотрим НЕСКОЛЬКО последних документов (не только самый свежий), потому
  // что у одного контрагента может быть несколько РАЗНЫХ типов операций
  // (например, разные виды переводов) — берём не просто "любой заполненный",
  // а тот, чьё назначение платежа больше всего похоже на текущую операцию.
  const filter = encodeURIComponent(`Контрагент_Key eq guid'${counterpartyKey}'`);
  const url = `${base}/${docType}?$format=json&$filter=${filter}&$orderby=Date desc&$top=15&$expand=РасшифровкаПлатежа`;

  const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
  if (!response.ok) return null;
  const data = await response.json().catch(() => null);
  const docs = (data && data.value) || [];

  // Похожесть текста назначения платежа: доля общих слов (без учёта
  // регистра), длиной от 3 букв — так короткие союзы/предлоги не мешают.
  function wordSet(text) {
    return new Set(String(text || '').toLowerCase().match(/[а-яёіңғүұқөһa-z]{3,}/g) || []);
  }
  const currentWords = wordSet(currentPurpose);
  function similarity(otherText) {
    if (!currentWords.size) return 0;
    const otherWords = wordSet(otherText);
    if (!otherWords.size) return 0;
    let common = 0;
    for (const w of currentWords) if (otherWords.has(w)) common++;
    return common / currentWords.size;
  }

  // Собираем всех кандидатов, у кого есть хотя бы статья ДДС, и сортируем
  // по похожести назначения платежа (лучшее совпадение — первым). При
  // равной похожести побеждает наличие ещё и заполненного вида операции.
  const candidates = [];
  for (const candidate of docs) {
    const candidateRow = candidate.РасшифровкаПлатежа && candidate.РасшифровкаПлатежа[0];
    const candidateCategoryKey = candidateRow && candidateRow.СтатьяДвиженияДенежныхСредств_Key;
    if (!candidateCategoryKey) continue;
    candidates.push({
      doc: candidate,
      row: candidateRow,
      categoryKey: candidateCategoryKey,
      score: similarity(candidate.НазначениеПлатежа || candidate.Комментарий) + (candidate.ВидОперации ? 0.01 : 0),
    });
  }
  candidates.sort((a, b) => b.score - a.score);
  const best = candidates[0];
  if (!best) return null;
  const { doc, row, categoryKey } = best;
  if (!doc || !categoryKey) return null;

  const debtType = (row && (row.ВидЗадолженности || row.ВидЗадолженности_Key)) || '';

  // Статья хранится как GUID — подтягиваем её человекочитаемое название
  const categoryCatalogs = ['Catalog_СтатьиДвиженияДенежныхСредств', 'Catalog_СтатьиДДС'];
  for (const catalog of categoryCatalogs) {
    const catUrl = `${base}/${catalog}(guid'${categoryKey}')?$format=json&$select=Description`;
    const catResp = await fetch(catUrl, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (catResp.ok) {
      const catData = await catResp.json().catch(() => null);
      if (catData && catData.Description) {
        return { categoryKey, categoryName: catData.Description, operationKind: doc.ВидОперации || '', debtType };
      }
    }
  }
  return { categoryKey, categoryName: '', operationKind: doc.ВидОперации || '' };
}

module.exports = { findCategoryKeyByName, findAccountKeyByCode, findHistoricalCategory };
