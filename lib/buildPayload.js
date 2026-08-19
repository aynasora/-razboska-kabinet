// Собирает OData-payload платёжного поручения (шапка + строка табличной
// части "Расшифровка платежа"), но НЕ отправляет его в 1С — отправка
// вынесена в writeTo1C. Раньше сборка payload и его отправка были одной
// функцией (createDraftInOnec в server.js); разделены, чтобы можно было
// собрать payload и посмотреть на него без реальной отправки (см. dryRun).
// Логика заполнения полей не изменена — это чистое перемещение.
const { toIsoDate, guessOperationKindLiteral } = require('./shared');
const { findCategoryKeyByName, findAccountKeyByCode } = require('./resolveCashFlowArticle');

async function buildPayload(op, settings, resolvedCategory, resolvedAccount, resolvedOperationKind, resolvedDebtType) {
  const base = settings.baseUrl.replace(/\/+$/, ''); // убираем лишний / на конце
  const docType = op.amount >= 0 ? 'Document_ПлатежноеПоручениеВходящее' : 'Document_ПлатежноеПоручениеИсходящее';
  const endpoint = `${base}/${docType}`;

  if (!op.counterpartyKey) {
    throw new Error('Контрагент не сопоставлен со справочником 1С — сначала выберите контрагента вручную');
  }

  const payload = {
    Date: toIsoDate(op.date),
    Posted: false, // черновик — не проводится автоматически
    Организация_Key: settings.orgKey || '',
    СчетОрганизации_Key: settings.accountKey || '',
    Контрагент_Key: op.counterpartyKey || '',
    СуммаДокумента: Math.abs(op.amount),
    НазначениеПлатежа: op.purpose,
    Комментарий: op.purpose, // должно совпадать с назначением платежа, как в вашей 1С
  };

  // Статья ДДС: если вы её НЕ правили руками и в истории 1С уже есть
  // готовый GUID — используем его напрямую. Иначе (ручная правка, правило,
  // или текст без готового GUID) ищем статью по названию в справочнике.
  let categoryKey = null;
  if (!op.manualCategory && op.historicalCategoryKey) {
    categoryKey = op.historicalCategoryKey;
  } else if (resolvedCategory) {
    categoryKey = await findCategoryKeyByName(resolvedCategory, settings);
    // Если не нашли даже по имени — не страшно: текст статьи уже есть в
    // комментарии и в интерфейсе, вы сможете проставить её вручную в 1С.
  }
  if (categoryKey) payload.СтатьяДвиженияДенежныхСредств_Key = categoryKey;

  // "Вид операции" — строгий список в 1С. Приоритет:
  //   1. История (скопировано с уже существующего документа — 100% верно)
  //   2. Подтверждённое точное значение (проверено напрямую через вашу базу)
  //   3. Обоснованная попытка по тому же принципу именования — если 1С её
  //      отклонит, мы автоматически повторим запрос без этого поля (см. writeTo1C).
  let operationKindLiteral = null;
  if (op.historicalOperationKind) {
    operationKindLiteral = op.historicalOperationKind;
  } else if (resolvedOperationKind) {
    operationKindLiteral = guessOperationKindLiteral(resolvedOperationKind);
  } else if (op.amount < 0 && (resolvedCategory === 'Расчёты с поставщиками и подрядчиками' || resolvedCategory === 'Расчеты с поставщиками и подрядчиками')) {
    // Базовый сценарий "Оплата поставщику" (нет ни истории, ни более
    // специфичного правила вроде аренды/зарплаты/налогов) — ставим вид
    // операции по умолчанию. Если точное написание в вашей 1С отличается,
    // это поле безопасно "отвалится" через OPTIONAL_FIELDS в writeTo1C.
    operationKindLiteral = guessOperationKindLiteral('Оплата поставщику');
  }
  if (operationKindLiteral) {
    payload.ВидОперации = operationKindLiteral;
  }

  // Договор — привязываем, только если найден однозначно. Если у
  // контрагента несколько договоров и непонятно, какой из них — лучше
  // остановиться и попросить вас выбрать вручную, чем угадать неверно.
  // Если нет ни одного договора, ни "Без договора" — тоже не гадаем и не
  // создаём автоматически здесь: просто оставляем поле пустым и явно
  // сообщаем об этом через missingNoContract → droppedFields в writeTo1C.
  if (op.contractStatus === 'ambiguous') {
    throw new Error('У контрагента несколько договоров, и не удалось определить нужный по назначению платежа — выберите договор вручную перед подтверждением.');
  }
  const missingNoContract = op.contractStatus === 'need_create_no_contract' && !op.contractKey;
  if (op.contractKey) {
    payload.Договор_Key = op.contractKey;
  }

  // ВАЖНО: сумма, статья ДДС и договор в этом документе хранятся не
  // только на уровне самого документа, но и в отдельной табличной части
  // "Расшифровка платежа" (видна как таблица внутри документа в 1С).
  // Без нее табличная часть остаётся пустой строкой, даже если общая
  // сумма наверху документа заполнена правильно.
  //
  // ВАЖНО №2: если GUID-поле (Договор_Key, СтатьяДвиженияДенежныхСредств_Key)
  // неизвестно — его нужно просто НЕ включать в объект, а не отправлять
  // пустую строку. 1С понимает "поля нет" нормально, а вот пустую строку
  // вместо GUID отвергает ошибкой "Не удалось разобрать строку '' как
  // значение типа Edm.Guid".
  const amountAbs = Math.abs(op.amount);
  const lineItem = {
    LineNumber: 1,
    СуммаПлатежа: amountAbs,
    КурсВзаиморасчетов: 1,
    СуммаВзаиморасчетов: amountAbs,
  };
  if (op.contractKey) lineItem.Договор_Key = op.contractKey;
  if (categoryKey) lineItem.СтатьяДвиженияДенежныхСредств_Key = categoryKey;

  // "Вид задолженности" — раньше бралось ТОЛЬКО из истории 1С; если у
  // контрагента ещё нет истории, поле молча оставалось пустым, даже если
  // resolvedDebtType (из правила/ручной правки/подстановки по умолчанию)
  // был известен. Приоритет: история 1С (100% верно) > то, что определили
  // для этой операции.
  const debtTypeValue = op.historicalDebtType || resolvedDebtType || '';
  if (debtTypeValue) lineItem.ВидЗадолженности = debtTypeValue;

  // "Счёт учёта (БУ)" в табличной части — ищем GUID счёта по коду
  // (resolvedAccount, например "1251").
  const accountKey = await findAccountKeyByCode(resolvedAccount, settings);
  if (accountKey) lineItem.СчетУчета_Key = accountKey;

  // "Подотчётник" — отдельное поле табличной части именно для операции
  // "Перечисление денежных средств подотчётнику": получатель там обычно
  // указывается той же ссылкой, что и Контрагент_Key в шапке документа.
  // Добавляем это поле, только когда вид операции действительно про
  // подотчётные средства — на остальных документах такого поля нет.
  const isAccountablePersonPayment = /подотчет/i.test(
    String(resolvedOperationKind || '') + ' ' + String(op.historicalOperationKind || '')
  );
  if (isAccountablePersonPayment && op.counterpartyKey) {
    lineItem.Подотчетник_Key = op.counterpartyKey;
  }

  payload.РасшифровкаПлатежа = [lineItem];

  return { endpoint, payload, missingNoContract };
}

module.exports = { buildPayload };
