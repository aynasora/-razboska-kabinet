// Отправляет уже собранный payload (см. buildPayload) в 1С через OData —
// с самолечащимся повтором при отклонении необязательных полей. Перенесено
// из server.js (вторая половина бывшей createDraftInOnec) без изменений
// логики.

// Точные названия некоторых полей (Подотчетник_Key, СчетУчета_Key и т.п.)
// мы не можем угадать со 100% гарантией без доступа к вашей конкретной базе
// 1С — они могут называться иначе в вашей конфигурации. Поэтому вместо
// того чтобы один раз попробовать и в случае ошибки просто провалить
// создание черновика целиком — если 1С отвечает ошибкой, ссылающейся на
// одно из НЕОБЯЗАТЕЛЬНЫХ полей ниже, мы убираем именно это поле (и из
// шапки документа, и из строки табличной части) и пробуем снова. Так
// черновик создаётся почти всегда, а какие поля не прижились — видно в
// истории действий, чтобы можно было подсказать нам точное название поля
// под вашу конфигурацию.
const OPTIONAL_FIELDS = [
  'ВидОперации',
  'Договор_Key',
  'СтатьяДвиженияДенежныхСредств_Key',
  'ВидЗадолженности',
  'Подотчетник_Key',
  'СчетУчета_Key',
  'РасшифровкаПлатежа',
];

function fieldIsPresent(payloadObj, fieldName) {
  if (fieldName in payloadObj) return true;
  return Array.isArray(payloadObj.РасшифровкаПлатежа) && payloadObj.РасшифровкаПлатежа.some(row => fieldName in row);
}
function stripOptionalField(payloadObj, fieldName) {
  if (fieldName === 'РасшифровкаПлатежа') {
    const { РасшифровкаПлатежа, ...rest } = payloadObj;
    return rest;
  }
  const clone = { ...payloadObj };
  delete clone[fieldName];
  if (Array.isArray(clone.РасшифровкаПлатежа)) {
    clone.РасшифровкаПлатежа = clone.РасшифровкаПлатежа.map(row => {
      if (!(fieldName in row)) return row;
      const rowCopy = { ...row };
      delete rowCopy[fieldName];
      return rowCopy;
    });
  }
  return clone;
}

// endpoint — полный URL коллекции документа (например
// https://.../odata/standard.odata/Document_ПлатежноеПоручениеИсходящее),
// как его вернул buildPayload. opts.missingNoContract — если true, сразу
// добавляем предупреждение об отсутствующем договоре в droppedFields (это
// поле и так не попало в payload, просто честно сообщаем об этом).
async function writeTo1C(endpoint, payload, settings, opts = {}) {
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  async function postDocument(body) {
    return fetch(`${endpoint}?$format=json`, {
      method: 'POST',
      headers: {
        Authorization: `Basic ${auth}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      body: JSON.stringify(body),
    });
  }

  let currentPayload = payload;
  let response = await postDocument(currentPayload);
  const droppedFields = [];
  if (opts.missingNoContract) {
    droppedFields.push('Договор_Key (нет договора и нет «Без договора» — создайте кнопкой «Без договора» или в 1С)');
  }
  let safetyCounter = 0;
  while (!response.ok && safetyCounter < OPTIONAL_FIELDS.length) {
    safetyCounter++;
    const text = await response.text().catch(() => '');
    const badField = OPTIONAL_FIELDS.find(f => fieldIsPresent(currentPayload, f) && text.includes(f));
    if (!badField) {
      throw new Error(`1С ответила ${response.status}: ${text.slice(0, 300)}`);
    }
    droppedFields.push(badField);
    currentPayload = stripOptionalField(currentPayload, badField);
    response = await postDocument(currentPayload);
  }
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`1С ответила ${response.status}: ${text.slice(0, 300)}`);
  }

  const data = await response.json().catch(() => ({}));
  return { docNumber: data.Number || null, docRef: data.Ref_Key || null, droppedFields };
}

module.exports = { writeTo1C };
