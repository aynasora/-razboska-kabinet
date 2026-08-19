// Защита от повторного создания одного и того же документа: "отпечаток"
// операции + сверка с уже существующими документами в 1С + сверка со
// списком уже загруженных операций. Перенесено из server.js без изменений
// логики (checkExistingInOnec, makeFingerprint) плюс одна новая тонкая
// обёртка isDuplicateAmongOperations, которая раньше была инлайновой
// проверкой внутри confirmOperation в server.js.
const { normalizeCounterpartyName } = require('./shared');

// "Отпечаток" операции — используется, чтобы не создать в 1С второй
// одинаковый документ, если одна и та же выписка (или пересекающийся
// период) была загружена дважды и попала в список операций как два разных
// объекта (с разными id). Берём только день из даты (без времени), сумму
// без знака (направление уже видно по документу), БИН/ИИН только цифрами
// и "ядро" названия контрагента — так же, как две почти одинаковые строки
// из разных выгрузок банка распознаются как одна и та же операция.
function makeFingerprint(op) {
  const day = String(op.date || '').trim().split(' ')[0];
  const amount = Math.abs(Number(op.amount) || 0).toFixed(2);
  const binDigits = String(op.bin || '').replace(/\D/g, '');
  const normName = normalizeCounterpartyName(op.counterparty);
  const purpose = String(op.purpose || '').toLowerCase().replace(/\s+/g, ' ').trim();
  return `${day}|${amount}|${binDigits}|${normName}|${purpose}`;
}

// Проверяет, нет ли уже среди документов 1С (Платёжное поручение
// входящее/исходящее) операции с такой же датой и суммой — чтобы не
// предлагать создавать то, что уже разнесено (вручную или иначе).
async function checkExistingInOnec(op, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const docType = op.amount >= 0 ? 'Document_ПлатежноеПоручениеВходящее' : 'Document_ПлатежноеПоручениеИсходящее';
  const amount = Math.abs(op.amount);
  // Дата у нас хранится как ДД.ММ.ГГГГ — переводим в ГГГГ-ММ-ДД для фильтра OData
  const parts = String(op.date).split('.');
  if (parts.length !== 3) return false;
  const isoDate = `${parts[2]}-${parts[1]}-${parts[0]}`;
  const filter = encodeURIComponent(
    `Date ge datetime'${isoDate}T00:00:00' and Date le datetime'${isoDate}T23:59:59' and СуммаДокумента eq ${amount}`
  );
  const url = `${base}/${docType}?$format=json&$filter=${filter}&$select=Ref_Key`;

  const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
  if (!response.ok) return false; // если сверить не удалось — не блокируем, просто считаем, что не нашли
  const data = await response.json().catch(() => null);
  if (!data) return false;
  return (data.value || []).length > 0;
}

// Проверка дубликатов НА УРОВНЕ НАШЕГО СПИСКА операций (в дополнение к
// checkExistingInOnec, которая смотрит в саму 1С): если среди уже
// загруженных операций есть другая, с таким же отпечатком, и по ней
// документ УЖЕ был создан — значит, эта операция, скорее всего, попала в
// список второй раз (например, из пересекающегося периода двух выгрузок).
// Возвращает найденную операцию-оригинал или null.
function findDuplicateAmongOperations(op, all) {
  const fingerprint = makeFingerprint(op);
  return all.find(o => o.id !== op.id && o.documentCreated && makeFingerprint(o) === fingerprint) || null;
}

module.exports = { makeFingerprint, checkExistingInOnec, findDuplicateAmongOperations };
