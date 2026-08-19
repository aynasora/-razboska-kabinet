// НОВАЯ возможность (раньше отсутствовала): после того как writeTo1C
// отчитался об успехе, лишний раз перезапрашиваем 1С и убеждаемся, что
// документ действительно там — OData иногда возвращает 200 OK на запрос,
// который на стороне 1С по факту не сохранился (например, из-за таймаута
// синхронной записи). Это НЕ блокирующая проверка: если она сама не
// удалась технически, считаем документ созданным (writeTo1C уже получил
// от 1С успешный ответ) и просто помечаем verified:false.
async function verifyCreatedDocument(op, settings, created) {
  if (!created || (!created.docRef && !created.docNumber)) {
    return { verified: false, reason: 'нет ни Ref_Key, ни номера созданного документа для проверки' };
  }
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const docType = op.amount >= 0 ? 'Document_ПлатежноеПоручениеВходящее' : 'Document_ПлатежноеПоручениеИсходящее';

  try {
    if (created.docRef) {
      const url = `${base}/${docType}(guid'${created.docRef}')?$format=json&$select=Ref_Key,Number`;
      const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
      if (response.ok) {
        const data = await response.json().catch(() => null);
        if (data && data.Ref_Key) return { verified: true, ref: data.Ref_Key, number: data.Number || created.docNumber };
      }
    }
    if (created.docNumber) {
      const filter = encodeURIComponent(`Number eq '${String(created.docNumber).replace(/'/g, "''")}'`);
      const url = `${base}/${docType}?$format=json&$filter=${filter}&$select=Ref_Key,Number&$top=1`;
      const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
      if (response.ok) {
        const data = await response.json().catch(() => null);
        const found = data && data.value && data.value[0];
        if (found) return { verified: true, ref: found.Ref_Key, number: found.Number };
      }
    }
    return { verified: false, reason: 'документ с таким Ref_Key/номером не нашёлся при повторном запросе' };
  } catch (e) {
    return { verified: false, reason: 'проверка технически не удалась: ' + e.message };
  }
}

module.exports = { verifyCreatedDocument };
