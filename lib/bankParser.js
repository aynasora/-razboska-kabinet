// Парсер банковской выписки (.xlsx) → единый объект операции, независимый
// от формата конкретного банка. Дальнейшая бухгалтерская логика
// (classifyOperation, resolveCounterparty, resolveContract, ...) читает
// только это единое поле, а не сырые колонки банка.
//
// Единый ("канонический") объект операции:
//   { date, documentNumber, direction, counterpartyName, counterpartyId,
//     counterpartyIban, knp, purpose, amount }
// где direction — 'in' (поступление) | 'out' (списание), amount — БЕЗ знака.
//
// Для совместимости с уже работающей частью приложения (findCounterpartyByBin,
// createDraftInOnec, makeFingerprint, весь клиентский интерфейс и т.д.,
// которые исторически ожидают поля counterparty/bin и amount СО ЗНАКОМ),
// к каждой операции ДОПОЛНИТЕЛЬНО добавляются legacy-алиасы:
//   counterparty === counterpartyName
//   bin           === counterpartyId
//   amount        — знак сохранён (поступление положительное, списание отрицательное)
// Ничего из существующей логики переписывать не нужно — она продолжает
// читать op.counterparty/op.bin/op.amount как раньше.
const XLSX = require('xlsx');
const crypto = require('crypto');
const { fixKazakhMojibake } = require('./shared');

const norm = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();

function parseAmount(v) {
  if (v === '' || v === undefined || v === null) return 0;
  const s = String(v).replace(/\s/g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// Ищет строку-шапку таблицы: банки (БЦК, Kaspi Business и похожие) часто
// кладут перед таблицей несколько строк с реквизитами, поэтому шапку нужно
// искать, а не считать первой строкой. Строка-шапка — та, где одновременно
// есть что-то похожее на "дата" и на дебет/кредит/сумму.
function findHeaderRow(grid) {
  for (let r = 0; r < Math.min(grid.length, 40); r++) {
    const cells = (grid[r] || []).map(norm);
    const findCol = (test) => cells.findIndex(test);
    const dateCol = findCol(c => c.includes('дата') && !c.includes('валют'));
    const debitCol = findCol(c => c.includes('дебет'));
    const creditCol = findCol(c => c.includes('кредит') && !c.includes('корреспондент'));
    const amountCol = findCol(c => c.includes('сумма') && !c.includes('конверт'));
    if (dateCol === -1 || (debitCol === -1 && creditCol === -1 && amountCol === -1)) continue;

    return {
      rowIndex: r,
      cols: {
        date: dateCol,
        debit: debitCol,
        credit: creditCol,
        amount: amountCol,
        documentNumber: findCol(c => c.includes('номер') && c.includes('док')),
        purpose: findCol(c => c.includes('мақсат') || c.includes('назначен') || c.includes('комментарий')),
        knp: findCol(c => c.includes('кнп') || c.includes('тмк')),
        // Общее поле "Корреспондент" — рабочий вариант для банков без
        // отдельных колонок плательщика/получателя (запасной вариант ниже).
        counterparty: findCol(c => c.includes('корреспондент') && !c.includes('банк') && !c.includes('бик') && !c.includes('иик')),
        binGeneric: findCol(c => c.includes('бин') || c.includes('иин')),
        // Плательщик — тот, кто ОТПРАВЛЯЕТ деньги (в БЦК — "Плательщик...",
        // в некоторых других выписках — "Отправитель..."; оба варианта
        // означают одно и то же поле с точки зрения парсера).
        payerName: findCol(c => (c.includes('плательщик') || c.includes('отправ')) && (c.includes('наимен') || c.includes('название'))),
        payerBin: findCol(c => (c.includes('бин') || c.includes('иин')) && (c.includes('плательщик') || c.includes('отправ'))),
        payerIban: findCol(c => (c.includes('плательщик') || c.includes('отправ')) && c.includes('иик')),
        // Получатель — тот, кому деньги ПРИХОДЯТ.
        receiverName: findCol(c => c.includes('получ') && (c.includes('наимен') || c.includes('название'))),
        receiverBin: findCol(c => (c.includes('бин') || c.includes('иин')) && c.includes('получ')),
        receiverIban: findCol(c => c.includes('получ') && c.includes('иик')),
      },
    };
  }
  return null;
}

function cell(row, colIndex) {
  return colIndex !== -1 ? row[colIndex] : '';
}

// Разбирает выписку в "табличном" формате с найденной шапкой — основной
// путь для БЦК/Kaspi Business и похожих банков.
function parseHeaderedGrid(grid, header, filename) {
  const { cols } = header;
  const operations = [];

  for (let r = header.rowIndex + 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row.every(c => String(c).trim() === '')) continue;

    const debit = parseAmount(cell(row, cols.debit));
    const credit = parseAmount(cell(row, cols.credit));
    const single = parseAmount(cell(row, cols.amount));
    const signedAmount = credit || (debit ? -debit : single);
    const dateRaw = String(cell(row, cols.date) || '');
    if (!dateRaw && !signedAmount) continue;

    const direction = signedAmount >= 0 ? 'in' : 'out';

    // Если пришли деньги (direction=in) — контрагент это ПЛАТЕЛЬЩИК (тот,
    // кто нам заплатил). Если ушли (direction=out) — контрагент это
    // ПОЛУЧАТЕЛЬ. Предпочитаем отдельные колонки плательщика/получателя,
    // если банк их даёт (точнее), иначе — общая колонка "Корреспондент" и
    // общая колонка БИН/ИИН (как было раньше).
    const nameCol = direction === 'in' ? cols.payerName : cols.receiverName;
    const binCol = direction === 'in' ? cols.payerBin : cols.receiverBin;
    const ibanCol = direction === 'in' ? cols.payerIban : cols.receiverIban;

    const counterpartyName = fixKazakhMojibake(
      nameCol !== -1 ? cell(row, nameCol) : cell(row, cols.counterparty)
    );
    const counterpartyId = String(binCol !== -1 ? cell(row, binCol) : cell(row, cols.binGeneric) || '').trim();
    const counterpartyIban = String(cell(row, ibanCol) || '').trim();

    operations.push(toOperation({
      date: dateRaw.split(' ')[0] || dateRaw,
      documentNumber: String(cell(row, cols.documentNumber) || '').trim(),
      direction,
      counterpartyName,
      counterpartyId,
      counterpartyIban,
      knp: String(cell(row, cols.knp) || ''),
      purpose: fixKazakhMojibake(cell(row, cols.purpose)),
      amount: Math.abs(signedAmount),
    }, filename, r));
  }
  return operations;
}

// Запасной вариант: простые выписки без "шапки банка", где заголовки — в
// самой первой строке (объектный разбор через sheet_to_json без header:1).
function parseFallbackSheet(sheet, filename) {
  const objRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  const pick = (row, names) => {
    for (const n of names) {
      const key = Object.keys(row).find(k => k.trim().toLowerCase() === n);
      if (key) return row[key];
    }
    return '';
  };

  return objRows.map((row, i) => {
    const dateRaw = pick(row, ['дата', 'дата операции', 'дата документа']);
    const amountIn = parseFloat(pick(row, ['приход', 'сумма прихода']) || 0) || 0;
    const amountOut = parseFloat(pick(row, ['расход', 'сумма расхода']) || 0) || 0;
    const amountSingle = parseFloat(pick(row, ['сумма']) || 0) || 0;
    const signedAmount = amountIn || (amountOut ? -amountOut : amountSingle);
    const direction = signedAmount >= 0 ? 'in' : 'out';

    const counterpartyName = fixKazakhMojibake(pick(row, [
      'контрагент', 'наименование контрагента', 'плательщик/получатель',
      'плательщикнаименование', 'получательнаименование',
    ]));
    const counterpartyId = String(pick(row, [
      'бин', 'иин', 'бин/иин', 'плательщикбин_иин', 'получательбин_иин',
    ]) || '');
    const counterpartyIban = String(pick(row, [
      'иик', 'плательщикиик', 'получательиик',
    ]) || '');

    return toOperation({
      date: String(dateRaw || ''),
      documentNumber: String(pick(row, ['номер документа', 'номер платежного документа', '№ документа']) || ''),
      direction,
      counterpartyName,
      counterpartyId,
      counterpartyIban,
      knp: String(pick(row, ['кнп', 'код назначения платежа', 'тмк']) || ''),
      purpose: fixKazakhMojibake(pick(row, ['назначение платежа', 'назначение', 'комментарий'])),
      amount: Math.abs(signedAmount),
    }, filename, i);
  }).filter(op => op.date || op.amount);
}

// Превращает канонический объект в то, что реально хранится/используется
// приложением дальше: канонические поля + legacy-алиасы (counterparty, bin,
// amount со знаком) + служебные поля списка операций.
function toOperation(canonical, filename, rowIndex) {
  const signedAmount = canonical.direction === 'out' ? -Math.abs(canonical.amount) : Math.abs(canonical.amount);
  return {
    id: crypto.randomUUID(),
    // --- единый ("канонический") набор полей ---
    date: canonical.date,
    documentNumber: canonical.documentNumber || '',
    direction: canonical.direction,
    counterpartyName: canonical.counterpartyName || '',
    counterpartyId: canonical.counterpartyId || '',
    counterpartyIban: canonical.counterpartyIban || '',
    knp: canonical.knp || '',
    purpose: canonical.purpose || '',
    // --- legacy-алиасы для остального (уже работающего) кода ---
    counterparty: canonical.counterpartyName || '',
    bin: canonical.counterpartyId || '',
    amount: signedAmount,
    // --- служебные поля ---
    suggestedCategory: '',
    status: 'review',
    sourceFile: filename,
    rowIndex,
  };
}

// Главная точка входа модуля. buffer — содержимое загруженного .xlsx,
// filename — исходное имя файла (попадает в op.sourceFile и в историю).
// Возвращает { operations, format } — format только для диагностики
// (какой путь разбора сработал), на бизнес-логику не влияет.
function parseStatement(buffer, filename) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });

  const header = findHeaderRow(grid);
  if (header) {
    return { operations: parseHeaderedGrid(grid, header, filename), format: 'headered-grid' };
  }
  return { operations: parseFallbackSheet(sheet, filename), format: 'fallback-object-rows' };
}

module.exports = { parseStatement };
