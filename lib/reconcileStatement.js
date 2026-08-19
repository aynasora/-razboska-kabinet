// Сверяет свежезагруженные операции с 1С: контрагент по БИН/ИИН, история
// категоризации, договор — и заодно проверяет, не разнесена ли операция
// уже в самой 1С. Раньше это был инлайновый блок внутри
// app.post('/api/upload', ...) в server.js — перенесено без изменений
// логики в отдельную функцию.
const { checkExistingInOnec } = require('./duplicateGuard');
const { findCounterpartyByBin } = require('./resolveCounterparty');
const { resolveSupplierContract } = require('./resolveContract');
const { findHistoricalCategory } = require('./resolveCashFlowArticle');

// newOps — только что добавленные операции (без дублей), all — полный
// список операций в хранилище (уже включает newOps по тем же объектам —
// мутируем их поля прямо в этом массиве, как и раньше). settings —
// настройки подключения к 1С. Возвращает { total, alreadyIn1c,
// newCounterparty, connected }.
async function reconcileStatement(newOps, all, settings) {
  if (!settings.baseUrl || !settings.login) {
    return { total: newOps.length, alreadyIn1c: 0, newCounterparty: 0, connected: false };
  }

  for (const op of newOps) {
    const target = all.find(o => o.id === op.id);

    try {
      const alreadyExists = await checkExistingInOnec(op, settings);
      if (alreadyExists) {
        target.status = 'already_in_1c';
        continue; // не тратим время на сопоставление контрагента — операция и так пропускается
      }
    } catch (e) {
      // Если сверка не удалась — не блокируем, просто идём дальше как обычно.
    }

    try {
      const found = await findCounterpartyByBin(op.bin, settings, op.counterparty);
      if (found) {
        target.counterpartyKey = found.Ref_Key;
        target.counterpartyMatchedName = found.Description || '';

        // Смотрим, как этот контрагент категоризировался раньше в 1С —
        // если найдём паттерн, используем ту же статью ДДС автоматически.
        try {
          const historical = await findHistoricalCategory(found.Ref_Key, op.amount, settings, op.purpose);
          if (historical) {
            target.historicalCategory = historical.categoryName || '';
            target.historicalCategoryKey = historical.categoryKey || '';
            target.historicalOperationKind = historical.operationKind || '';
            target.historicalDebtType = historical.debtType || '';
          }
        } catch (e) {
          // Не критично — просто не будет автопредложения по истории для этой операции.
        }

        // Ищем договор контрагента — если он есть в справочнике и найден
        // однозначно, привяжем его к документу автоматически.
        try {
          const contract = await resolveSupplierContract(found.Ref_Key, op.purpose, settings);
          target.contractStatus = contract.status;
          if (contract.status === 'matched') {
            target.contractKey = contract.key;
            target.contractName = contract.name;
          } else if (contract.status === 'ambiguous') {
            target.contractOptions = contract.options;
          }
        } catch (e) {
          // Не критично — просто без привязки к договору.
        }
      } else {
        target.status = 'new_counterparty';
      }
    } catch (e) {
      if (e.ambiguousCounterparty) {
        target.status = 'ambiguous_counterparty';
        target.counterpartyOptions = e.ambiguousCounterparty.map(o => ({ key: o.Ref_Key, name: o.Description }));
      }
      // Иначе — проверка не удалась (например, неверный пароль) — не
      // блокируем загрузку, просто оставляем операцию как есть, на проверке.
    }
  }

  return {
    total: newOps.length,
    alreadyIn1c: newOps.filter(o => all.find(u => u.id === o.id)?.status === 'already_in_1c').length,
    newCounterparty: newOps.filter(o => all.find(u => u.id === o.id)?.status === 'new_counterparty').length,
    connected: true,
  };
}

module.exports = { reconcileStatement };
