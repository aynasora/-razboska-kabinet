// Движок правил категоризации: типовые правила по регламенту + применение
// пользовательских правил к операции. Перенесено из server.js без изменений
// логики — только переезд в отдельный модуль.
const { normalizeScript, baseDefault } = require('./shared');

const DEFAULT_RULES = [
  { field: 'purpose', contains: 'аренд', category: 'Аренда', account: '3360/1710' },
  { field: 'purpose', contains: 'зарплат', category: 'Заработная плата', account: '3350' },
  { field: 'purpose', contains: 'налог', category: 'Налоги', account: '3130' },
  { field: 'purpose', contains: 'кпн', category: 'Налоги', account: '3130' },
  { field: 'purpose', contains: 'ипн', category: 'Налоги', account: '3120' },
  { field: 'purpose', contains: 'осмс', category: 'Налоги', account: '3150' },
  { field: 'purpose', contains: 'опвр', category: 'Налоги', account: '3150' },
  { field: 'purpose', contains: 'комисси', category: 'Банковские услуги', account: '3310/1710' },
  { field: 'purpose', contains: 'возврат', category: 'Возврат денежных средств', account: '3310/1710' },
];

// Правило может проверять назначение платежа, имя контрагента ИЛИ КНП
// (код назначения платежа — стандартный классификатор, который банки РК
// присылают прямо в выписке). normalizeRuleField приводит любое сырое
// значение поля (из формы, bulk-импорта, RULES_TEXT) к одному из трёх
// допустимых значений.
function normalizeRuleField(raw) {
  if (raw === 'counterparty') return 'counterparty';
  if (raw === 'knp') return 'knp';
  return 'purpose';
}
function ruleFieldValue(op, field) {
  if (field === 'counterparty') return op.counterparty;
  if (field === 'knp') return op.knp;
  return op.purpose;
}

// Применяет ваши правила категоризации к операции: смотрит назначение
// платежа, имя контрагента и/или КНП, и если находит совпадение —
// возвращает категорию и счёт. Правила проверяются по порядку, первое
// совпадение побеждает. Если ни одно правило не подошло — возвращает
// базовую логику «покупатель/поставщик» по регламенту, а не пустоту:
// минимум статья по направлению платежа должна быть определена всегда.
function applyRules(op, rules) {
  for (const rule of rules) {
    const raw = ruleFieldValue(op, rule.field) || '';
    const haystack = normalizeScript(raw).toLowerCase();
    const needle = normalizeScript(String(rule.contains || '')).toLowerCase();
    if (needle && haystack.includes(needle)) {
      return {
        category: rule.category || '',
        account: rule.account || '',
        operationKind: rule.operationKind || '',
        debtType: rule.debtType || '',
        source: 'rule',
      };
    }
  }
  const base = baseDefault(op.amount);
  return { ...base, operationKind: '', debtType: '', source: 'default' };
}

module.exports = { DEFAULT_RULES, normalizeRuleField, ruleFieldValue, applyRules };
