// Определяет итоговую статью ДДС / счёт / вид операции / вид задолженности
// для операции. Раньше этот приоритет (ручная правка > история 1С > правило
// > базовая логика) был продублирован дословно в двух местах server.js
// (GET /api/operations и confirmOperation) — здесь он объединён в одну
// функцию, которой пользуются оба места. Поведение не изменилось: те же
// четыре источника, тот же порядок приоритета.
const { applyRules } = require('./rules');

function classifyOperation(op, rules) {
  const ruleMatch = applyRules(op, rules);
  const category = op.manualCategory || op.historicalCategory || ruleMatch.category;
  const account = op.manualAccount || op.historicalAccount || ruleMatch.account;
  const operationKind = op.manualOperationKind || op.historicalOperationKind || ruleMatch.operationKind;
  const debtType = op.manualDebtType || op.historicalDebtType || ruleMatch.debtType;
  const isDefaultOnly = !op.manualCategory && !op.historicalCategory && ruleMatch.source === 'default';
  return {
    category,
    account,
    operationKind,
    debtType,
    isDefaultOnly, // категория определена только базовой логикой, стоит перепроверить глазами
    needsAttention: !category,
  };
}

module.exports = { classifyOperation };
