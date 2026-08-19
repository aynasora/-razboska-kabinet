// НОВАЯ возможность (раньше отсутствовала): собрать payload платёжного
// поручения точно так же, как перед реальной отправкой, но НЕ отправлять
// его в 1С. Полезно для отладки ("что именно мы туда пошлём") и для
// проверки конфигурации, не создавая лишних черновиков документов.
const { buildPayload } = require('./buildPayload');

async function dryRun(op, settings, resolvedCategory, resolvedAccount, resolvedOperationKind, resolvedDebtType) {
  const { endpoint, payload, missingNoContract } = await buildPayload(
    op, settings, resolvedCategory, resolvedAccount, resolvedOperationKind, resolvedDebtType
  );
  return {
    endpoint,
    payload,
    warnings: missingNoContract
      ? ['Нет договора и нет «Без договора» — при реальной отправке поле Договор_Key останется пустым']
      : [],
  };
}

module.exports = { dryRun };
