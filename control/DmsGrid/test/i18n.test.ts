import { createI18n } from '../src/services/i18n';
test('localised placeholders preserve zero and replace all occurrences', () => {
  const t = createI18n(() => '{count} of {count} items');
  expect(t('items', { count: 0 })).toBe('0 of 0 items');
});
