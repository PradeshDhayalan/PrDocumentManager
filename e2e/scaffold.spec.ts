import { test, expect } from '@playwright/test';
test('virtual control shell mounts in every theme and handles an unsaved parent', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/');
  await expect(page.getByRole('region', { name: 'Documents', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Blue', exact: true }).focus();
  await page.keyboard.press('ArrowRight');
  await expect(page.getByRole('button', { name: 'Teal', exact: true })).toBeFocused();
  await page.keyboard.press('Enter');
  for (const theme of ['Blue', 'Teal', 'Dark', 'High contrast']) {
    await page.getByRole('button', { name: theme, exact: true }).click();
    await expect(page.getByText('No documents yet', { exact: true })).toBeVisible();
  }
  await page.getByRole('button', { name: 'Contoso Ltd', exact: true }).click();
  await expect(
    page.getByText('Save this record to manage documents.', { exact: true }),
  ).toBeVisible();
  expect(errors).toEqual([]);
});
