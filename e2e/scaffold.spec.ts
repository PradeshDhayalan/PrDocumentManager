import { test, expect } from '@playwright/test';
test('control uses light Blue without harness header and handles empty and unsaved parents', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/');
  await expect(page.getByRole('grid', { name: 'Documents', exact: true })).toBeVisible();
  await expect(page.getByRole('toolbar', { name: 'DMS Grid test harness' })).toHaveCount(0);
  await expect(page.getByRole('combobox', { name: 'Record', exact: true })).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Dark', exact: true })).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Teal', exact: true })).toHaveCount(0);
  const colors = await page
    .getByRole('region', { name: 'Documents', exact: true })
    .evaluate((element) => ({
      brand: getComputedStyle(element).getPropertyValue('--colorBrandBackground').trim(),
      background: getComputedStyle(element).backgroundColor,
    }));
  expect(colors.brand.toLowerCase()).toBe('#0f6cbd');
  expect(colors.background).toBe('rgb(255, 255, 255)');
  await page.goto('/?record=empty');
  await expect(page.getByText('No documents yet', { exact: true })).toBeVisible();
  await page.goto('/?record=unsaved');
  await expect(
    page.getByText('Save this record to manage documents.', { exact: true }),
  ).toBeVisible();
  await expect(page.getByRole('button', { name: 'Upload', exact: true })).toHaveCount(0);
  expect(errors).toEqual([]);
});
