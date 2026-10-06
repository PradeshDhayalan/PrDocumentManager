import { test, expect } from '@playwright/test';
let createdColumn: string | undefined;
test.afterEach(async ({ request }) => {
  if (createdColumn) await request.delete(`/__mock/columns/${createdColumn}`);
  createdColumn = undefined;
});

test('pagination defaults to ten, replaces pages and resets on search', async ({ page }) => {
  await page.goto('/');
  const rows = page.getByRole('grid').locator('[data-document-id]');
  await expect(rows).toHaveCount(10);
  const first = await rows.first().getAttribute('data-document-id');
  await page.getByRole('button', { name: 'Next page', exact: true }).click();
  await expect(page.getByText(/Page 2 of/)).toBeVisible();
  expect(await rows.first().getAttribute('data-document-id')).not.toBe(first);
  await page.getByRole('button', { name: 'Previous page', exact: true }).click();
  await expect(rows.first()).toHaveAttribute('data-document-id', first!);
  await page
    .getByRole('textbox', { name: 'Search documents', exact: true })
    .fill('Master Services');
  await expect(rows).toHaveCount(1);
  await expect(page.getByRole('button', { name: 'Next page', exact: true })).toBeDisabled();
});

test('configuration creates a persistent column, edits it on double click and searches and filters it', async ({
  page,
  request,
}) => {
  const name = `dms_test${Date.now()}`;
  createdColumn = name;
  const title = `Project ${Date.now()}`;
  await page.goto('/');
  await page.getByRole('button', { name: 'PCF Configuration', exact: true }).click();
  const config = page.getByRole('dialog', { name: 'PCF Configuration', exact: true });
  await config.getByRole('spinbutton', { name: 'Documents per page' }).fill('5');
  await config.getByRole('textbox', { name: 'Column label', exact: true }).fill(title);
  await config.getByRole('textbox', { name: 'Logical name', exact: true }).fill(name);
  await config.getByRole('button', { name: 'Create column', exact: true }).click();
  await config.getByRole('checkbox', { name: `Filter ${title}`, exact: true }).check();
  await config.getByRole('combobox', { name: 'Search by field', exact: true }).click();
  await page.getByRole('option', { name: title, exact: true }).click();
  await config.getByRole('button', { name: 'Apply configuration', exact: true }).click();
  const rows = page.getByRole('grid').locator('[data-document-id]');
  await expect(rows).toHaveCount(5);
  await expect(page.getByRole('columnheader', { name: title, exact: false })).toBeVisible();
  const id = await rows.first().getAttribute('data-document-id');
  await rows.first().dblclick();
  const edit = page.getByRole('dialog', { name: 'Edit details', exact: true });
  await edit.getByRole('textbox', { name: title, exact: true }).fill('Project-alpha');
  await edit.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(edit).not.toBeVisible();
  await page.reload();
  await page.getByRole('textbox', { name: 'Search documents', exact: true }).fill('Project-alpha');
  await expect(rows).toHaveCount(1);
  await expect(rows.first()).toHaveAttribute('data-document-id', id!);
  await page.getByRole('button', { name: 'Filters', exact: true }).click();
  await page.getByRole('textbox', { name: `Filter ${title}`, exact: true }).fill('no match');
  await expect(rows).toHaveCount(0);
  const metadata = await request.get(
    "/api/data/v9.2/EntityDefinitions(LogicalName='dms_document')?$expand=Attributes",
  );
  expect(
    (await metadata.json()).Attributes.some((a: { LogicalName: string }) => a.LogicalName === name),
  ).toBe(true);
});

test('tile sizes apply, tiles have no pencil and file drag shows an overlay until drop', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'PCF Configuration', exact: true }).click();
  const config = page.getByRole('dialog', { name: 'PCF Configuration', exact: true });
  await config.getByRole('combobox', { name: 'Tile size', exact: true }).click();
  await page.getByRole('option', { name: 'large', exact: true }).click();
  await config.getByRole('button', { name: 'Apply configuration', exact: true }).click();
  await page.getByRole('button', { name: 'Switch view', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Tiles view', exact: true }).click();
  const tile = page
    .getByRole('listbox', { name: 'Documents', exact: true })
    .getByRole('option')
    .first();
  await expect(tile.locator('button')).toHaveCount(0);
  expect(
    await tile
      .locator(':scope > div')
      .first()
      .evaluate((element) => element.getBoundingClientRect().height),
  ).toBe(260);
  await tile.dblclick();
  await expect(page.getByRole('dialog', { name: 'Edit details', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  const transfer = await page.evaluateHandle(() => {
    const data = new DataTransfer();
    data.items.add(new File(['Drop test'], 'drop-test.txt', { type: 'text/plain' }));
    return data;
  });
  const region = page.getByRole('region', { name: 'Documents', exact: true });
  await region.dispatchEvent('dragenter', { dataTransfer: transfer });
  await expect(page.getByText('Drop documents to upload', { exact: true })).toBeVisible();
  await region.dispatchEvent('drop', { dataTransfer: transfer });
  await expect(page.getByText('Drop documents to upload', { exact: true })).toHaveCount(0);
  await expect(page.getByRole('dialog', { name: 'Upload', exact: true })).toBeVisible();
});
