import { test, expect, Page } from '@playwright/test';
test.describe.configure({ mode: 'serial' });
async function start(page: Page) {
  await page.goto('/');
  await expect(page.locator('[data-document-id]')).toHaveCount(22);
}
async function tiles(page: Page) {
  await page.getByRole('button', { name: 'Views', exact: true }).first().click();
  await page.getByRole('menuitem', { name: 'Tiles', exact: true }).click();
}
async function record(page: Page, name: string) {
  await page.getByRole('combobox', { name: 'Record', exact: true }).click();
  await page.getByRole('option', { name, exact: true }).click();
}
test.beforeEach(async ({ request }) => {
  await request.post('/__mock/reset');
});
for (const view of ['List', 'Tiles']) {
  test(`${view}: shared selection, ranges, context menu and thumbnails`, async ({ page }) => {
    await start(page);
    if (view === 'Tiles') await tiles(page);
    const rows = page.locator('[data-document-id]'),
      selected = page.locator('[data-document-id][aria-selected=true]');
    await rows.nth(0).click();
    await expect(selected).toHaveCount(1);
    await rows.nth(2).click({ modifiers: ['Shift'] });
    await expect(selected).toHaveCount(3);
    await rows.nth(4).click({ modifiers: ['Control'] });
    await expect(selected).toHaveCount(4);
    await rows.nth(2).click({ button: 'right' });
    await expect(
      page.getByRole('menuitem', { name: 'Edit properties', exact: true }),
    ).toBeVisible();
    await expect(selected).toHaveCount(4);
    await page.keyboard.press('Escape');
    await rows.nth(6).click({ button: 'right' });
    await expect(selected).toHaveCount(1);
    await page.keyboard.press('Escape');
    await rows.nth(0).focus();
    await page.keyboard.press('Control+a');
    await expect(selected).toHaveCount(22);
    await page.keyboard.press('Escape');
    await expect(selected).toHaveCount(0);
    const image = page.getByAltText('Thumbnail preview of Site Photo.jpg');
    await expect(image).toHaveAttribute('src', /^blob:/);
    if (view === 'Tiles') {
      expect(
        await rows.first().evaluate((element) => getComputedStyle(element).boxShadow),
      ).not.toBe('none');
    }
  });
}
test('themes, search, large-record paging and unsaved/empty parents', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await start(page);
  for (const theme of ['Blue', 'Teal', 'Dark', 'High contrast']) {
    await page.getByRole('button', { name: theme, exact: true }).click();
    await expect(page.locator('[data-document-id]')).toHaveCount(22);
  }
  await page.getByPlaceholder('Search documents').fill('Insurance');
  await expect(page.locator('[data-document-id]')).toHaveCount(1);
  await page.getByPlaceholder('Search documents').fill('');
  await expect(page.locator('[data-document-id]')).toHaveCount(22);
  await record(page, 'Fabrikam · 1,200 documents');
  await expect(page.locator('[data-document-id]')).toHaveCount(50);
  await expect(page.getByText('50 of 1200 items', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Load more', exact: true }).click();
  await expect(page.locator('[data-document-id]')).toHaveCount(100);
  await record(page, 'Empty contact');
  await expect(page.getByText('No documents yet', { exact: true })).toBeVisible();
  await record(page, 'Unsaved record');
  await expect(
    page.getByText('Save this record to manage documents.', { exact: true }),
  ).toBeVisible();
  expect(errors).toEqual([]);
});
test('upload, preview, download, metadata edit and deletion persist through API', async ({
  page,
}) => {
  await start(page);
  const name = 'Review upload.txt',
    bytes = Buffer.from('Document review test\n');
  await page
    .locator('input[type=file]')
    .setInputFiles({ name, mimeType: 'text/plain', buffer: bytes });
  const row = page.locator('[data-document-id]').filter({ hasText: name });
  await expect(row).toHaveCount(1);
  await expect(page.getByText('Uploaded', { exact: true })).toBeVisible();
  await page
    .locator('input[type=file]')
    .setInputFiles({ name, mimeType: 'text/plain', buffer: bytes });
  await expect(page.getByText('A matching file already exists', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Skip', exact: true }).click();
  await row.dblclick();
  await expect(page.getByRole('dialog').locator('pre')).toHaveText('Document review test');
  await page.getByRole('button', { name: 'Close', exact: true }).click();
  await row.click();
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download', exact: true }).first().click();
  const file = await downloadPromise;
  expect(file.suggestedFilename()).toBe(name);
  const stream = await file.createReadStream();
  const chunks: Buffer[] = [];
  for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
  expect(Buffer.concat(chunks)).toEqual(bytes);
  await page.getByRole('button', { name: 'Edit properties', exact: true }).first().click();
  await page.getByRole('textbox', { name: 'Description', exact: true }).fill('Reviewed metadata');
  await page.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.getByText('Properties saved.', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Delete', exact: true }).first().click();
  await page.getByRole('dialog').getByRole('button', { name: 'Delete', exact: true }).click();
  await expect(row).toHaveCount(0);
  await page.getByRole('button', { name: 'Refresh', exact: true }).click();
  await expect(page.locator('[data-document-id]')).toHaveCount(22);
});
test('SharePoint capabilities, URL validation and reference creation', async ({ page }) => {
  await start(page);
  await page.getByRole('combobox', { name: 'Active storage', exact: true }).click();
  await page.getByRole('option', { name: 'SharePoint', exact: true }).click();
  await expect(page.getByRole('button', { name: 'New', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Upload', exact: true })).toHaveCount(0);
  await page.getByRole('button', { name: 'New', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Add SharePoint link', exact: true }).click();
  await page
    .getByRole('textbox', { name: 'SharePoint URL', exact: true })
    .fill('https://evil.test/doc.pdf');
  await page.getByRole('button', { name: 'Add SharePoint link', exact: true }).click();
  await expect(
    page.getByText('Enter an HTTPS URL on an allowed SharePoint host.', { exact: true }),
  ).toBeVisible();
  await page
    .getByRole('textbox', { name: 'SharePoint URL', exact: true })
    .fill('https://contoso.sharepoint.com/sites/review/sample.pdf');
  await page
    .getByRole('textbox', { name: 'Display name (optional)', exact: true })
    .fill('Review reference');
  await page.getByRole('button', { name: 'Add SharePoint link', exact: true }).click();
  await expect(
    page.locator('[data-document-id]').filter({ hasText: 'Review reference' }),
  ).toHaveCount(1);
  await page.getByRole('combobox', { name: 'Active storage', exact: true }).click();
  await page.getByRole('option', { name: 'Dataverse Notes', exact: true }).click();
});
test('narrow command bar keeps commands in Fluent overflow', async ({ page }) => {
  await start(page);
  await page.setViewportSize({ width: 375, height: 812 });
  await page.getByRole('button', { name: 'More commands', exact: true }).click();
  await expect(page.getByRole('menuitem', { name: 'Edit columns', exact: true })).toBeVisible();
  await expect(page.getByRole('menuitem', { name: 'Search documents', exact: true })).toBeVisible();
});
test('production PCF bundle mounts with platform React/Fluent and opens the date picker', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await start(page);
  // This path is supplied by the test runner, never by production input.
  const bundlePath = '/@fs/' + process.cwd() + '/control/out/controls/DmsGrid/bundle.js';
  await page.evaluate(async (path) => {
    const load = (url: string) => import(/* @vite-ignore */ url);
    const React = await load('/node_modules/.vite/deps/react.js'),
      Fluent = await load('/node_modules/.vite/deps/@fluentui_react-components.js'),
      ReactDOM = await load('/node_modules/.vite/deps/react-dom.js'),
      mock = await load('/src/createMockContext.ts');
    const platform = window as unknown as Record<string, unknown>;
    platform.Reactv16 = React.default;
    platform.FluentUIReactv940 = Fluent;
    platform.ComponentFramework = {
      registerControl: (
        _name: string,
        Constructor: typeof import('../control/DmsGrid').DmsGrid,
      ) => {
        const control = new Constructor(),
          context = mock.createMockPcfContext(
            mock.createMockContext(Fluent.webLightTheme, '11111111-1111-4111-8111-111111111111'),
          );
        control.init(context);
        ReactDOM.default.render(control.updateView(context), document.getElementById('root'));
      },
    };
    await new Promise<void>((resolve, reject) => {
      const script = document.createElement('script');
      script.src = path;
      script.onload = () => resolve();
      script.onerror = () => reject(new Error('Bundle could not be loaded'));
      document.head.appendChild(script);
    });
  }, bundlePath);
  await expect(page.locator('[data-document-id]')).toHaveCount(22);
  await page.locator('[data-document-id]').first().click();
  await page.getByRole('button', { name: 'Edit properties', exact: true }).first().click();
  await expect(page.getByRole('textbox', { name: 'Description', exact: true })).toBeVisible();
  await page.getByPlaceholder('Choose a date').click();
  await expect(page.getByRole('button', { name: 'Go to today', exact: true })).toBeVisible();
  expect(errors).toEqual([]);
});

test('metadata conflicts preserve changed fields and require an explicit overwrite', async ({
  page,
  request,
}) => {
  await start(page);
  const row = page.locator('[data-document-id]').first(),
    id = await row.getAttribute('data-document-id');
  await row.click();
  await page.getByRole('button', { name: 'Edit properties', exact: true }).first().click();
  await page.getByRole('textbox', { name: 'Description', exact: true }).fill('My review');
  await request.patch('/api/data/v9.2/dms_documents(' + id + ')', {
    headers: { 'If-Match': '*' },
    data: { dms_description: 'A concurrent edit' },
  });
  await page.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(
    page.getByText(
      'Someone changed this document. Reload current values or confirm overwriting your changed fields.',
      { exact: true },
    ),
  ).toBeVisible();
  await expect(page.getByRole('textbox', { name: 'Description', exact: true })).toHaveValue(
    'My review',
  );
  await page.getByRole('button', { name: 'Overwrite', exact: true }).click();
  await page
    .getByRole('dialog', { name: 'Overwrite', exact: true })
    .getByRole('button', { name: 'Overwrite', exact: true })
    .click();
  await expect(page.getByText('Properties saved.', { exact: true })).toBeVisible();
  const response = await request.get('/api/data/v9.2/dms_documents(' + id + ')');
  expect((await response.json()).dms_description).toBe('My review');
});
