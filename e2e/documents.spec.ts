import { test, expect, Page, APIRequestContext } from '@playwright/test';
import fs from 'node:fs/promises';
const api = '/api/data/v9.2';
const parent = '11111111-1111-4111-8111-111111111111';
async function rows(request: APIRequestContext, name: string) {
  const response = await request.get(`${api}/dms_documents`, {
    params: {
      $filter: `dms_regardingid eq '${parent}' and contains(dms_name,'${name.replace(/'/g, "''")}')`,
    },
  });
  expect(response.ok()).toBeTruthy();
  return (await response.json()).value;
}
async function cleanup(request: APIRequestContext, name: string) {
  for (const row of await rows(request, name))
    await request.delete(`${api}/dms_documents(${row.dms_documentid})`);
}
async function search(page: Page, name: string) {
  await page.getByRole('textbox', { name: 'Search documents', exact: true }).fill(name);
  await expect(
    page.getByRole('region', { name: 'Documents', exact: true }).locator('[aria-busy]'),
  ).toHaveAttribute('aria-busy', 'false');
  if (!name.includes("doesn't"))
    await expect(page.getByText('1 of 1 items', { exact: true })).toBeVisible();
}
test.afterEach(async ({ request }) => {
  await request.post('/__mock/config', {
    data: { activeProvider: 'Note', failNextUploadAtBlock: null, failWith: null },
  });
});
test('upload, preview, byte-equal download, edit, persist and delete a local file', async ({
  page,
  request,
}) => {
  const name = `e2e-lifecycle-${Date.now()}.txt`,
    contents = Buffer.from('Local document control lifecycle\nPersist this exact content.\n');
  try {
    await page.goto('/');
    await expect(page.getByRole('grid')).toBeVisible();
    await page
      .locator('input[type=file]')
      .setInputFiles({ name, mimeType: 'text/plain', buffer: contents });
    await page.getByRole('button', { name: 'Upload files', exact: true }).click();
    await expect(page.getByText('1 file uploaded.', { exact: true })).toBeVisible();
    await page.getByRole('dialog').getByRole('button', { name: 'Close', exact: true }).click();
    await search(page, name);
    await page.getByRole('checkbox', { name: `Select ${name}`, exact: true }).check();
    await page.getByRole('button', { name: 'Preview', exact: true }).click();
    await expect(page.getByRole('dialog').locator('pre')).toHaveText(contents.toString().trim());
    const downloadPromise = page.waitForEvent('download');
    await page.getByRole('dialog').getByRole('button', { name: 'Download', exact: true }).click();
    const downloaded = await downloadPromise;
    expect(downloaded.suggestedFilename()).toBe(name);
    expect(await fs.readFile((await downloaded.path())!)).toEqual(contents);
    await page.getByRole('dialog').getByRole('button', { name: 'Close', exact: true }).click();
    await page.getByRole('button', { name: 'Edit details', exact: true }).click();
    const edit = page.getByRole('dialog', { name: 'Edit details', exact: true });
    await edit
      .getByRole('textbox', { name: 'Description', exact: true })
      .fill('Saved description from the full control');
    await edit.getByRole('combobox', { name: 'Document type', exact: true }).click();
    await page.getByRole('option', { name: 'Legal', exact: true }).click();
    await edit.getByRole('button', { name: 'Save', exact: true }).click();
    await expect(edit).not.toBeVisible();
    const saved = (await rows(request, name))[0];
    expect(saved.dms_description).toBe('Saved description from the full control');
    expect(saved.dms_documenttype).toBe(100000003);
    await page.reload();
    await search(page, name);
    await page.getByRole('checkbox', { name: `Select ${name}`, exact: true }).check();
    await page.getByRole('button', { name: 'Details', exact: true }).click();
    await expect(
      page.getByText('Saved description from the full control', { exact: true }),
    ).toBeVisible();
    await page.getByRole('button', { name: 'Delete', exact: true }).click();
    await page
      .getByRole('dialog', { name: 'Delete', exact: true })
      .getByRole('button', { name: 'Delete', exact: true })
      .click();
    await expect(page.getByText('1 document deleted.', { exact: true })).toBeVisible();
    expect(await rows(request, name)).toHaveLength(0);
  } finally {
    await cleanup(request, name);
  }
});
test('large block upload retries a throttled block and downloads identical bytes', async ({
  page,
  request,
}) => {
  const name = `e2e-blocks-${Date.now()}.bin`,
    bytes = Buffer.alloc(5 * 1024 * 1024 + 17, 73);
  try {
    await request.post('/__mock/config', { data: { failNextUploadAtBlock: 2, failWith: 429 } });
    await page.goto('/');
    await expect(page.getByRole('grid')).toBeVisible();
    await page
      .locator('input[type=file]')
      .setInputFiles({ name, mimeType: 'application/octet-stream', buffer: bytes });
    await page.getByRole('button', { name: 'Upload files', exact: true }).click();
    await expect(page.getByText('1 file uploaded.', { exact: true })).toBeVisible({
      timeout: 20000,
    });
    await page.getByRole('dialog').getByRole('button', { name: 'Close', exact: true }).click();
    await search(page, name);
    await page.getByRole('checkbox', { name: `Select ${name}`, exact: true }).check();
    const downloadPromise = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Download', exact: true }).click();
    expect(await fs.readFile((await (await downloadPromise).path())!)).toEqual(bytes);
  } finally {
    await cleanup(request, name);
  }
});
test('validates SharePoint links, creates a persistent reference and preserves the external file on delete', async ({
  page,
  request,
}) => {
  const name = `e2e-link-${Date.now()}`;
  try {
    await page.goto('/');
    await expect(page.getByRole('grid')).toBeVisible();
    await request.post('/__mock/config', { data: { activeProvider: 'SharePoint' } });
    await page.reload();
    await page.getByRole('button', { name: 'New', exact: true }).click();
    await page.getByRole('menuitem', { name: 'Link', exact: true }).click();
    const dialog = page.getByRole('dialog', { name: 'Add link', exact: true });
    await dialog
      .getByRole('textbox', { name: 'SharePoint URL', exact: true })
      .fill('https://example.com/document');
    await dialog.getByRole('textbox', { name: 'Display name', exact: true }).fill(name);
    await dialog.getByRole('button', { name: 'Add link', exact: true }).click();
    await expect(dialog.getByRole('alert')).toContainText('allowed SharePoint host');
    await dialog
      .getByRole('textbox', { name: 'SharePoint URL', exact: true })
      .fill('https://contoso.sharepoint.com/sites/Sales/Test.docx');
    await dialog.getByRole('button', { name: 'Add link', exact: true }).click();
    await expect(dialog).not.toBeVisible();
    await search(page, name);
    await page.getByRole('checkbox', { name: `Select ${name}`, exact: true }).check();
    expect((await rows(request, name))[0].dms_provider).toBe(100000001);
    await page.getByRole('button', { name: 'Delete', exact: true }).click();
    const confirmation = page.getByRole('dialog', { name: 'Delete', exact: true });
    await expect(confirmation).toContainText('SharePoint file remains');
    await confirmation.getByRole('button', { name: 'Delete', exact: true }).click();
    await expect(page.getByText('1 document deleted.', { exact: true })).toBeVisible();
    expect(await rows(request, name)).toHaveLength(0);
  } finally {
    await cleanup(request, name);
  }
});
test('pages large libraries, escapes apostrophes in search, shares selection across views and handles mobile overflow', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.getByRole('grid')).toBeVisible();
  await page.getByRole('checkbox', { name: 'Select Site Photo.jpg', exact: true }).check();
  await page.getByRole('button', { name: 'Switch view', exact: true }).click();
  await page.getByRole('menuitem', { name: /Tiles view/ }).click();
  await expect(page.getByRole('option').filter({ hasText: 'Site Photo.jpg' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await page.getByRole('button', { name: '1 selected', exact: true }).click();
  await page.goto('/?record=fabrikam');
  await expect(page.getByText('10 of 1200 items', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Next page', exact: true }).click();
  await expect(page.getByText(/Page 2 of 120/)).toBeVisible();
  await search(page, "doesn't exist");
  await expect(page.getByText('No matching documents', { exact: true })).toBeVisible();
  await page.setViewportSize({ width: 320, height: 800 });
  await page.getByRole('button', { name: 'More commands', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Search documents', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Search documents', exact: true });
  await dialog
    .getByRole('textbox', { name: 'Search documents', exact: true })
    .fill('Fabrikam 0001');
  await dialog.getByRole('button', { name: 'Done', exact: true }).click();
  await expect(page.getByText('1 of 1 items', { exact: true })).toBeVisible();
  expect(
    await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth),
  ).toBeTruthy();
});

test('drag-selects rows and cards, keeps card names fixed and shows dates without expiry tags', async ({
  page,
}) => {
  await page.goto('/');
  const grid = page.getByRole('grid', { name: 'Documents', exact: true });
  await expect(grid).toBeVisible();
  await expect(grid.getByRole('gridcell').filter({ hasText: /^Expired$/ })).toHaveCount(0);
  await expect(grid.getByText(/Expires in \d+ days/)).toHaveCount(0);
  const listItems = grid.locator('[data-document-id]');
  const first = (await listItems.nth(0).boundingBox())!;
  const third = (await listItems.nth(2).boundingBox())!;
  await page.mouse.move(first.x + 180, first.y + 12);
  await page.mouse.down();
  await page.mouse.move(first.x + 320, third.y + third.height - 8, { steps: 12 });
  await expect(page.getByTestId('selection-marquee')).toBeVisible();
  await page.mouse.up();
  await expect(page.getByRole('button', { name: '3 selected', exact: true })).toBeVisible();
  await expect(page.getByTestId('selection-marquee')).toHaveCount(0);
  await page.getByRole('button', { name: 'Switch view', exact: true }).click();
  await page.getByRole('menuitem', { name: /Tiles view/ }).click();
  await page.getByRole('button', { name: '3 selected', exact: true }).click();
  const tiles = page
    .getByRole('listbox', { name: 'Documents', exact: true })
    .locator('[data-document-id]');
  const dimensions = await tiles.evaluateAll((elements) =>
    elements.map((element) => element.getBoundingClientRect().height),
  );
  expect(Math.max(...dimensions) - Math.min(...dimensions)).toBeLessThan(1);
  await expect(
    page
      .getByRole('listbox', { name: 'Documents', exact: true })
      .getByText('Active', { exact: true }),
  ).toHaveCount(0);
  await expect(
    page
      .getByRole('listbox', { name: 'Documents', exact: true })
      .getByText('Draft', { exact: true }),
  ).toHaveCount(0);
  const card1 = (await tiles.nth(0).boundingBox())!;
  const card2 = (await tiles.nth(1).boundingBox())!;
  await page.mouse.move(card1.x + 12, card1.y + 90);
  await page.mouse.down();
  await page.mouse.move(card2.x + card2.width - 12, card2.y + 150, { steps: 12 });
  await page.mouse.up();
  await expect(page.getByRole('button', { name: '2 selected', exact: true })).toBeVisible();
  await expect(tiles.nth(0)).toHaveAttribute('aria-selected', 'true');
  await expect(tiles.nth(1)).toHaveAttribute('aria-selected', 'true');
  await page.evaluate(() => document.fonts.ready);
  const session = await page.context().newCDPSession(page);
  await session.send('DOM.enable');
  await session.send('CSS.enable');
  const documentNode = await session.send('DOM.getDocument');
  const textNode = await session.send('DOM.querySelector', {
    nodeId: documentNode.root.nodeId,
    selector: '[data-document-id] [title]',
  });
  const renderedFonts = await session.send('CSS.getPlatformFontsForNode', {
    nodeId: textNode.nodeId,
  });
  expect(
    renderedFonts.fonts.some((font) => font.familyName === 'Segoe UI' && font.glyphCount > 0),
  ).toBeTruthy();
  const family = await page
    .getByRole('region', { name: 'Documents', exact: true })
    .evaluate((element) => getComputedStyle(element).getPropertyValue('--fontFamilyBase').trim());
  expect(family).toBe(
    "'Segoe UI','Segoe UI Web (West European)',-apple-system,BlinkMacSystemFont,Roboto,'Helvetica Neue',sans-serif",
  );
});

test('JSON custom buttons confirm, invoke a registered form handler, persist changes and enforce selection limits', async ({
  page,
  request,
}) => {
  const name = `e2e-custom-${Date.now()}.txt`;
  const created = await request.post(`${api}/dms_documents`, {
    data: {
      dms_name: name,
      dms_description: 'Original description',
      dms_regardingid: parent,
      dms_regardingtype: 'account',
      dms_provider: 100000000,
      dms_uploadstate: 100000001,
    },
  });
  expect(created.ok()).toBeTruthy();
  const clickCommand = async (label: string) => {
    await page.getByRole('button', { name: label, exact: true }).click();
  };
  try {
    await page.goto('/');
    await expect(page.getByRole('grid')).toBeVisible();
    await search(page, name);
    await page.getByRole('checkbox', { name: `Select ${name}`, exact: true }).check();
    await clickCommand('Mark as reviewed');
    const confirmation = page.getByRole('dialog', { name: 'Mark as reviewed', exact: true });
    await expect(confirmation).toContainText('1 document(s) selected.');
    await confirmation.getByRole('button', { name: 'Close', exact: true }).click();
    expect((await rows(request, name))[0].dms_description).toBe('Original description');
    await clickCommand('Mark as reviewed');
    await confirmation.getByRole('button', { name: 'Mark as reviewed', exact: true }).click();
    await expect(confirmation).not.toBeVisible();
    await expect(
      page.getByText('1 document(s) marked as reviewed.', { exact: true }),
    ).toBeVisible();
    expect((await rows(request, name))[0].dms_description).toBe('[Reviewed] Original description');
    await page.getByRole('checkbox', { name: `Select ${name}`, exact: true }).check();
    await clickCommand('Show selection');
    await expect(page.getByRole('status')).toContainText(name);
    await page.setViewportSize({ width: 320, height: 800 });
    await page.getByRole('button', { name: 'More commands', exact: true }).click();
    await expect(
      page.getByRole('menuitem', { name: 'Mark as reviewed', exact: true }),
    ).toBeVisible();
    await page.keyboard.press('Escape');
    await page.setViewportSize({ width: 1440, height: 1050 });
    await page.evaluate(() =>
      localStorage.setItem(
        'dms-pcf-configuration',
        JSON.stringify({
          pageSize: 50,
          grid: {
            filterColumns: ['dms_documenttype', 'dms_documentstatus'],
            searchField: 'dms_name',
            tileSize: 'medium',
          },
        }),
      ),
    );
    await page.goto('/?record=fabrikam');
    await expect(page.getByText('50 of 1200 items', { exact: true })).toBeVisible();
    await page
      .getByRole('checkbox', { name: 'Select all documents on this page', exact: true })
      .check();
    await expect(
      page.getByRole('button', { name: 'Mark as reviewed', exact: true }),
    ).toBeDisabled();
  } finally {
    await cleanup(request, name);
  }
});

test('resizes and reorders columns, configures built-ins and metadata, and spaces tiles with Microsoft file icons', async ({
  page,
}) => {
  await page.goto('/');
  const grid = page.getByRole('grid');
  await expect(grid).toBeVisible();
  const nameHeader = grid.locator('[role="columnheader"][data-column-id="name"]');
  const before = (await nameHeader.boundingBox())!;
  const handle = nameHeader.locator('.fui-TableResizeHandle');
  const bounds = (await handle.boundingBox())!;
  await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
  await page.mouse.down();
  await page.mouse.move(bounds.x + bounds.width / 2 + 85, bounds.y + bounds.height / 2, {
    steps: 10,
  });
  await page.mouse.up();
  await expect
    .poll(async () => (await nameHeader.boundingBox())!.width)
    .toBeGreaterThan(before.width + 70);
  await expect
    .poll(
      async () =>
        (await grid.locator('[role="gridcell"][data-column-id="name"]').first().boundingBox())!
          .width,
    )
    .toBeGreaterThan(before.width + 70);
  expect(
    await page
      .getByRole('checkbox', { name: /^Select / })
      .filter({ has: page.locator('[aria-checked="true"]') })
      .count(),
  ).toBe(0);
  const modifiedHeader = grid.locator('[role="columnheader"][data-column-id="modified"]');
  await modifiedHeader
    .getByRole('button', { name: 'Drag modified column', exact: true })
    .dragTo(nameHeader.getByRole('button', { name: 'Drag name column', exact: true }));
  const ids = () =>
    grid
      .locator('[role="columnheader"][data-column-id]')
      .evaluateAll((elements) => elements.map((element) => element.getAttribute('data-column-id')));
  await expect.poll(async () => (await ids()).indexOf('modified')).toBe(1);
  await page.getByRole('button', { name: 'More commands', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Edit columns', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Edit columns', exact: true });
  await dialog.getByRole('checkbox', { name: 'Name', exact: true }).uncheck();
  await dialog.getByRole('checkbox', { name: 'Status', exact: true }).uncheck();
  await dialog.getByRole('checkbox', { name: 'Description', exact: true }).check();
  await dialog.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(nameHeader).toHaveCount(0);
  await expect(grid.locator('[role="columnheader"][data-column-id="status"]')).toHaveCount(0);
  await expect(
    grid.locator('[role="columnheader"][data-column-id="dms_description"]'),
  ).toBeVisible();
  await page.getByRole('button', { name: 'More commands', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Edit columns', exact: true }).click();
  await dialog.getByRole('button', { name: 'Reset columns', exact: true }).click();
  await dialog.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(nameHeader).toBeVisible();
  await expect.poll(async () => (await nameHeader.boundingBox())!.width).toBe(before.width);
  const word = page.getByRole('img', { name: 'Word file', exact: true }).first();
  await expect(word).toHaveAttribute('src', /office\.net.*\/docx\.svg/);
  await expect
    .poll(() =>
      word.evaluate((element: HTMLImageElement) => element.complete && element.naturalWidth > 0),
    )
    .toBe(true);
  await page.getByRole('button', { name: 'Switch view', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Tiles view', exact: true }).click();
  const tiles = page.locator('[data-document-id]');
  const first = (await tiles.nth(0).boundingBox())!,
    second = (await tiles.nth(1).boundingBox())!;
  expect(second.x - first.x - first.width).toBeGreaterThanOrEqual(23);
});
