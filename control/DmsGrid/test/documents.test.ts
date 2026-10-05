import { sanitizeFilename, validateSharePointUrl } from '../src/services/format';
import { mapDocument, DocumentRepository } from '../src/services/DocumentRepository';
import { DataverseClient } from '../src/services/DataverseClient';
test('filenames handle reserved names, invalid characters, whitespace, length and extension casing', () => {
  expect(sanitizeFilename(' CON.PDF ')).toBe('_CON.pdf');
  expect(sanitizeFilename(' a:<b>\u0001.JPG ')).toBe('a b.jpg');
  expect(sanitizeFilename('x'.repeat(300) + '.TXT')).toHaveLength(255);
});
test('SharePoint links enforce HTTPS, host boundaries, credentials and URL length', () => {
  expect(
    validateSharePointUrl('https://tenant.sharepoint.com/sites/a/file.pdf', ['*.sharepoint.com']),
  ).toContain('tenant.sharepoint.com');
  for (const url of [
    'http://tenant.sharepoint.com/a',
    'https://sharepoint.com.evil.test/a',
    'https://user:pass@tenant.sharepoint.com/a',
  ])
    expect(() => validateSharePointUrl(url, ['*.sharepoint.com'])).toThrow();
});
test('document mapper preserves raw choice values, formatted personas and unavailable provider identity', () => {
  const row = mapDocument({
    dms_documentid: 'a',
    dms_name: 'file.txt',
    dms_provider: 100000002,
    dms_uploadstate: 100000001,
    dms_documenttype: 100000004,
    '_modifiedby_value@OData.Community.Display.V1.FormattedValue': 'Megan Bowen',
    '@odata.etag': 'W/"4"',
  });
  expect(row.provider).toBe('AzureBlob');
  expect(row.modifiedBy).toBe('Megan Bowen');
  expect(row.documentType).toBe(100000004);
  expect(row.etag).toBe('W/"4"');
});
test('every server query scopes both parent ID and entity type and escapes apostrophes', async () => {
  const client = new DataverseClient('https://org.test');
  const get = jest.spyOn(client, 'getPage').mockResolvedValue({ value: [] });
  await new DocumentRepository(client).list({
    recordId: 'parent',
    entityName: 'account',
    search: "O'Brien",
    orderBy: 'modifiedon desc',
    preset: 'all',
    typeFilter: '',
    pageSize: 50,
  });
  const params = new URL(get.mock.calls[0][0], 'https://org.test').searchParams;
  expect(params.get('$filter')).toBe(
    "dms_regardingid eq 'parent' and dms_regardingtype eq 'account' and (contains(dms_name,'O''Brien') or contains(dms_description,'O''Brien'))",
  );
});
test('API rejects cross-origin paging links before fetching', async () => {
  const client = new DataverseClient('https://org.test');
  await expect(client.get('https://evil.test/api/data/v9.2/dms_documents')).rejects.toMatchObject({
    kind: 'validation',
  });
});
