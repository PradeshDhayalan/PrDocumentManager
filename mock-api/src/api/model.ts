export type Value = string | number | boolean | null;
export type Entity = Record<string, Value>;
export type MigrationState = 'NotStarted' | 'Running' | 'Completed' | 'Failed';
export interface Migration {
  state: MigrationState;
  processed: number;
  total: number;
  startedAt?: number;
}
export interface Snapshot {
  customColumns?: { name: string; label: string; type: string }[];
  documents: Entity[];
  annotations: Entity[];
  migrations: Record<string, Migration>;
  binarySeeds?: Record<string, string>;
}
export interface MockConfig {
  latencyMs: number;
  failureRate: number;
  failNextUploadAtBlock: number | null;
  failWith: 429 | 500 | 412 | 'timeout' | 'offline' | null;
  maxUploadFileSizeBytes: number;
  activeProvider: 'Note' | 'SharePoint';
  spBlockFraming: boolean;
}
export const defaultConfig: MockConfig = {
  latencyMs: 0,
  failureRate: 0,
  failNextUploadAtBlock: null,
  failWith: null,
  maxUploadFileSizeBytes: 128 * 1024 * 1024,
  activeProvider: 'Note',
  spBlockFraming: false,
};
export const records = {
  contoso: '11111111-1111-4111-8111-111111111111',
  fabrikam: '22222222-2222-4222-8222-222222222222',
  empty: '33333333-3333-4333-8333-333333333333',
  legacy: '44444444-4444-4444-8444-444444444444',
};
export const choices: Record<string, Record<number, string>> = {
  dms_provider: {
    100000000: 'Note',
    100000001: 'SharePoint',
    100000002: 'AzureBlob',
    100000003: 'External',
  },
  dms_documenttype: {
    100000000: 'Contract',
    100000001: 'Proposal',
    100000002: 'Finance',
    100000003: 'Legal',
    100000004: 'Media',
    100000005: 'Other',
  },
  dms_documentstatus: { 100000000: 'Draft', 100000001: 'Active', 100000002: 'Archived' },
  dms_uploadstate: { 100000000: 'Pending', 100000001: 'Available', 100000002: 'Failed' },
};
export const users = [
  'MOD Administrator',
  'Michael Bose',
  'Alex Wilber',
  'Irvin Sayers',
  'Megan Bowen',
  'Adele Vance',
];
export const userIds = users.map(
  (_, i) => `aaaaaaaa-aaaa-4aaa-8aaa-${String(i + 1).padStart(12, '0')}`,
);
export const formattedSuffix = '@OData.Community.Display.V1.FormattedValue';
export function formatEntity(entity: Entity): Entity {
  const output = { ...entity };
  for (const [key, value] of Object.entries(entity)) {
    if (choices[key] && typeof value === 'number')
      output[key + formattedSuffix] = choices[key][value] || '';
    if (['_modifiedby_value', '_createdby_value', '_ownerid_value'].includes(key))
      output[key + formattedSuffix] = users[userIds.indexOf(String(value))] || 'Unknown user';
    if (['createdon', 'modifiedon', 'dms_expirydate'].includes(key) && typeof value === 'string')
      output[key + formattedSuffix] = new Intl.DateTimeFormat('en-US', {
        dateStyle: 'medium',
        timeZone: 'UTC',
      }).format(new Date(value));
  }
  return output;
}
export class ApiError extends Error {
  constructor(
    public status: number,
    public code: string,
    message: string,
  ) {
    super(message);
  }
}
export const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export function requireGuid(value: unknown): string {
  if (typeof value !== 'string' || !guid.test(value))
    throw new ApiError(400, 'InvalidGuid', 'A valid GUID is required.');
  return value.toLowerCase();
}
export function payload(value: unknown): Entity {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new ApiError(400, 'InvalidPayload', 'A JSON object is required.');
  const result: Entity = {};
  for (const [key, field] of Object.entries(value)) {
    if (field !== null && !['string', 'number', 'boolean'].includes(typeof field))
      throw new ApiError(400, 'InvalidPayload', 'Only primitive attribute values are accepted.');
    result[key] = field as Value;
  }
  return result;
}
export function validateDocument(row: Entity): void {
  for (const key of ['dms_name', 'dms_regardingid', 'dms_regardingtype'])
    if (typeof row[key] !== 'string' || !String(row[key]).trim())
      throw new ApiError(400, 'RequiredField', `Required attribute: ${key}`);
  requireGuid(row.dms_documentid);
  requireGuid(row.dms_regardingid);
  for (const [key, options] of Object.entries(choices)) {
    if (
      row[key] !== undefined &&
      row[key] !== null &&
      (typeof row[key] !== 'number' || !options[Number(row[key])])
    )
      throw new ApiError(400, 'InvalidChoice', `Invalid choice: ${key}`);
  }
  if (row.dms_provider === undefined)
    throw new ApiError(400, 'RequiredField', 'Storage provider is required.');
  for (const [key, limit] of Object.entries({
    dms_name: 255,
    dms_originalfilename: 255,
    dms_description: 4000,
    dms_regardingtype: 64,
    dms_storageref: 2000,
    dms_contenttype: 128,
    dms_checksum: 64,
  }))
    if (typeof row[key] === 'string' && String(row[key]).length > limit)
      throw new ApiError(400, 'MaxLength', `Attribute exceeds length: ${key}`);
  if (row.dms_expirydate && !/^\d{4}-\d{2}-\d{2}$/.test(String(row.dms_expirydate)))
    throw new ApiError(400, 'InvalidDate', 'Date-only value is required.');
}
