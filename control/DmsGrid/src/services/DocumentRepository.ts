import { DataverseClient, ApiError } from './DataverseClient';
import {
  RawEntity,
  DocumentRow,
  DocumentPage,
  ListQuery,
  ClientConfig,
  Migration,
} from '../types/Documents';
export const providerChoices = {
  Note: 100000000,
  SharePoint: 100000001,
  AzureBlob: 100000002,
  External: 100000003,
};
const formatted = '@OData.Community.Display.V1.FormattedValue';
export function mapDocument(raw: RawEntity): DocumentRow {
  const provider = Object.entries(providerChoices).find(
    ([, value]) => value === raw.dms_provider,
  )?.[0] as DocumentRow['provider'] | undefined;
  return {
    id: String(raw.dms_documentid),
    name: String(raw.dms_name || ''),
    originalName: String(raw.dms_originalfilename || ''),
    description: String(raw.dms_description || ''),
    provider: provider || 'External',
    storageRef: String(raw.dms_storageref || ''),
    contentType: String(raw.dms_contenttype || ''),
    sizeKb: Number(raw.dms_filesizekb || 0),
    checksum: String(raw.dms_checksum || ''),
    documentType: typeof raw.dms_documenttype === 'number' ? raw.dms_documenttype : null,
    status: typeof raw.dms_documentstatus === 'number' ? raw.dms_documentstatus : null,
    expiry: String(raw.dms_expirydate || ''),
    uploadState: Number(raw.dms_uploadstate),
    modified: String(raw.modifiedon || ''),
    created: String(raw.createdon || ''),
    modifiedBy: String(raw['_modifiedby_value' + formatted] || ''),
    createdBy: String(raw['_createdby_value' + formatted] || ''),
    modifiedById: String(raw._modifiedby_value || ''),
    etag: String(raw['@odata.etag']),
    raw,
  };
}
export const quote = (value: string) => `'${value.replace(/'/g, "''")}'`;
export class DocumentRepository {
  constructor(public client: DataverseClient) {}
  async list(query: ListQuery, cursor?: string, signal?: AbortSignal): Promise<DocumentPage> {
    const conditions = [
      `dms_regardingid eq ${quote(query.recordId)}`,
      `dms_regardingtype eq ${quote(query.entityName)}`,
    ];
    if (query.search)
      conditions.push(
        `(contains(dms_name,${quote(query.search)}) or contains(dms_description,${quote(query.search)}))`,
      );
    if (query.typeFilter) conditions.push(`dms_documenttype eq ${Number(query.typeFilter)}`);
    const now = new Date(),
      today = now.toISOString().slice(0, 10),
      month = new Date(now.getTime() + 30 * 86400000).toISOString().slice(0, 10);
    if (query.preset === 'expired') conditions.push(`dms_expirydate lt ${quote(today)}`);
    if (query.preset === 'expiring')
      conditions.push(`(dms_expirydate ge ${quote(today)} and dms_expirydate le ${quote(month)})`);
    if (query.preset === 'recent')
      conditions.push(
        `modifiedon ge ${quote(new Date(now.getTime() - 7 * 86400000).toISOString())}`,
      );
    // TODO(SPIKE-S4): the host supplies the current user in real Dataverse; mock uses its seeded administrator.
    if (query.preset === 'mine')
      conditions.push(`_createdby_value eq 'aaaaaaaa-aaaa-4aaa-8aaa-000000000001'`);
    const params = new URLSearchParams({
      $filter: conditions.join(' and '),
      $orderby: query.orderBy,
      $count: 'true',
    });
    const page = await this.client.getPage<RawEntity>(cursor || `dms_documents?${params}`, signal);
    return {
      rows: page.value.map(mapDocument),
      nextCursor: page['@odata.nextLink'],
      totalCount: page['@odata.count'] ?? page.value.length,
    };
  }
  async get(id: string, signal?: AbortSignal): Promise<DocumentRow> {
    return mapDocument(await this.client.get<RawEntity>(`dms_documents(${id})`, signal));
  }
  async create(attributes: RawEntity, signal?: AbortSignal): Promise<DocumentRow> {
    const id = String(attributes.dms_documentid);
    try {
      return mapDocument(
        await this.client.request<RawEntity>(`dms_documents(${id})`, 'PATCH', attributes, signal, {
          'If-None-Match': '*',
        }),
      );
    } catch (error) {
      if (error instanceof ApiError && error.kind === 'conflict') {
        const existing = await this.get(id, signal);
        if (
          existing.storageRef === attributes.dms_storageref &&
          existing.name === attributes.dms_name
        )
          return existing;
      }
      throw error;
    }
  }
  async update(
    row: DocumentRow,
    values: RawEntity,
    overwrite = false,
    signal?: AbortSignal,
  ): Promise<DocumentRow> {
    return mapDocument(
      await this.client.patch<RawEntity>(
        `dms_documents(${row.id})`,
        values,
        overwrite ? '*' : row.etag,
        signal,
      ),
    );
  }
  delete(row: DocumentRow): Promise<void> {
    return this.client.delete(`dms_documents(${row.id})`, row.etag);
  }
  async config(entityName: string, recordId: string, signal?: AbortSignal): Promise<ClientConfig> {
    const { ConfigJson } = await this.client.action<{ ConfigJson: string }>(
      'dms_GetClientConfig',
      { EntityLogicalName: entityName, RegardingId: recordId },
      signal,
    );
    return JSON.parse(ConfigJson) as ClientConfig;
  }
  migration(
    entityName: string,
    recordId: string,
    ensure = false,
    signal?: AbortSignal,
  ): Promise<Migration> {
    return this.client.action<Migration>(
      ensure ? 'dms_EnsureMigration' : 'dms_GetMigrationStatus',
      { EntityLogicalName: entityName, RegardingId: recordId },
      signal,
    );
  }
}
export async function parallelLimit<T, R>(
  values: T[],
  limit: number,
  work: (value: T) => Promise<R>,
): Promise<PromiseSettledResult<R>[]> {
  const results: PromiseSettledResult<R>[] = new Array(values.length);
  let next = 0;
  await Promise.all(
    Array.from({ length: Math.min(limit, values.length) }, async () => {
      while (next < values.length) {
        const index = next++;
        try {
          results[index] = { status: 'fulfilled', value: await work(values[index]) };
        } catch (reason) {
          results[index] = { status: 'rejected', reason };
        }
      }
    }),
  );
  return results;
}
