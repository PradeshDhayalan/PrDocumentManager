import { HostContext } from '../types/HostContext';
export type Value = string | number | boolean | null;
export type Entity = Record<string, Value>;
export interface DocumentRow {
  id: string;
  name: string;
  fileType: string;
  type: string;
  status: string;
  expiry: string;
  modified: string;
  by: string;
  source: string;
  size: string;
  raw: Entity;
}
export interface Attribute {
  LogicalName: string;
  AttributeType: string;
  DisplayName: { UserLocalizedLabel: { Label: string } };
  RequiredLevel: { Value: string };
  MaxLength?: number;
  IsValidForUpdate: boolean;
  OptionSet?: { Options?: { Value: number; Label: { UserLocalizedLabel: { Label: string } } }[] };
}
export interface ClientConfig {
  activeProvider: 'Note' | 'SharePoint';
  entity: {
    enabled: boolean;
    visibleColumns: string[];
    editableColumns: string[];
    maxFileSizeMb: number;
    blockedExtensions: string[];
    allowedExtensions: string[];
  };
  org: { maxUploadFileSizeBytes: number };
  sharePoint: { allowedHosts: string[] };
  migration: { state: string; processed: number; total: number };
}
export interface Query {
  searchField?: string;
  fields?: Entity;
  search: string;
  type: string;
  status: string;
  preset: string;
  sort: string;
  descending: boolean;
}
export const formattedSuffix = '@OData.Community.Display.V1.FormattedValue';
export function label(row: Entity, field: string): string {
  return String(row[field + formattedSuffix] ?? row[field] ?? '—');
}
export function fileType(name: string, contentType = ''): string {
  if (contentType === 'application/x-sharepoint-link') return 'Link';
  const ext = name.split('.').pop()?.toLowerCase();
  return (
    (
      {
        doc: 'Word',
        docx: 'Word',
        xls: 'Excel',
        xlsx: 'Excel',
        ppt: 'PowerPoint',
        pptx: 'PowerPoint',
        pdf: 'PDF',
        jpg: 'Photo',
        jpeg: 'Photo',
        png: 'Photo',
        gif: 'Photo',
        webp: 'Photo',
        txt: 'Text',
        md: 'Text',
        csv: 'Text',
        json: 'Text',
        mp4: 'Video',
      } as Record<string, string>
    )[ext || ''] || 'File'
  );
}
export function mapDocument(raw: Entity): DocumentRow {
  const upload = label(raw, 'dms_uploadstate');
  return {
    id: String(raw.dms_documentid),
    name: String(raw.dms_name),
    fileType: fileType(
      String(raw.dms_originalfilename || raw.dms_name),
      String(raw.dms_contenttype || ''),
    ),
    type: label(raw, 'dms_documenttype'),
    status: upload === 'Failed' || upload === 'Pending' ? upload : label(raw, 'dms_documentstatus'),
    expiry: raw.dms_expirydate ? String(raw.dms_expirydate) : '',
    modified: label(raw, 'modifiedon'),
    by: label(raw, '_modifiedby_value'),
    source: label(raw, 'dms_provider'),
    size:
      raw.dms_filesizekb == null
        ? '—'
        : Number(raw.dms_filesizekb) >= 1024
          ? `${(Number(raw.dms_filesizekb) / 1024).toFixed(1)} MB`
          : `${raw.dms_filesizekb} KB`,
    raw,
  };
}
export class RequestError extends Error {
  constructor(
    public status: number,
    message: string,
  ) {
    super(message);
  }
}
const quote = (value: string) => `'${value.replace(/'/g, "''")}'`;
const blockSize = 4 * 1024 * 1024;
function encode(bytes: Uint8Array): string {
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 8192)
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  return btoa(binary);
}
function decode(value: string): Uint8Array<ArrayBuffer> {
  return Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
}
export class DocumentClient {
  private base: string;
  constructor(private host: HostContext) {
    this.base = `${host.clientUrl.replace(/\/$/, '')}/api/data/v9.2/`;
  }
  async request<T>(
    route: string,
    options: RequestInit = {},
    signal?: AbortSignal,
    retry = false,
  ): Promise<T> {
    const url = new URL(route, this.base);
    if (url.origin !== new URL(this.base).origin || !url.pathname.startsWith('/api/data/v9.2/'))
      throw new Error('Invalid API page URL.');
    for (let attempt = 0; ; attempt++) {
      const response = await fetch(url, {
        ...options,
        signal,
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
          Prefer: 'odata.include-annotations="*"',
          ...options.headers,
        },
      });
      if (retry && [429, 500, 502, 503, 504].includes(response.status) && attempt < 3) {
        await new Promise<void>((resolve, reject) => {
          const timer = setTimeout(
            () => {
              signal?.removeEventListener('abort', abort);
              resolve();
            },
            Math.min(10000, Number(response.headers.get('Retry-After') || 2 ** attempt) * 1000),
          );
          function abort() {
            clearTimeout(timer);
            reject(new DOMException('Cancelled', 'AbortError'));
          }
          signal?.addEventListener('abort', abort, { once: true });
          if (signal?.aborted) abort();
        });
        continue;
      }
      if (!response.ok) {
        const body = await response.json().catch(() => null);
        throw new RequestError(
          response.status,
          response.status === 412
            ? 'This document changed since it was loaded. Refresh before trying again.'
            : body?.error?.message || `Request failed (${response.status}).`,
        );
      }
      return response.status === 204 ? (undefined as T) : ((await response.json()) as T);
    }
  }
  async config(signal?: AbortSignal): Promise<ClientConfig> {
    const data = await this.request<{ ConfigJson: string }>(
      'dms_GetClientConfig',
      {
        method: 'POST',
        body: JSON.stringify({
          EntityLogicalName: this.host.entityName,
          RegardingId: this.host.recordId,
        }),
      },
      signal,
    );
    return JSON.parse(data.ConfigJson);
  }
  async attributes(signal?: AbortSignal): Promise<Attribute[]> {
    const data = await this.request<{ Attributes: Attribute[] }>(
      "EntityDefinitions(LogicalName='dms_document')?$expand=Attributes",
      {},
      signal,
    );
    return data.Attributes;
  }
  async list(
    query: Query,
    signal?: AbortSignal,
    next?: string,
  ): Promise<{ rows: DocumentRow[]; count: number; next?: string }> {
    const filters = [
      `dms_regardingid eq ${quote(this.host.recordId)}`,
      `dms_regardingtype eq ${quote(this.host.entityName)}`,
    ];
    const field = query.searchField || 'dms_name';
    if (!/^[a-z_][a-z0-9_]*$/.test(field)) throw new Error('Invalid search column.');
    if (query.search.trim()) filters.push(`contains(${field},${quote(query.search.trim())})`);
    for (const [key, value] of Object.entries(query.fields || {})) {
      if (!/^[a-z_][a-z0-9_]*$/.test(key)) throw new Error('Invalid filter column.');
      if (value === null || value === '') continue;
      filters.push(
        typeof value === 'string' ? `contains(${key},${quote(value)})` : `${key} eq ${value}`,
      );
    }
    if (query.type) filters.push(`dms_documenttype eq ${Number(query.type)}`);
    if (query.status) filters.push(`dms_documentstatus eq ${Number(query.status)}`);
    const today = new Date().toISOString().slice(0, 10);
    if (query.preset === 'Expired') filters.push(`dms_expirydate lt ${quote(today)}`);
    if (query.preset === 'Expiring soon') {
      const soon = new Date(Date.now() + 30 * 86400000).toISOString().slice(0, 10);
      filters.push(`dms_expirydate ge ${quote(today)} and dms_expirydate le ${quote(soon)}`);
    }
    if (query.preset === 'Recently modified')
      filters.push(`modifiedon ge ${quote(new Date(Date.now() - 7 * 86400000).toISOString())}`);
    if (query.preset === 'Added by me' && this.host.userId)
      filters.push(`_createdby_value eq ${quote(this.host.userId)}`);
    const parameters = new URLSearchParams({
      $filter: filters.join(' and '),
      $orderby: `${query.sort} ${query.descending ? 'desc' : 'asc'}`,
      $count: 'true',
    });
    const data = await this.request<{
      value: Entity[];
      '@odata.count': number;
      '@odata.nextLink'?: string;
    }>(
      next || `dms_documents?${parameters}`,
      {
        headers: {
          Prefer: `odata.include-annotations="*",odata.maxpagesize=${Math.max(1, Math.min(250, this.host.pageSize || 10))}`,
        },
      },
      signal,
    );
    return {
      rows: data.value.map(mapDocument),
      count: data['@odata.count'],
      // The mock sits behind Vite; keep continuation requests on the host's proxy.
      next: data['@odata.nextLink'] ? this.pageRoute(data['@odata.nextLink']) : undefined,
    };
  }
  private pageRoute(link: string): string {
    const url = new URL(link, this.base);
    if (url.pathname !== '/api/data/v9.2/dms_documents')
      throw new Error('Invalid document page URL.');
    return url.pathname + url.search;
  }
  async save(row: DocumentRow, changes: Entity, signal?: AbortSignal): Promise<void> {
    await this.request(
      `dms_documents(${row.id})`,
      {
        method: 'PATCH',
        headers: { 'If-Match': String(row.raw['@odata.etag']) },
        body: JSON.stringify(changes),
      },
      signal,
    );
  }
  async remove(row: DocumentRow, signal?: AbortSignal): Promise<void> {
    await this.request(
      `dms_documents(${row.id})`,
      { method: 'DELETE', headers: { 'If-Match': String(row.raw['@odata.etag']) } },
      signal,
    );
  }
  async addLink(
    url: string,
    name: string,
    config: ClientConfig,
    signal?: AbortSignal,
  ): Promise<void> {
    const parsed = new URL(url);
    if (
      parsed.protocol !== 'https:' ||
      parsed.username ||
      parsed.password ||
      !config.sharePoint.allowedHosts.some((host) =>
        host.startsWith('*.')
          ? parsed.hostname.endsWith(host.slice(1)) && parsed.hostname !== host.slice(2)
          : parsed.hostname === host,
      )
    )
      throw new Error('Enter an HTTPS link on an allowed SharePoint host.');
    await this.request(
      'dms_documents',
      {
        method: 'POST',
        body: JSON.stringify({
          dms_name:
            name.trim() ||
            decodeURIComponent(
              parsed.pathname.split('/').filter(Boolean).pop() || 'SharePoint document',
            ),
          dms_regardingid: this.host.recordId,
          dms_regardingtype: this.host.entityName,
          dms_provider: 100000001,
          dms_storageref: parsed.href,
          dms_contenttype: 'application/x-sharepoint-link',
          dms_documenttype: 100000005,
          dms_documentstatus: 100000001,
          dms_uploadstate: 100000001,
        }),
      },
      signal,
    );
  }
  async upload(
    file: File,
    config: ClientConfig,
    progress: (value: number) => void,
    signal: AbortSignal,
  ): Promise<void> {
    const ext = file.name.split('.').pop()?.toLowerCase() || '';
    if (
      config.entity.blockedExtensions.includes(ext) ||
      (config.entity.allowedExtensions.length && !config.entity.allowedExtensions.includes(ext))
    )
      throw new Error(`${file.name}: this file extension is not allowed.`);
    if (
      file.size >
      Math.min(config.entity.maxFileSizeMb * 1024 * 1024, config.org.maxUploadFileSizeBytes)
    )
      throw new Error(`${file.name}: file exceeds the upload size limit.`);
    if (file.name.length > 255) throw new Error('The filename must be 255 characters or fewer.');
    const id = crypto.randomUUID(),
      annotationId = crypto.randomUUID();
    const metadata = await this.request<{ EntitySetName: string }>(
      `EntityDefinitions(LogicalName=${quote(this.host.entityName)})`,
      {},
      signal,
    );
    const target = {
      annotationid: annotationId,
      filename: file.name,
      mimetype: file.type || 'application/octet-stream',
      [`objectid_${this.host.entityName}@odata.bind`]: `/${metadata.EntitySetName}(${this.host.recordId})`,
    };
    // Create an owned pending document first so compensation can cascade through the server.
    const pending = await this.request<Entity>(
      'dms_documents',
      {
        method: 'POST',
        body: JSON.stringify({
          dms_documentid: id,
          dms_name: file.name,
          dms_originalfilename: file.name,
          dms_regardingid: this.host.recordId,
          dms_regardingtype: this.host.entityName,
          dms_provider: 100000000,
          dms_storageref: annotationId,
          dms_contenttype: target.mimetype,
          dms_filesizekb: Math.ceil(file.size / 1024),
          dms_documenttype: 100000005,
          dms_documentstatus: 100000001,
          dms_uploadstate: 100000000,
        }),
      },
      signal,
    );
    try {
      if (file.size <= blockSize) {
        await this.request(
          'annotations',
          {
            method: 'POST',
            body: JSON.stringify({
              ...target,
              documentbody: encode(new Uint8Array(await file.arrayBuffer())),
            }),
          },
          signal,
        );
        progress(0.9);
      } else {
        const session = await this.request<{ FileContinuationToken: string }>(
          'InitializeAnnotationBlocksUpload',
          { method: 'POST', body: JSON.stringify({ Target: target }) },
          signal,
        );
        const ids: string[] = [];
        for (let offset = 0; offset < file.size; offset += blockSize) {
          const blockId = btoa(String(ids.length).padStart(8, '0'));
          ids.push(blockId);
          const bytes = new Uint8Array(await file.slice(offset, offset + blockSize).arrayBuffer());
          await this.request(
            'UploadBlock',
            {
              method: 'POST',
              body: JSON.stringify({
                FileContinuationToken: session.FileContinuationToken,
                BlockId: blockId,
                BlockData: encode(bytes),
              }),
            },
            signal,
            true,
          );
          progress(Math.min(0.9, ((offset + bytes.length) / file.size) * 0.9));
        }
        await this.request(
          'CommitAnnotationBlocksUpload',
          {
            method: 'POST',
            body: JSON.stringify({
              FileContinuationToken: session.FileContinuationToken,
              BlockList: ids,
            }),
          },
          signal,
          true,
        );
      }
      await this.save(mapDocument(pending), { dms_uploadstate: 100000001 }, signal);
      progress(1);
    } catch (error) {
      // Never delete annotations from the browser; the server checks ownership.
      try {
        await this.remove(mapDocument(pending));
      } catch {
        throw new Error(
          `${file.name}: upload did not finish and cleanup failed. Refresh to review the pending document.`,
        );
      }
      throw error;
    }
  }
  async download(row: DocumentRow, signal?: AbortSignal): Promise<Blob> {
    if (row.source !== 'Note' || !row.raw.dms_storageref || row.raw.dms_uploadstate !== 100000001)
      throw new Error('A downloadable local file is not available for this document.');
    const session = await this.request<{ FileContinuationToken: string; FileSizeInBytes: number }>(
      'InitializeAnnotationBlocksDownload',
      {
        method: 'POST',
        body: JSON.stringify({ Target: { annotationid: row.raw.dms_storageref } }),
      },
      signal,
    );
    const parts: Uint8Array<ArrayBuffer>[] = [];
    for (let offset = 0; offset < session.FileSizeInBytes; offset += blockSize) {
      const block = await this.request<{ Data: string }>(
        'DownloadBlock',
        {
          method: 'POST',
          body: JSON.stringify({
            FileContinuationToken: session.FileContinuationToken,
            Offset: offset,
            BlockLength: Math.min(blockSize, session.FileSizeInBytes - offset),
          }),
        },
        signal,
        true,
      );
      parts.push(decode(block.Data));
    }
    return new Blob(parts, { type: String(row.raw.dms_contenttype || 'application/octet-stream') });
  }
  async migration(start: boolean, signal?: AbortSignal): Promise<ClientConfig['migration']> {
    return this.request(
      start ? 'dms_EnsureMigration' : 'dms_GetMigrationStatus',
      { method: 'POST', body: JSON.stringify({ RegardingId: this.host.recordId }) },
      signal,
    );
  }
}
export function expiryState(
  value: string,
): { label: string; color: 'danger' | 'warning' } | undefined {
  if (!value) return;
  const days = Math.ceil(
    (new Date(value + 'T00:00:00').getTime() - new Date(new Date().toDateString()).getTime()) /
      86400000,
  );
  if (days < 0) return { label: 'Expired', color: 'danger' };
  if (days <= 30)
    return { label: days === 0 ? 'Expires today' : `Expires in ${days} days`, color: 'warning' };
}
export function safeExternalUrl(row: DocumentRow): string {
  const url = new URL(String(row.raw.dms_storageref));
  if (url.protocol !== 'https:' || url.username || url.password)
    throw new Error('The document link is not a valid HTTPS URL.');
  return url.href;
}
