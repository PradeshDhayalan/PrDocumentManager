export type AttributeValue = string | number | boolean | null;
export type RawEntity = Record<string, AttributeValue>;
export type ProviderId = 'Note' | 'SharePoint' | 'AzureBlob' | 'External';
export interface DocumentRow {
  id: string;
  name: string;
  originalName: string;
  description: string;
  provider: ProviderId;
  storageRef: string;
  contentType: string;
  sizeKb: number;
  checksum: string;
  documentType: number | null;
  status: number | null;
  expiry: string;
  uploadState: number;
  modified: string;
  created: string;
  modifiedBy: string;
  createdBy: string;
  modifiedById: string;
  etag: string;
  raw: RawEntity;
}
export interface Migration {
  state: 'NotStarted' | 'Running' | 'Completed' | 'Failed';
  processed: number;
  total: number;
}
export interface ClientConfig {
  schemaVersion: 1;
  activeProvider: 'Note' | 'SharePoint';
  entity: {
    logicalName: string;
    enabled: boolean;
    visibleColumns: string[];
    editableColumns: string[];
    maxFileSizeMb: number;
    allowedExtensions: string[];
    blockedExtensions: string[];
  };
  org: { maxUploadFileSizeBytes: number };
  sharePoint: { allowedHosts: string[] };
  migration: Migration;
}
export interface Page<T> {
  value: T[];
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
}
export interface DocumentPage {
  rows: DocumentRow[];
  nextCursor?: string;
  totalCount: number;
}
export type Preset = 'all' | 'recent' | 'expiring' | 'expired' | 'mine';
export interface ListQuery {
  recordId: string;
  entityName: string;
  search: string;
  orderBy: string;
  preset: Preset;
  typeFilter: string;
  pageSize: number;
}
export interface Capabilities {
  canUpload: boolean;
  canAddLink: boolean;
  canDownload: boolean;
  canOpen: boolean;
  canPreviewInline: boolean;
  canDelete: boolean;
  ownsBinary: boolean;
  supportsBlockTransfer: boolean;
}
export type FileContent =
  | { kind: 'blob'; blob: Blob; name: string; contentType: string }
  | { kind: 'url'; url: string };
export type T = (key: string, values?: Record<string, string | number>) => string;
