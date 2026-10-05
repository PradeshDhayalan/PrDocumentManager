import { Capabilities, DocumentRow, FileContent, ProviderId } from '../types/Documents';
import { DataverseClient } from '../services/DataverseClient';
export interface StorageProvider {
  readonly id: ProviderId;
  readonly capabilities: Capabilities;
  getContent(
    row: DocumentRow,
    progress?: (percent: number) => void,
    signal?: AbortSignal,
  ): Promise<FileContent>;
  upload?(
    file: File,
    annotationId: string,
    entityName: string,
    recordId: string,
    progress: (percent: number) => void,
    signal: AbortSignal,
  ): Promise<void>;
}
type Factory = (client: DataverseClient) => StorageProvider;
const factories = new Map<ProviderId, Factory>();
export function registerProvider(id: ProviderId, factory: Factory): void {
  factories.set(id, factory);
}
export function resolveProvider(id: ProviderId, client: DataverseClient): StorageProvider {
  return (
    factories.get(id)?.(client) || {
      id,
      capabilities: {
        canUpload: false,
        canAddLink: false,
        canDownload: false,
        canOpen: false,
        canPreviewInline: false,
        canDelete: true,
        ownsBinary: false,
        supportsBlockTransfer: false,
      },
      getContent: () => Promise.reject(new Error('provider.unavailable')),
    }
  );
}
export function usable(row: DocumentRow, provider: StorageProvider): boolean {
  return row.uploadState === 100000001 && !!row.storageRef && provider.capabilities.canOpen;
}
