import { StorageProvider, registerProvider } from './registry';
import { DataverseClient, ApiError } from '../services/DataverseClient';
import { DocumentRow, FileContent } from '../types/Documents';
const blockSize = 4 * 1024 * 1024;
export function encode(bytes: Uint8Array): string {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 8192)
    binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return btoa(binary);
}
export function decode(base64: string): Uint8Array {
  const binary = atob(base64),
    bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
// TODO(SPIKE-S3/S6): all annotation action shapes and polymorphic parent bindings are isolated here.
class NoteProvider implements StorageProvider {
  readonly id = 'Note' as const;
  readonly capabilities = {
    canUpload: true,
    canAddLink: false,
    canDownload: true,
    canOpen: true,
    canPreviewInline: true,
    canDelete: true,
    ownsBinary: true,
    supportsBlockTransfer: true,
  };
  constructor(private client: DataverseClient) {}
  async upload(
    file: File,
    annotationId: string,
    entityName: string,
    recordId: string,
    progress: (percent: number) => void,
    signal: AbortSignal,
  ): Promise<void> {
    const metadata = await this.client.get<{ EntitySetName: string }>(
      `EntityDefinitions(LogicalName='${entityName}')?$select=EntitySetName`,
      signal,
    );
    const target = {
      annotationid: annotationId,
      subject: file.name,
      filename: file.name,
      mimetype: file.type || 'application/octet-stream',
      isdocument: true,
      [`objectid_${entityName}@odata.bind`]: `/${metadata.EntitySetName}(${recordId})`,
    };
    if (file.size <= blockSize) {
      try {
        await this.client.post(
          'annotations',
          { ...target, documentbody: encode(new Uint8Array(await file.arrayBuffer())) },
          signal,
        );
      } catch (error) {
        if (!(error instanceof ApiError && error.kind === 'conflict')) throw error;
        const existing = await this.client.get<{ filesize: number }>(
          `annotations(${annotationId})?$select=filesize`,
          signal,
        );
        if (existing.filesize !== file.size) throw error;
      }
      progress(100);
      return;
    }
    const { FileContinuationToken } = await this.client.action<{ FileContinuationToken: string }>(
        'InitializeAnnotationBlocksUpload',
        { Target: target },
        signal,
      ),
      BlockList: string[] = [];
    for (let offset = 0; offset < file.size; offset += blockSize) {
      if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
      const BlockId = btoa(String(BlockList.length).padStart(8, '0'));
      const data = new Uint8Array(await file.slice(offset, offset + blockSize).arrayBuffer());
      await this.client.action(
        'UploadBlock',
        { FileContinuationToken, BlockId, BlockData: encode(data) },
        signal,
      );
      BlockList.push(BlockId);
      progress(Math.round((Math.min(file.size, offset + blockSize) / file.size) * 95));
    }
    await this.client.action(
      'CommitAnnotationBlocksUpload',
      { FileContinuationToken, BlockList, Target: target },
      signal,
    );
    progress(100);
  }
  async getContent(
    row: DocumentRow,
    progress: (percent: number) => void = () => undefined,
    signal?: AbortSignal,
  ): Promise<FileContent> {
    if (row.sizeKb <= 4096) {
      const annotation = await this.client.get<{
        documentbody: string;
        filename: string;
        mimetype: string;
      }>(`annotations(${row.storageRef})?$select=documentbody,filename,mimetype`, signal);
      const bytes = decode(annotation.documentbody);
      return {
        kind: 'blob',
        blob: new Blob([bytes], { type: annotation.mimetype }),
        name: row.name,
        contentType: annotation.mimetype,
      };
    }
    const init = await this.client.action<{
        FileContinuationToken: string;
        FileSizeInBytes: number;
        FileName: string;
      }>(
        'InitializeAnnotationBlocksDownload',
        { Target: { annotationid: row.storageRef } },
        signal,
      ),
      parts: Uint8Array[] = [];
    for (let offset = 0; offset < init.FileSizeInBytes; offset += blockSize) {
      const { Data } = await this.client.action<{ Data: string }>(
        'DownloadBlock',
        {
          FileContinuationToken: init.FileContinuationToken,
          Offset: offset,
          BlockLength: blockSize,
        },
        signal,
      );
      parts.push(decode(Data));
      progress(
        Math.round(
          (Math.min(init.FileSizeInBytes, offset + blockSize) / init.FileSizeInBytes) * 100,
        ),
      );
    }
    return {
      kind: 'blob',
      blob: new Blob(parts, { type: row.contentType }),
      name: row.name,
      contentType: row.contentType,
    };
  }
}
registerProvider('Note', (client) => new NoteProvider(client));
