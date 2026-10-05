import { StorageProvider, registerProvider } from './registry';
import { DocumentRow, FileContent } from '../types/Documents';
class SharePointProvider implements StorageProvider {
  readonly id = 'SharePoint' as const;
  // TODO(SPIKE-S2): embedding stays disabled until tenant CSP and framing are verified; open is always available.
  readonly capabilities = {
    canUpload: false,
    canAddLink: true,
    canDownload: true,
    canOpen: true,
    canPreviewInline: false,
    canDelete: true,
    ownsBinary: false,
    supportsBlockTransfer: false,
  };
  async getContent(row: DocumentRow): Promise<FileContent> {
    const url = new URL(row.storageRef);
    if (url.protocol !== 'https:') throw new Error('link.invalid');
    return { kind: 'url', url: url.href };
  }
}
registerProvider('SharePoint', () => new SharePointProvider());
