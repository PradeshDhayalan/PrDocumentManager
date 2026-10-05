import { sha256 } from '@noble/hashes/sha256';
import { DocumentRepository, providerChoices, quote } from './DocumentRepository';
import { ClientConfig, DocumentRow, RawEntity } from '../types/Documents';
import { resolveProvider } from '../providers/registry';
import { sanitizeFilename, newGuid, extension } from './format';
import { aborted } from './DataverseClient';
export type UploadState =
  | 'queued'
  | 'hashing'
  | 'duplicate'
  | 'uploading'
  | 'complete'
  | 'failed'
  | 'cancelled';
export interface UploadJob {
  id: string;
  file: File;
  state: UploadState;
  progress: number;
  row?: DocumentRow;
  controller?: AbortController;
  error?: unknown;
  checksum?: string;
  force?: boolean;
}
export class UploadQueue {
  private jobs: UploadJob[] = [];
  private listeners = new Set<() => void>();
  private running = 0;
  private disposed = false;
  constructor(
    private repo: DocumentRepository,
    private config: ClientConfig,
    private recordId: string,
    private entityName: string,
    private onRow: (row: DocumentRow) => void,
  ) {}
  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };
  snapshot = (): UploadJob[] => this.jobs.map((job) => ({ ...job }));
  private emitRow(row: DocumentRow): void {
    if (!this.disposed) this.onRow(row);
  }
  private notify(): void {
    this.listeners.forEach((listener) => listener());
  }
  add(files: File[]): string[] {
    const errors: string[] = [];
    for (const file of files) {
      const ext = extension(file.name),
        max = Math.min(
          this.config.entity.maxFileSizeMb * 1048576,
          this.config.org.maxUploadFileSizeBytes,
        );
      let error = '';
      if (!file.size) error = 'upload.zero';
      else if (file.size > max) error = 'upload.tooLarge';
      else if (
        this.config.entity.blockedExtensions.includes(ext) ||
        (this.config.entity.allowedExtensions.length &&
          !this.config.entity.allowedExtensions.includes(ext))
      )
        error = 'upload.extension';
      if (error) {
        errors.push(error);
        continue;
      }
      this.jobs.push({ id: newGuid(), file, state: 'queued', progress: 0 });
    }
    this.notify();
    this.drain();
    return errors;
  }
  retry(id: string, force = false): void {
    const job = this.jobs.find((job) => job.id === id);
    if (!job || !['failed', 'cancelled', 'duplicate'].includes(job.state)) return;
    job.force = force || job.force;
    job.state = 'queued';
    job.error = undefined;
    job.progress = 0;
    this.notify();
    this.drain();
  }
  cancel(id: string): void {
    const job = this.jobs.find((job) => job.id === id);
    if (!job) return;
    if (['queued', 'duplicate'].includes(job.state)) {
      job.state = 'cancelled';
      this.notify();
    } else job.controller?.abort();
  }
  removeCompleted(): void {
    this.jobs = this.jobs.filter((job) => !['complete', 'cancelled'].includes(job.state));
    this.notify();
  }
  dispose(): void {
    this.disposed = true;
    this.jobs.forEach((job) => job.controller?.abort());
    this.listeners.clear();
  }
  private drain(): void {
    if (this.disposed) return;
    while (this.running < 3) {
      const job = this.jobs.find((job) => job.state === 'queued');
      if (!job) break;
      this.running++;
      job.state = 'hashing';
      this.notify();
      void this.run(job).finally(() => {
        this.running--;
        this.drain();
      });
    }
  }
  private async run(job: UploadJob): Promise<void> {
    const controller = new AbortController();
    job.controller = controller;
    try {
      if (!job.checksum) {
        const hash = sha256.create();
        for (let offset = 0; offset < job.file.size; offset += 4 * 1048576) {
          if (controller.signal.aborted) throw new DOMException('Aborted', 'AbortError');
          hash.update(
            new Uint8Array(await job.file.slice(offset, offset + 4 * 1048576).arrayBuffer()),
          );
        }
        job.checksum = Array.from(hash.digest())
          .map((byte) => byte.toString(16).padStart(2, '0'))
          .join('');
      }
      if (!job.force && !job.row) {
        const filter = `dms_regardingid eq ${quote(this.recordId)} and dms_regardingtype eq ${quote(this.entityName)} and dms_checksum eq ${quote(job.checksum)}`;
        const duplicates = await this.repo.client.getPage<RawEntity>(
          `dms_documents?$filter=${encodeURIComponent(filter)}&$top=1`,
          controller.signal,
        );
        if (duplicates.value.length) {
          job.state = 'duplicate';
          this.notify();
          return;
        }
      }
      const safeFile = new File([job.file], sanitizeFilename(job.file.name), {
        type: job.file.type,
      });
      if (job.row) {
        await this.repo.client.action('dms_CleanupFailedUpload', { DocumentId: job.row.id });
        job.row = await this.repo.update(await this.repo.get(job.row.id), {
          dms_uploadstate: 100000000,
        });
      } else
        job.row = await this.repo.create(
          {
            dms_documentid: job.id,
            dms_name: safeFile.name,
            dms_originalfilename: job.file.name.slice(0, 255),
            dms_regardingid: this.recordId,
            dms_regardingtype: this.entityName,
            dms_provider: providerChoices.Note,
            dms_storageref: newGuid(),
            dms_contenttype: safeFile.type || 'application/octet-stream',
            dms_filesizekb: Math.ceil(safeFile.size / 1024),
            dms_uploadstate: 100000000,
            dms_documenttype: 100000005,
            dms_documentstatus: 100000000,
          },
          controller.signal,
        );
      this.emitRow(job.row);
      job.state = 'uploading';
      this.notify();
      const provider = resolveProvider(job.row.provider, this.repo.client);
      await provider.upload?.(
        safeFile,
        job.row.storageRef,
        this.entityName,
        this.recordId,
        (percent) => {
          job.progress = percent;
          this.notify();
        },
        controller.signal,
      );
      job.row = await this.repo.update(
        job.row,
        { dms_uploadstate: 100000001, dms_checksum: job.checksum },
        false,
        controller.signal,
      );
      this.emitRow(job.row);
      job.state = 'complete';
      this.notify();
    } catch (error) {
      job.error = error;
      job.state = aborted(error) || controller.signal.aborted ? 'cancelled' : 'failed';
      if (job.row) {
        try {
          const latest = await this.repo.get(job.row.id);
          if (latest.uploadState === 100000001 && latest.checksum === job.checksum) {
            job.row = latest;
            job.state = 'complete';
            job.progress = 100;
            this.emitRow(latest);
          } else {
            await this.repo.client.action('dms_CleanupFailedUpload', { DocumentId: job.row.id });
            job.row = await this.repo.get(job.row.id);
            this.emitRow(job.row);
          }
        } catch {
          /* Keep Pending if server compensation cannot be confirmed. */
        }
      }
      this.notify();
    }
  }
}
