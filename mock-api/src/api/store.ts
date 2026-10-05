import fs from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import {
  ApiError,
  Entity,
  Migration,
  MockConfig,
  Snapshot,
  defaultConfig,
  requireGuid,
  userIds,
  validateDocument,
} from './model';
interface UploadSession {
  target: Entity;
  blocks: Record<string, string>;
  received: number;
  committed?: number;
  committedBlocks?: string[];
}
interface DownloadSession {
  annotationId: string;
}
export class Store {
  public snapshot: Snapshot;
  public config: MockConfig = { ...defaultConfig };
  public uploads = new Map<string, UploadSession>();
  public downloads = new Map<string, DownloadSession>();
  public generation = 0;
  constructor(
    public directory: string,
    public seedDirectory: string,
    public now: () => number = Date.now,
  ) {
    fs.mkdirSync(this.directory, { recursive: true });
    fs.mkdirSync(this.binaryDirectory, { recursive: true });
    this.snapshot = fs.existsSync(this.statePath)
      ? (JSON.parse(fs.readFileSync(this.statePath, 'utf8')) as Snapshot)
      : { documents: [], annotations: [], migrations: {} };
  }
  get statePath(): string {
    return path.join(this.directory, 'state.json');
  }
  get binaryDirectory(): string {
    return path.join(this.directory, 'uploads');
  }
  file(id: string): string {
    return path.join(this.binaryDirectory, requireGuid(id));
  }
  persist(): void {
    const temporary = this.statePath + '.tmp';
    fs.writeFileSync(temporary, JSON.stringify(this.snapshot));
    fs.renameSync(temporary, this.statePath);
  }
  reset(): void {
    const seedPath = path.join(this.seedDirectory, 'state.json');
    if (!fs.existsSync(seedPath)) throw new ApiError(503, 'SeedMissing', 'Run npm run seed first.');
    this.snapshot = JSON.parse(fs.readFileSync(seedPath, 'utf8')) as Snapshot;
    fs.rmSync(this.binaryDirectory, { recursive: true, force: true });
    fs.mkdirSync(this.binaryDirectory, { recursive: true });
    fs.rmSync(path.join(this.directory, 'blocks'), { recursive: true, force: true });
    this.uploads.clear();
    this.downloads.clear();
    this.config = { ...defaultConfig };
    this.generation++;
    for (const [id, name] of Object.entries(this.snapshot.binarySeeds || {}))
      fs.copyFileSync(path.join(this.seedDirectory, name), this.file(id));
    this.persist();
  }
  entity(collection: 'documents' | 'annotations', id: string): Entity {
    const entity = this.snapshot[collection].find(
      (row) =>
        row[collection === 'documents' ? 'dms_documentid' : 'annotationid'] === requireGuid(id),
    );
    if (!entity) throw new ApiError(404, 'NotFound', 'Entity was not found.');
    return entity;
  }
  checkMatch(row: Entity, match: string | undefined): void {
    if (match && match !== '*' && match !== row['@odata.etag'])
      throw new ApiError(
        412,
        'ConcurrencyVersionMismatch',
        'The record has changed. Reload or explicitly overwrite.',
      );
  }
  bump(row: Entity): void {
    row['@odata.etag'] = `W/"${Number(/\d+/.exec(String(row['@odata.etag']))?.[0] || 0) + 1}"`;
    row.modifiedon = new Date(this.now()).toISOString();
    row._modifiedby_value = userIds[0];
  }
  saveDocument(
    id: string,
    attributes: Entity,
    headers: { match?: string; noneMatch?: string },
  ): { row: Entity; created: boolean } {
    id = requireGuid(id);
    const existing = this.snapshot.documents.find((row) => row.dms_documentid === id);
    if (existing && headers.noneMatch === '*')
      throw new ApiError(412, 'EntityAlreadyExists', 'Entity already exists.');
    if (!existing && headers.match) throw new ApiError(404, 'NotFound', 'Entity was not found.');
    if (existing) this.checkMatch(existing, headers.match);
    const row: Entity = {
      dms_uploadstate: 100000000,
      dms_documentstatus: 100000000,
      createdon: new Date(this.now()).toISOString(),
      _createdby_value: userIds[0],
      _ownerid_value: userIds[0],
      ...existing,
      ...attributes,
      dms_documentid: id,
    };
    validateDocument(row);
    this.bump(row);
    if (existing) Object.assign(existing, row);
    else this.snapshot.documents.push(row);
    this.persist();
    return { row, created: !existing };
  }
  deleteDocument(id: string, match?: string): void {
    const row = this.entity('documents', id);
    this.checkMatch(row, match);
    // TODO(SPIKE-S6): real server pre-delete plugin must enforce annotation ownership.
    if (row.dms_provider === 100000000 && row.dms_storageref) {
      const otherOwner = this.snapshot.documents.some(
        (other) => other !== row && other.dms_storageref === row.dms_storageref,
      );
      if (
        !otherOwner &&
        this.snapshot.annotations.some(
          (note) =>
            note.annotationid === row.dms_storageref &&
            note._objectid_value === row.dms_regardingid,
        )
      )
        this.deleteAnnotation(String(row.dms_storageref));
    }
    this.snapshot.documents = this.snapshot.documents.filter((other) => other !== row);
    this.persist();
  }
  cleanupUpload(id: string): void {
    const row = this.entity('documents', id);
    if (row.dms_provider !== 100000000)
      throw new ApiError(400, 'InvalidProvider', 'Only owned Note uploads can be cleaned up.');
    if (row.dms_uploadstate === 100000001)
      throw new ApiError(409, 'AlreadyAvailable', 'A completed document cannot be compensated.');
    const annotationId = String(row.dms_storageref || '');
    if (annotationId) {
      const note = this.snapshot.annotations.find((note) => note.annotationid === annotationId);
      if (
        note &&
        note._objectid_value === row.dms_regardingid &&
        !this.snapshot.documents.some(
          (other) => other !== row && other.dms_storageref === annotationId,
        )
      )
        this.deleteAnnotation(annotationId);
      for (const [token, session] of this.uploads)
        if (
          session.target.annotationid === annotationId &&
          session.target._objectid_value === row.dms_regardingid
        ) {
          fs.rmSync(path.join(this.directory, 'blocks', token), { recursive: true, force: true });
          this.uploads.delete(token);
        }
    }
    row.dms_uploadstate = 100000002;
    this.bump(row);
    this.persist();
  }
  deleteAnnotation(id: string, match?: string): void {
    const row = this.entity('annotations', id);
    this.checkMatch(row, match);
    this.snapshot.annotations = this.snapshot.annotations.filter((other) => other !== row);
    fs.rmSync(this.file(id), { force: true });
    this.persist();
  }
  annotationTarget(attributes: Entity): Entity {
    const id = requireGuid(attributes.annotationid || randomUUID()),
      bind = Object.entries(attributes).find(([key]) =>
        /^objectid_[a-z0-9_]+@odata.bind$/.test(key),
      );
    const parent = bind && /^\/[a-z0-9_]+\(([a-f0-9-]+)\)$/i.exec(String(bind[1]));
    if (!parent)
      throw new ApiError(400, 'InvalidParent', 'A polymorphic parent binding is required.');
    requireGuid(parent[1]);
    if (
      typeof attributes.filename !== 'string' ||
      !attributes.filename ||
      attributes.filename.length > 255
    )
      throw new ApiError(400, 'InvalidFilename', 'A filename of up to 255 characters is required.');
    if (this.snapshot.annotations.some((row) => row.annotationid === id))
      throw new ApiError(412, 'EntityAlreadyExists', 'Annotation already exists.');
    const { documentbody: _documentBody, ...metadata } = attributes;
    return {
      ...metadata,
      annotationid: id,
      isdocument: true,
      _objectid_value: parent[1],
      createdon: new Date(this.now()).toISOString(),
      _createdby_value: userIds[0],
      _ownerid_value: userIds[0],
    };
  }
  sizeLimit(size: number): void {
    if (size > this.config.maxUploadFileSizeBytes)
      throw new ApiError(
        413,
        'AttachmentTooLarge',
        'The attachment exceeds the organisation size limit.',
      );
  }
  saveAnnotation(target: Entity, size: number): Entity {
    this.sizeLimit(size);
    const row = { ...target, filesize: size };
    this.bump(row);
    this.snapshot.annotations.push(row);
    this.persist();
    return row;
  }
  singleUpload(attributes: Entity): Entity {
    const target = this.annotationTarget(attributes),
      data = decodeBase64(attributes.documentbody);
    this.sizeLimit(data.length);
    fs.writeFileSync(this.file(String(target.annotationid)), data);
    return this.saveAnnotation(target, data.length);
  }
  initUpload(target: Entity): string {
    const token = randomUUID();
    this.uploads.set(token, { target: this.annotationTarget(target), blocks: {}, received: 0 });
    return token;
  }
  uploadBlock(token: string, blockId: string, data: Buffer): void {
    const session = this.uploads.get(token);
    if (!session) throw new ApiError(404, 'InvalidContinuationToken', 'Unknown upload token.');
    if (session.committed !== undefined)
      throw new ApiError(409, 'UploadCommitted', 'Upload is already committed.');
    if (!blockId || blockId.length > 128 || !/^[A-Za-z0-9+/=]+$/.test(blockId))
      throw new ApiError(400, 'InvalidBlockId', 'A base64 block identifier is required.');
    if (data.length > 4 * 1024 * 1024)
      throw new ApiError(413, 'BlockTooLarge', 'Blocks may not exceed 4 MiB.');
    if (!session.blocks[blockId]) session.received++;
    if (this.config.failNextUploadAtBlock === session.received) {
      this.config.failNextUploadAtBlock = null;
      throw new ApiError(
        typeof this.config.failWith === 'number' ? this.config.failWith : 500,
        'InjectedBlockFailure',
        'Injected block failure. Retry this block with the same token.',
      );
    }
    const total = Object.entries(session.blocks).reduce(
      (sum, [id, file]) => sum + (id === blockId ? 0 : fs.statSync(file).size),
      data.length,
    );
    this.sizeLimit(total);
    const folder = path.join(this.directory, 'blocks', token);
    fs.mkdirSync(folder, { recursive: true });
    const file = session.blocks[blockId] || path.join(folder, randomUUID());
    fs.writeFileSync(file, data);
    session.blocks[blockId] = file;
  }
  commit(token: string, blockList: string[]): number {
    const session = this.uploads.get(token);
    if (!session) throw new ApiError(404, 'InvalidContinuationToken', 'Unknown upload token.');
    if (session.committed !== undefined) {
      if (JSON.stringify(blockList) !== JSON.stringify(session.committedBlocks))
        throw new ApiError(
          409,
          'UploadCommitted',
          'A committed upload cannot change its ordered block list.',
        );
      return session.committed;
    }
    if (
      !blockList.length ||
      new Set(blockList).size !== blockList.length ||
      blockList.some((id) => !session.blocks[id])
    )
      throw new ApiError(
        400,
        'MissingBlock',
        'The ordered block list is incomplete or duplicated.',
      );
    const size = blockList.reduce((sum, id) => sum + fs.statSync(session.blocks[id]).size, 0);
    this.sizeLimit(size);
    const id = String(session.target.annotationid);
    if (this.snapshot.annotations.some((row) => row.annotationid === id))
      throw new ApiError(412, 'EntityAlreadyExists', 'Annotation already exists.');
    const destination = this.file(id),
      descriptor = fs.openSync(destination, 'w');
    try {
      for (const blockId of blockList) {
        const block = fs.readFileSync(session.blocks[blockId]);
        let written = 0;
        while (written < block.length)
          written += fs.writeSync(descriptor, block, written, block.length - written);
      }
    } finally {
      fs.closeSync(descriptor);
    }
    this.saveAnnotation(session.target, size);
    session.committed = size;
    session.committedBlocks = [...blockList];
    fs.rmSync(path.join(this.directory, 'blocks', token), { recursive: true, force: true });
    return size;
  }
  initDownload(id: string): { token: string; row: Entity } {
    const row = this.entity('annotations', id);
    if (!fs.existsSync(this.file(id)))
      throw new ApiError(404, 'BinaryMissing', 'Attachment binary was not found.');
    const token = randomUUID();
    this.downloads.set(token, { annotationId: id });
    return { token, row };
  }
  downloadBlock(token: string, offset: number, length: number): Buffer {
    const session = this.downloads.get(token);
    if (!session) throw new ApiError(404, 'InvalidContinuationToken', 'Unknown download token.');
    if (
      !Number.isSafeInteger(offset) ||
      offset < 0 ||
      !Number.isSafeInteger(length) ||
      length < 1 ||
      length > 4 * 1024 * 1024
    )
      throw new ApiError(
        400,
        'InvalidRange',
        'A non-negative offset and block length up to 4 MiB are required.',
      );
    const file = this.file(session.annotationId);
    if (!fs.existsSync(file))
      throw new ApiError(404, 'BinaryMissing', 'Attachment binary was not found.');
    const size = fs.statSync(file).size;
    if (offset > size) throw new ApiError(416, 'InvalidRange', 'Offset exceeds file length.');
    const buffer = Buffer.alloc(Math.min(length, size - offset)),
      descriptor = fs.openSync(file, 'r');
    try {
      fs.readSync(descriptor, buffer, 0, buffer.length, offset);
    } finally {
      fs.closeSync(descriptor);
    }
    return buffer;
  }
  migration(id: string, ensure = false): Migration {
    id = requireGuid(id);
    let state = this.snapshot.migrations[id];
    if (!state) {
      state = {
        state: 'NotStarted',
        processed: 0,
        total: this.snapshot.annotations.filter(
          (row) =>
            row._objectid_value === id &&
            !this.snapshot.documents.some((doc) => doc.dms_storageref === row.annotationid),
        ).length,
      };
      this.snapshot.migrations[id] = state;
    }
    if (ensure && state.state === 'NotStarted') {
      state.startedAt = this.now();
      state.state = state.total ? 'Running' : 'Completed';
    }
    if (state.state === 'Running') {
      const target = Math.min(
        state.total,
        Math.floor(((this.now() - (state.startedAt ?? this.now())) / 15000) * state.total),
      );
      const legacy = this.snapshot.annotations.filter(
        (row) =>
          row._objectid_value === id &&
          !this.snapshot.documents.some((doc) => doc.dms_storageref === row.annotationid),
      );
      for (const note of legacy.slice(0, Math.max(0, target - state.processed))) {
        this.saveDocument(
          randomUUID(),
          {
            dms_name: note.filename,
            dms_originalfilename: note.filename,
            dms_regardingid: id,
            dms_regardingtype: 'opportunity',
            dms_provider: 100000000,
            dms_storageref: note.annotationid,
            dms_contenttype: note.mimetype,
            dms_filesizekb: Math.ceil(Number(note.filesize) / 1024),
            dms_uploadstate: 100000001,
            dms_documenttype: 100000005,
          },
          {},
        );
        state.processed++;
      }
      if (state.processed === state.total) state.state = 'Completed';
    }
    this.persist();
    return state;
  }
}
export function decodeBase64(value: unknown): Buffer {
  if (typeof value !== 'string' || value.length % 4 !== 0 || /[^A-Za-z0-9+/=]/.test(value))
    throw new ApiError(400, 'InvalidBase64', 'Valid base64 data is required.');
  const data = Buffer.from(value, 'base64');
  if (data.toString('base64') !== value)
    throw new ApiError(400, 'InvalidBase64', 'Canonical base64 data is required.');
  return data;
}
