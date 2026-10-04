import request from 'supertest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { randomUUID, createHash } from 'node:crypto';
import { PDFDocument } from 'pdf-lib';
import JSZip from 'jszip';
import sharp from 'sharp';
import { execFileSync } from 'node:child_process';
import { createApp } from '../src/server';
import { Store } from '../src/api/store';
import { records, formattedSuffix } from '../src/api/model';
import { generateSeed } from '../src/seed/generate';
const api = '/api/data/v9.2';
let directory: string,
  seedDirectory: string,
  store: Store,
  app: ReturnType<typeof createApp>,
  clock: number;
beforeAll(async () => {
  directory = fs.mkdtempSync(path.join(os.tmpdir(), 'dms-api-test-'));
  seedDirectory = path.join(directory, 'seed');
  await generateSeed(seedDirectory, Date.UTC(2026, 9, 4));
}, 30000);
beforeEach(() => {
  clock = Date.UTC(2026, 9, 4);
  store = new Store(path.join(directory, 'data'), seedDirectory, () => clock);
  store.reset();
  app = createApp(store);
});
afterAll(() => fs.rmSync(directory, { recursive: true, force: true }));
const baseRow = () => ({
  dms_documentid: randomUUID(),
  dms_name: 'Test.txt',
  dms_regardingid: records.contoso,
  dms_regardingtype: 'account',
  dms_provider: 100000000,
});
const target = (id = randomUUID()) => ({
  annotationid: id,
  filename: 'Test.txt',
  mimetype: 'text/plain',
  isdocument: true,
  [`objectid_account@odata.bind`]: `/accounts(${records.contoso})`,
});
test('seed has valid Office packages, three-page PDFs, JPEG, PNG, ZIP and a two-second MP4', async () => {
  for (const name of [
    'Master Services Agreement.docx',
    'Q3 Proposal.pptx',
    'Pricing Model.xlsx',
    'Archive.zip',
  ]) {
    const zip = await JSZip.loadAsync(fs.readFileSync(path.join(seedDirectory, name)));
    expect(Object.keys(zip.files).length).toBeGreaterThan(0);
    if (name.endsWith('docx')) expect(zip.file('word/document.xml')).not.toBeNull();
    if (name.endsWith('pptx')) expect(zip.file('ppt/slides/slide1.xml')).not.toBeNull();
    if (name.endsWith('xlsx')) expect(zip.file('xl/workbook.xml')).not.toBeNull();
  }
  for (const name of ['Signed NDA.pdf', 'Insurance Certificate.pdf'])
    expect(
      (await PDFDocument.load(fs.readFileSync(path.join(seedDirectory, name)))).getPageCount(),
    ).toBe(3);
  expect((await sharp(path.join(seedDirectory, 'Site Photo.jpg')).metadata()).format).toBe('jpeg');
  expect((await sharp(path.join(seedDirectory, 'Logo.png')).metadata()).format).toBe('png');
  const duration = Number(
    execFileSync(
      process.env.DMS_FFPROBE_PATH || 'ffprobe',
      [
        '-v',
        'error',
        '-show_entries',
        'format=duration',
        '-of',
        'default=noprint_wrappers=1:nokey=1',
        path.join(seedDirectory, 'Kickoff Recording.mp4'),
      ],
      { encoding: 'utf8' },
    ),
  );
  expect(duration).toBeCloseTo(2, 0);
  expect(
    store.snapshot.documents.filter((row) => row.dms_regardingid === records.fabrikam),
  ).toHaveLength(1200);
  expect(
    store.snapshot.documents.filter((row) => row.dms_regardingid === records.contoso),
  ).toHaveLength(22);
});
test('CRUD, validation, PATCH upsert, ETags and explicit overwrite', async () => {
  const row = baseRow(),
    create = await request(app)
      .post(api + '/dms_documents')
      .send(row)
      .expect(201),
    id = row.dms_documentid;
  expect(create.headers['odata-entityid']).toContain(id);
  expect(create.body['@odata.etag']).toBe(create.headers.etag);
  await request(app)
    .patch(`${api}/dms_documents(${id})`)
    .set('If-None-Match', '*')
    .send({ dms_name: 'Duplicate' })
    .expect(412);
  const updated = await request(app)
    .patch(`${api}/dms_documents(${id})`)
    .set('If-Match', create.headers.etag)
    .send({ dms_name: 'Renamed.txt' })
    .expect(200);
  expect(updated.headers.etag).not.toBe(create.headers.etag);
  await request(app)
    .patch(`${api}/dms_documents(${id})`)
    .set('If-Match', create.headers.etag)
    .send({ dms_name: 'Stale.txt' })
    .expect(412);
  await request(app)
    .delete(`${api}/dms_documents(${id})`)
    .set('If-Match', create.headers.etag)
    .expect(412);
  await request(app)
    .patch(`${api}/dms_documents(${id})`)
    .set('If-Match', '*')
    .send({ dms_name: 'Explicit.txt' })
    .expect(200);
  await request(app).delete(`${api}/dms_documents(${id})`).set('If-Match', '*').expect(204);
  await request(app).get(`${api}/dms_documents(${id})`).expect(404);
  const other = baseRow();
  await request(app)
    .patch(`${api}/dms_documents(${other.dms_documentid})`)
    .set('If-None-Match', '*')
    .send(other)
    .expect(201);
  await request(app)
    .patch(`${api}/dms_documents(${randomUUID()})`)
    .set('If-Match', '*')
    .send(other)
    .expect(404);
  await request(app)
    .post(api + '/dms_documents')
    .send({ ...baseRow(), dms_name: '' })
    .expect(400);
  await request(app)
    .post(api + '/dms_documents')
    .send({ ...baseRow(), dms_documentstatus: 99 })
    .expect(400);
  await request(app)
    .post(api + '/dms_documents')
    .send({ ...baseRow(), dms_regardingid: '../../bad' })
    .expect(400);
});
test('OData select/filter/contains/date comparison/order/count/top/nextLink and formatted values', async () => {
  const response = await request(app)
    .get(api + '/dms_documents')
    .set('Prefer', 'odata.maxpagesize=7')
    .query({
      $filter: `dms_regardingid eq '${records.fabrikam}' and (contains(dms_name,'Pricing') or dms_documentstatus eq 100000001) and dms_expirydate ge '2026-09-01'`,
      $orderby: 'dms_name desc',
      $select: 'dms_documentid,dms_name,dms_documentstatus,_modifiedby_value,modifiedon',
      $count: 'true',
      $top: 19,
    })
    .expect(200);
  expect(response.body.value).toHaveLength(7);
  expect(response.body['@odata.count']).toBeGreaterThan(19);
  expect(response.body.value[0]).not.toHaveProperty('dms_regardingid');
  expect(response.body.value[0]['dms_documentstatus' + formattedSuffix]).toBeTruthy();
  expect(response.body.value[0]['_modifiedby_value' + formattedSuffix]).toBeTruthy();
  expect(response.body.value[0]['modifiedon' + formattedSuffix]).toBeTruthy();
  const all = [...response.body.value];
  let link = response.body['@odata.nextLink'];
  while (link) {
    const next = await request(app)
      .get(new URL(link).pathname + new URL(link).search)
      .set('Prefer', 'odata.maxpagesize=7')
      .expect(200);
    all.push(...next.body.value);
    link = next.body['@odata.nextLink'];
  }
  expect(all).toHaveLength(19);
  expect(new Set(all.map((row) => row.dms_documentid)).size).toBe(19);
  expect(all.map((row) => row.dms_name)).toEqual(
    all.map((row) => row.dms_name).sort((a, b) => b.localeCompare(a)),
  );
  for (const filter of [
    "dms_name eq 'a' trailing",
    'contains(dms_name)',
    "dms_name eq 'unterminated",
  ])
    await request(app)
      .get(api + '/dms_documents')
      .query({ $filter: filter })
      .expect(400);
  await request(app)
    .get(api + '/dms_documents')
    .query({ $skip: -1 })
    .expect(400);
  const empty = await request(app)
    .get(api + '/dms_documents')
    .query({ $filter: `dms_regardingid eq '${records.empty}'`, $count: 'true' })
    .expect(200);
  expect(empty.body.value).toEqual([]);
  expect(empty.body['@odata.count']).toBe(0);
});
test('small upload/download byte equality, persistence, limits and safe Note-only cascade', async () => {
  const data = Buffer.from('Hello world! \u00a9 \n\u0000\u0001'),
    annotation = target(),
    upload = await request(app)
      .post(api + '/annotations')
      .send({ ...annotation, documentbody: data.toString('base64') })
      .expect(201);
  const download = await request(app)
    .get(`${api}/annotations(${annotation.annotationid})`)
    .query({ $select: 'documentbody,filename,mimetype' })
    .expect(200);
  expect(Buffer.from(download.body.documentbody, 'base64')).toEqual(data);
  const persisted = new Store(store.directory, seedDirectory);
  expect(persisted.entity('annotations', annotation.annotationid).filesize).toBe(data.length);
  expect(fs.readFileSync(persisted.file(annotation.annotationid))).toEqual(data);
  await request(app)
    .post(api + '/annotations')
    .send({ ...annotation, documentbody: data.toString('base64') })
    .expect(412);
  await request(app)
    .delete(`${api}/annotations(${annotation.annotationid})`)
    .set('If-Match', 'W/"0"')
    .expect(412);
  const row = { ...baseRow(), dms_storageref: annotation.annotationid };
  await request(app)
    .post(api + '/dms_documents')
    .send(row)
    .expect(201);
  await request(app).delete(`${api}/dms_documents(${row.dms_documentid})`).expect(204);
  await request(app).get(`${api}/annotations(${annotation.annotationid})`).expect(404);
  expect(fs.existsSync(store.file(annotation.annotationid))).toBe(false);
  const reference = target();
  await request(app)
    .post(api + '/annotations')
    .send({ ...reference, documentbody: data.toString('base64') })
    .expect(201);
  const link = { ...baseRow(), dms_provider: 100000001, dms_storageref: reference.annotationid };
  await request(app)
    .post(api + '/dms_documents')
    .send(link)
    .expect(201);
  await request(app).delete(`${api}/dms_documents(${link.dms_documentid})`).expect(204);
  expect(fs.existsSync(store.file(reference.annotationid))).toBe(true);
  await request(app).post('/__mock/config').send({ maxUploadFileSizeBytes: 2 }).expect(200);
  await request(app)
    .post(api + '/annotations')
    .send({ ...target(), documentbody: data.toString('base64') })
    .expect(413);
  await request(app)
    .post(api + '/annotations')
    .send({ ...target(), documentbody: 'not base64!' })
    .expect(400);
  expect(upload.headers.etag).toBeTruthy();
});
test('30 MiB block transfer survives injected 429, resumes idempotently and downloads byte-for-byte', async () => {
  const data = Buffer.alloc(30 * 1024 * 1024);
  for (let i = 0; i < data.length; i += 4096) data.writeUInt32LE(i, i);
  const annotation = target();
  const init = await request(app)
      .post(api + '/InitializeAnnotationBlocksUpload')
      .send({ Target: annotation })
      .expect(200),
    token = init.body.FileContinuationToken,
    ids: string[] = [];
  await request(app)
    .post('/__mock/config')
    .send({ failNextUploadAtBlock: 2, failWith: 429 })
    .expect(200);
  for (let offset = 0; offset < data.length; offset += 4 * 1024 * 1024) {
    const id = Buffer.from(String(ids.length).padStart(8, '0')).toString('base64');
    ids.push(id);
    const body = {
      FileContinuationToken: token,
      BlockId: id,
      BlockData: data.subarray(offset, offset + 4 * 1024 * 1024).toString('base64'),
    };
    if (ids.length === 2) {
      const failed = await request(app)
        .post(api + '/UploadBlock')
        .send(body)
        .expect(429);
      expect(failed.headers['retry-after']).toBe('1');
    }
    await request(app)
      .post(api + '/UploadBlock')
      .send(body)
      .expect(204);
    if (ids.length === 1)
      await request(app)
        .post(api + '/UploadBlock')
        .send(body)
        .expect(204);
  }
  await request(app)
    .post(api + '/CommitAnnotationBlocksUpload')
    .send({ FileContinuationToken: token, BlockList: ['missing'] })
    .expect(400);
  const commit = await request(app)
    .post(api + '/CommitAnnotationBlocksUpload')
    .send({ FileContinuationToken: token, BlockList: ids })
    .expect(200);
  expect(commit.body.FileSizeInBytes).toBe(data.length);
  await request(app)
    .post(api + '/CommitAnnotationBlocksUpload')
    .send({ FileContinuationToken: token, BlockList: ids })
    .expect(200);
  const download = await request(app)
    .post(api + '/InitializeAnnotationBlocksDownload')
    .send({ Target: { annotationid: annotation.annotationid } })
    .expect(200);
  expect(download.body.FileSizeInBytes).toBe(data.length);
  const parts: Buffer[] = [];
  for (let offset = 0; offset < data.length; offset += 4 * 1024 * 1024) {
    const block = await request(app)
      .post(api + '/DownloadBlock')
      .send({
        FileContinuationToken: download.body.FileContinuationToken,
        Offset: offset,
        BlockLength: 4 * 1024 * 1024,
      })
      .expect(200);
    parts.push(Buffer.from(block.body.Data, 'base64'));
  }
  const result = Buffer.concat(parts);
  expect(createHash('sha256').update(result).digest('hex')).toBe(
    createHash('sha256').update(data).digest('hex'),
  );
  expect(result.equals(data)).toBe(true);
  await request(app)
    .post(api + '/DownloadBlock')
    .send({ FileContinuationToken: 'bad', Offset: 0, BlockLength: 1 })
    .expect(404);
  await request(app)
    .post(api + '/DownloadBlock')
    .send({
      FileContinuationToken: download.body.FileContinuationToken,
      Offset: -1,
      BlockLength: 1,
    })
    .expect(400);
}, 30000);
test('block assembly uses supplied order and enforces total size limit', async () => {
  const annotation = target(),
    init = await request(app)
      .post(api + '/InitializeAnnotationBlocksUpload')
      .send({ Target: annotation })
      .expect(200),
    token = init.body.FileContinuationToken;
  const ids = ['00000001', '00000002'].map((s) => Buffer.from(s).toString('base64'));
  for (let i = 0; i < 2; i++)
    await request(app)
      .post(api + '/UploadBlock')
      .send({
        FileContinuationToken: token,
        BlockId: ids[i],
        BlockData: Buffer.from(i === 0 ? 'first' : 'second').toString('base64'),
      })
      .expect(204);
  await request(app).post('/__mock/config').send({ maxUploadFileSizeBytes: 10 }).expect(200);
  await request(app)
    .post(api + '/CommitAnnotationBlocksUpload')
    .send({ FileContinuationToken: token, BlockList: ids })
    .expect(413);
  await request(app)
    .post('/__mock/config')
    .send({ maxUploadFileSizeBytes: 128 * 1024 * 1024 })
    .expect(200);
  await request(app)
    .post(api + '/CommitAnnotationBlocksUpload')
    .send({ FileContinuationToken: token, BlockList: ids.reverse() })
    .expect(200);
  expect(fs.readFileSync(store.file(annotation.annotationid)).toString()).toBe('secondfirst');
});
test('migration progresses in batches over 15 seconds and is idempotent', async () => {
  const body = { RegardingId: records.legacy, EntityLogicalName: 'opportunity' };
  const initial = await request(app)
    .post(api + '/dms_EnsureMigration')
    .send(body)
    .expect(200);
  expect(initial.body.state).toBe('Running');
  expect(initial.body.total).toBe(40);
  clock += 5000;
  const partial = await request(app)
    .post(api + '/dms_GetMigrationStatus')
    .send(body)
    .expect(200);
  expect(partial.body.processed).toBe(13);
  await request(app)
    .post(api + '/dms_EnsureMigration')
    .send(body)
    .expect(200);
  expect(
    store.snapshot.documents.filter((row) => row.dms_regardingid === records.legacy),
  ).toHaveLength(13);
  clock += 10000;
  const completed = await request(app)
    .post(api + '/dms_GetMigrationStatus')
    .send(body)
    .expect(200);
  expect(completed.body.state).toBe('Completed');
  expect(completed.body.processed).toBe(40);
  await request(app)
    .post(api + '/dms_EnsureMigration')
    .send(body)
    .expect(200);
  expect(
    store.snapshot.documents.filter((row) => row.dms_regardingid === records.legacy),
  ).toHaveLength(40);
});
test('configuration, metadata, faults, framing and reset are available without external calls', async () => {
  await request(app)
    .post('/__mock/config')
    .send({ activeProvider: 'SharePoint', spBlockFraming: true })
    .expect(200);
  const config = await request(app)
    .post(api + '/dms_GetClientConfig')
    .send({ EntityLogicalName: 'account', RegardingId: records.contoso })
    .expect(200);
  expect(JSON.parse(config.body.ConfigJson).activeProvider).toBe('SharePoint');
  const metadata = await request(app)
    .get(api + "/EntityDefinitions(LogicalName='dms_document')")
    .query({ $expand: 'Attributes' })
    .expect(200);
  expect(
    metadata.body.Attributes.map((field: { AttributeType: string }) => field.AttributeType),
  ).toEqual(
    expect.arrayContaining([
      'String',
      'Memo',
      'Picklist',
      'MultiSelectPicklist',
      'DateTime',
      'Boolean',
      'Money',
      'Integer',
      'Decimal',
      'Lookup',
    ]),
  );
  const entity = await request(app)
    .get(api + "/EntityDefinitions(LogicalName='account')")
    .query({ $select: 'EntitySetName' })
    .expect(200);
  expect(entity.body.EntitySetName).toBe('accounts');
  const throttled = await request(app)
    .get(api + '/dms_documents')
    .set('x-mock-fail', '429')
    .expect(429);
  expect(throttled.body.error.code).toBe('InjectedFailure');
  expect(throttled.headers['retry-after']).toBe('1');
  expect(throttled.headers['x-ms-service-request-id']).toBeTruthy();
  await request(app)
    .get(api + '/dms_documents')
    .set('x-mock-fail', '500')
    .expect(500);
  await expect(
    request(app)
      .get(api + '/dms_documents')
      .set('x-mock-fail', 'offline'),
  ).rejects.toThrow();
  const frame = await request(app).get('/mock-sharepoint/view/sample').expect(200);
  expect(frame.headers['x-frame-options']).toBe('DENY');
  const row = baseRow();
  await request(app)
    .post(api + '/dms_documents')
    .send(row)
    .expect(201);
  await request(app).post('/__mock/reset').expect(200);
  await request(app).get(`${api}/dms_documents(${row.dms_documentid})`).expect(404);
  expect(store.config.activeProvider).toBe('Note');
  expect(store.snapshot.documents).toHaveLength(1222);
  expect(store.uploads.size).toBe(0);
});
