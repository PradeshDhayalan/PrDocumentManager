import request from 'supertest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { createApp } from '../src/server';
import { Store } from '../src/api/store';

test('custom column metadata and typed document values survive a store restart', async () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'dms-columns-'));
  try {
    const store = new Store(directory, path.join(directory, 'seed'));
    const app = createApp(store);
    const column = { name: 'dms_projectcode', label: 'Project code', type: 'Integer' };
    await request(app).post('/__mock/columns').send(column).expect(201);
    await request(app).post('/__mock/columns').send(column).expect(409);
    await request(app)
      .post('/__mock/columns')
      .send({ ...column, name: 'dms_documentid' })
      .expect(409);
    await request(app)
      .post('/__mock/columns')
      .send({ ...column, name: 'dms_x) or true' })
      .expect(400);
    const id = randomUUID();
    const document = {
      dms_documentid: id,
      dms_name: 'Typed column.txt',
      dms_regardingid: randomUUID(),
      dms_regardingtype: 'account',
      dms_provider: 100000000,
      dms_projectcode: 42,
    };
    store.saveDocument(id, document, {});
    await request(app)
      .patch(`/api/data/v9.2/dms_documents(${id})`)
      .send({ dms_projectcode: 'invalid' })
      .expect(400);
    const reloaded = new Store(directory, path.join(directory, 'seed'));
    expect(reloaded.entity('documents', id).dms_projectcode).toBe(42);
    const response = await request(createApp(reloaded))
      .get("/api/data/v9.2/EntityDefinitions(LogicalName='dms_document')?$expand=Attributes")
      .expect(200);
    expect(response.body.Attributes).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ LogicalName: column.name, IsValidForUpdate: true }),
      ]),
    );
    await request(createApp(reloaded)).delete(`/__mock/columns/${column.name}`).expect(204);
    expect(reloaded.entity('documents', id).dms_projectcode).toBeUndefined();
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});
