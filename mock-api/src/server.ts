import express, { Request, Response, NextFunction } from 'express';
import cors from 'cors';
import fs from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { ApiError, Entity, MockConfig, payload, records, requireGuid } from './api/model';
import { Store, decodeBase64 } from './api/store';
import { queryEntities, selectEntity } from './api/query';
import { attributes, editableColumns, entitySets, visibleColumns } from './api/metadata';
export function createApp(store: Store) {
  const app = express(),
    api = express.Router();
  app.use(
    cors({ exposedHeaders: ['ETag', 'OData-EntityId', 'Retry-After', 'x-ms-service-request-id'] }),
  );
  app.use((_req, res, next) => {
    res.set('x-ms-service-request-id', randomUUID());
    next();
  });
  app.use(express.json({ limit: '12mb' }));
  const route =
    (handler: (req: Request, res: Response) => void | Promise<void>) =>
    (req: Request, res: Response, next: NextFunction) => {
      Promise.resolve()
        .then(() => handler(req, res))
        .catch(next);
    };
  app.get('/__mock/health', (_req, res) => res.json({ ready: true, milestone: 'M1' }));
  app.get('/__mock/state', (_req, res) =>
    res.json({
      documents: store.snapshot.documents.length,
      annotations: store.snapshot.annotations.length,
      uploads: store.uploads.size,
      migrations: store.snapshot.migrations,
      config: store.config,
      generation: store.generation,
    }),
  );
  app.post(
    '/__mock/reset',
    route((_req, res) => {
      store.reset();
      res.json({ reset: true });
    }),
  );
  app.post(
    '/__mock/config',
    route((req, res) => {
      const input = payload(req.body),
        next: MockConfig = { ...store.config };
      for (const key of [
        'latencyMs',
        'failureRate',
        'maxUploadFileSizeBytes',
        'failNextUploadAtBlock',
      ] as const)
        if (key in input) {
          const value = input[key];
          if (key === 'failNextUploadAtBlock' && value === null) next[key] = null;
          else {
            if (
              typeof value !== 'number' ||
              !Number.isFinite(value) ||
              value < 0 ||
              (key === 'failureRate' && value > 1) ||
              (key !== 'failureRate' && !Number.isInteger(value))
            )
              throw new ApiError(400, 'InvalidConfig', 'Invalid numeric configuration.');
            next[key] = value;
          }
        }
      if ('activeProvider' in input) {
        if (input.activeProvider !== 'Note' && input.activeProvider !== 'SharePoint')
          throw new ApiError(400, 'InvalidConfig', 'Invalid active provider.');
        next.activeProvider = input.activeProvider;
      }
      if ('spBlockFraming' in input) {
        if (typeof input.spBlockFraming !== 'boolean')
          throw new ApiError(400, 'InvalidConfig', 'Framing setting must be boolean.');
        next.spBlockFraming = input.spBlockFraming;
      }
      if ('failWith' in input) {
        if (![null, 429, 500, 412, 'timeout', 'offline'].some((value) => value === input.failWith))
          throw new ApiError(400, 'InvalidConfig', 'Invalid failure mode.');
        next.failWith = input.failWith as MockConfig['failWith'];
      }
      if (input.migrationState) {
        if (
          !['NotStarted', 'Running', 'Completed', 'Failed'].includes(String(input.migrationState))
        )
          throw new ApiError(400, 'InvalidConfig', 'Invalid migration state.');
        const state = store.snapshot.migrations[records.legacy];
        state.state = input.migrationState as typeof state.state;
        if (state.state === 'Running') state.startedAt = store.now();
        store.persist();
      }
      store.config = next;
      res.json(next);
    }),
  );
  app.get(
    '/__mock/files/:id',
    route((req, res) => {
      const row = store.entity('annotations', req.params.id);
      res
        .type(String(row.mimetype || 'application/octet-stream'))
        .sendFile(store.file(req.params.id));
    }),
  );
  app.get('/mock-sharepoint/view/:id', (req, res) => {
    const blocked = req.query.blocked === '1' || store.config.spBlockFraming;
    res.set(
      'Content-Security-Policy',
      blocked ? "frame-ancestors 'none'" : "frame-ancestors 'self'",
    );
    if (blocked) res.set('X-Frame-Options', 'DENY');
    res
      .type('html')
      .send(
        '<!doctype html><html lang="en"><title>Mock SharePoint viewer</title><body><h1>Mock SharePoint viewer</h1><p>This local sample represents an existing SharePoint document. No SharePoint API or authentication is used.</p></body></html>',
      );
  });
  api.use((req, res, next) => {
    void (async () => {
      const latency = Number(req.get('x-mock-latency-ms') ?? store.config.latencyMs);
      if (Number.isFinite(latency) && latency > 0)
        await new Promise((resolve) => setTimeout(resolve, Math.min(latency, 30000)));
      const fault =
        req.get('x-mock-fail') ||
        (store.config.failureRate > 0 && Math.random() < store.config.failureRate
          ? store.config.failWith || 500
          : null);
      if (fault === 'offline') {
        req.socket.destroy();
        return;
      }
      if (fault === 'timeout') {
        await new Promise((resolve) => setTimeout(resolve, 25000));
        throw new ApiError(504, 'MockTimeout', 'Injected timeout.');
      }
      if (fault) {
        const status = Number(fault);
        if (![429, 500, 412].includes(status))
          throw new ApiError(400, 'InvalidFault', 'Unsupported fault override.');
        if (status === 429) res.set('Retry-After', '1');
        throw new ApiError(status, 'InjectedFailure', 'Injected request failure.');
      }
      next();
    })().catch(next);
  });
  const requestUrl = (req: Request) =>
    new URL(req.originalUrl, `${req.protocol}://${req.get('host')}`);
  const sendEntity = (res: Response, row: Entity, select: string | null = null) => {
    res.set('ETag', String(row['@odata.etag'])).json(selectEntity(row, select));
  };
  api.get(
    '/dms_documents',
    route((req, res) => {
      res.json(queryEntities(store.snapshot.documents, requestUrl(req), req.get('Prefer')));
    }),
  );
  api.post(
    '/dms_documents',
    route((req, res) => {
      const input = payload(req.body),
        id = requireGuid(input.dms_documentid || randomUUID());
      if (store.snapshot.documents.some((row) => row.dms_documentid === id))
        throw new ApiError(412, 'EntityAlreadyExists', 'Entity already exists.');
      const result = store.saveDocument(id, input, { noneMatch: req.get('If-None-Match') });
      res
        .status(201)
        .set('OData-EntityId', `${requestUrl(req).origin}/api/data/v9.2/dms_documents(${id})`);
      sendEntity(res, result.row);
    }),
  );
  api.get(
    /^\/dms_documents\(([^)]+)\)$/,
    route((req, res) => {
      sendEntity(
        res,
        store.entity('documents', req.params[0]),
        requestUrl(req).searchParams.get('$select'),
      );
    }),
  );
  api.patch(
    /^\/dms_documents\(([^)]+)\)$/,
    route((req, res) => {
      const result = store.saveDocument(req.params[0], payload(req.body), {
        match: req.get('If-Match'),
        noneMatch: req.get('If-None-Match'),
      });
      res.status(result.created ? 201 : 200);
      sendEntity(res, result.row);
    }),
  );
  api.delete(
    /^\/dms_documents\(([^)]+)\)$/,
    route((req, res) => {
      store.deleteDocument(req.params[0], req.get('If-Match'));
      res.sendStatus(204);
    }),
  );
  api.get(
    '/annotations',
    route((req, res) => {
      res.json(queryEntities(store.snapshot.annotations, requestUrl(req), req.get('Prefer')));
    }),
  );
  api.post(
    '/annotations',
    route((req, res) => {
      const row = store.singleUpload(payload(req.body));
      res
        .status(201)
        .set(
          'OData-EntityId',
          `${requestUrl(req).origin}/api/data/v9.2/annotations(${row.annotationid})`,
        );
      sendEntity(res, row);
    }),
  );
  api.get(
    /^\/annotations\(([^)]+)\)$/,
    route((req, res) => {
      const row = { ...store.entity('annotations', req.params[0]) },
        select = requestUrl(req).searchParams.get('$select');
      if (
        !select ||
        select
          .split(',')
          .map((s) => s.trim())
          .includes('documentbody')
      ) {
        const file = store.file(req.params[0]);
        if (!fs.existsSync(file))
          throw new ApiError(404, 'BinaryMissing', 'Attachment binary was not found.');
        row.documentbody = fs.readFileSync(file).toString('base64');
      }
      sendEntity(res, row, select);
    }),
  );
  api.delete(
    /^\/annotations\(([^)]+)\)$/,
    route((req, res) => {
      store.deleteAnnotation(req.params[0], req.get('If-Match'));
      res.sendStatus(204);
    }),
  );
  api.post(
    '/InitializeAnnotationBlocksUpload',
    route((req, res) => {
      const body = req.body as Record<string, unknown>;
      res.json({ FileContinuationToken: store.initUpload(payload(body.Target)) });
    }),
  );
  api.post(
    '/UploadBlock',
    route((req, res) => {
      const body = payload(req.body);
      try {
        store.uploadBlock(
          String(body.FileContinuationToken),
          String(body.BlockId || ''),
          decodeBase64(body.BlockData),
        );
      } catch (error) {
        if (error instanceof ApiError && error.status === 429) res.set('Retry-After', '1');
        throw error;
      }
      res.sendStatus(204);
    }),
  );
  api.post(
    '/CommitAnnotationBlocksUpload',
    route((req, res) => {
      const body = req.body as Record<string, unknown>;
      if (
        !Array.isArray(body.BlockList) ||
        !body.BlockList.every((value) => typeof value === 'string')
      )
        throw new ApiError(400, 'InvalidBlocks', 'An ordered block list is required.');
      res.json({
        FileSizeInBytes: store.commit(
          String(body.FileContinuationToken),
          body.BlockList as string[],
        ),
      });
    }),
  );
  api.post(
    '/InitializeAnnotationBlocksDownload',
    route((req, res) => {
      const body = req.body as Record<string, unknown>,
        target = payload(body.Target),
        { token, row } = store.initDownload(requireGuid(target.annotationid));
      res.json({
        FileContinuationToken: token,
        FileSizeInBytes: row.filesize,
        FileName: row.filename,
      });
    }),
  );
  api.post(
    '/DownloadBlock',
    route((req, res) => {
      const body = payload(req.body);
      res.json({
        Data: store
          .downloadBlock(
            String(body.FileContinuationToken),
            Number(body.Offset),
            Number(body.BlockLength),
          )
          .toString('base64'),
      });
    }),
  );
  api.post(
    '/dms_GetClientConfig',
    route((req, res) => {
      const body = payload(req.body),
        logicalName = String(body.EntityLogicalName || 'account'),
        id = String(body.RegardingId || '');
      const migration = id ? store.migration(id) : { state: 'Completed', processed: 0, total: 0 };
      res.json({
        ConfigJson: JSON.stringify({
          schemaVersion: 1,
          activeProvider: store.config.activeProvider,
          entity: {
            logicalName,
            enabled: !!entitySets[logicalName],
            visibleColumns,
            editableColumns,
            maxFileSizeMb: 100,
            allowedExtensions: [],
            blockedExtensions: ['exe', 'bat', 'cmd', 'ps1', 'com', 'scr'],
          },
          org: { maxUploadFileSizeBytes: store.config.maxUploadFileSizeBytes },
          sharePoint: { allowedHosts: ['*.sharepoint.com', 'docs.contoso.com'] },
          migration,
          features: { telemetry: false },
        }),
      });
    }),
  );
  for (const action of ['dms_EnsureMigration', 'dms_GetMigrationStatus'])
    api.post(
      '/' + action,
      route((req, res) => {
        const body = payload(req.body);
        const migration = store.migration(
          requireGuid(body.RegardingId),
          action === 'dms_EnsureMigration',
        );
        res.json({ ...migration, MigrationJson: JSON.stringify(migration) });
      }),
    );
  api.get(
    /^\/EntityDefinitions\(LogicalName='([^']+)'\)$/,
    route((req, res) => {
      const logicalName = req.params[0],
        set = entitySets[logicalName];
      if (!set) throw new ApiError(404, 'UnknownEntity', 'Entity metadata was not found.');
      res.json({
        LogicalName: logicalName,
        EntitySetName: set,
        ...(requestUrl(req).searchParams.get('$expand') === 'Attributes'
          ? { Attributes: attributes }
          : {}),
        '@odata.etag': 'W/"1"',
      });
    }),
  );
  app.use('/api/data/v9.2', api);
  app.use((_req, _res, next) => next(new ApiError(404, 'NotFound', 'Endpoint was not found.')));
  app.use((error: unknown, _req: Request, res: Response, _next: NextFunction) => {
    if (res.headersSent) return;
    const parserStatus =
      error && typeof error === 'object' && 'status' in error ? Number(error.status) : 0;
    const apiError =
      error instanceof ApiError
        ? error
        : new ApiError(
            parserStatus >= 400 && parserStatus < 500 ? parserStatus : 500,
            'InvalidRequest',
            'The request could not be processed.',
          );
    res.status(apiError.status).json({ error: { code: apiError.code, message: apiError.message } });
  });
  return app;
}
const mockRoot = path.resolve(__dirname, '..');
export const store = new Store(path.join(mockRoot, 'data'), path.join(mockRoot, 'seed-files'));
export const app = createApp(store);
if (require.main === module) app.listen(Number(process.env.PORT || 5174), '0.0.0.0');
