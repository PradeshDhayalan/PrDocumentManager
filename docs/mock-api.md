# Mock API — M1

The mock is local Express/TypeScript, backed by `mock-api/data/state.json` and `data/uploads/<annotation-guid>`. It imitates the subset of Dataverse described in the brief; it is not a real Dataverse implementation.

## Run

Install Node 20+, npm and ffmpeg/ffprobe. Run `npm ci`, `npm run seed`, then `npm run dev` from the repository root. The harness is at http://localhost:5173, API at http://localhost:5174. Vite proxies `/api`, `/__mock`, and `/mock-sharepoint` for same-origin browser access. Production PCF HTTP wiring is M2; M0's control shell still shows its empty state.

`npm run seed` resets all local mock data. `POST /__mock/reset` restores generated fixtures and clears uploads, transfer sessions and injected faults. Uploaded files and document metadata persist across API process restarts. Uncommitted transfer tokens are deliberately process-local; S3 tracks real token lifetime/resume behavior.

## API examples

```sh
curl 'http://localhost:5174/api/data/v9.2/dms_documents?$count=true&$top=5'

curl -X POST http://localhost:5174/api/data/v9.2/annotations \
  -H 'Content-Type: application/json' \
  -d '{"annotationid":"bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb","filename":"Hello.txt","mimetype":"text/plain","documentbody":"SGVsbG8hCg==","objectid_account@odata.bind":"/accounts(11111111-1111-4111-8111-111111111111)"}'

curl -X POST http://localhost:5174/api/data/v9.2/dms_documents \
  -H 'Content-Type: application/json' \
  -d '{"dms_documentid":"cccccccc-cccc-4ccc-8ccc-cccccccccccc","dms_name":"Hello.txt","dms_regardingid":"11111111-1111-4111-8111-111111111111","dms_regardingtype":"account","dms_provider":100000000,"dms_storageref":"bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb","dms_uploadstate":100000001}'

curl -X PATCH 'http://localhost:5174/api/data/v9.2/dms_documents(cccccccc-cccc-4ccc-8ccc-cccccccccccc)' \
  -H 'Content-Type: application/json' -H 'If-Match: W/"1"' \
  -d '{"dms_name":"Edited.txt","dms_description":"An edited sample"}'

curl http://localhost:5174/__mock/files/bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb -o Hello.txt
```

Every document/annotation response includes `@odata.etag`. Entity GETs also expose an `ETag` header. Supply it with PATCH/DELETE: stale values return 412. `If-Match: *` explicitly overwrites an existing row; it does not create one. `If-None-Match: *` prevents overwrites. PATCH without If-Match can upsert a new GUID. Deleting a Note row removes its owned annotation/file on the server; deleting a SharePoint row leaves binaries untouched.

## Blocks

All action paths below start with `/api/data/v9.2/`.

1. `InitializeAnnotationBlocksUpload`: `{Target: <annotation metadata including parent bind>}` → `FileContinuationToken`.
2. `UploadBlock`: `{FileContinuationToken, BlockId: <base64 ID>, BlockData: <base64, ≤4 MiB decoded>}`. Re-sending a stored block is supported.
3. `CommitAnnotationBlocksUpload`: `{FileContinuationToken, BlockList: <ordered IDs>}` → `FileSizeInBytes`. A matching repeat commit returns the original result. Commit reads one block at a time rather than concatenating the whole file in server memory.
4. `InitializeAnnotationBlocksDownload`: `{Target:{annotationid}}` → token, file size/name.
5. `DownloadBlock`: `{FileContinuationToken, Offset, BlockLength}` → `{Data:<base64>}`; request blocks up to 4 MiB.

This simplified action contract is isolated in the mock Store and must be verified against real annotation actions (S3/S6).

## Queries and tools

- Document/annotation lists support `$select`, `$filter`, `$orderby`, `$top`, `$skip`, `$count=true`, `Prefer: odata.maxpagesize=N`, and `@odata.nextLink`. Filters support eq/ne/gt/ge/lt/le, and/or/parentheses, contains, escaped quoted strings, numbers, booleans and null. Date comparisons use quoted ISO values. Counts are before top/paging. Selected attributes include their formatted values.
- `dms_GetClientConfig`: `{EntityLogicalName, RegardingId?}` → `{ConfigJson}`.
- `dms_EnsureMigration` / `dms_GetMigrationStatus`: `{RegardingId, EntityLogicalName}`. Polling advances legacy rows in batches over approximately 15 seconds. Repeated calls are idempotent. No background migration engine is implemented.
- `EntityDefinitions(LogicalName='dms_document')?$expand=Attributes` includes all editor categories; `?$select=EntitySetName` resolves seeded entity names. Metadata is simplified (S5).
- `/__mock/state` returns counts and fault settings. `/__mock/config` accepts latency, failure rate, failure mode, block failure number, size limit, active provider, framing flag and migration state.
- `x-mock-fail: 429|500|412|timeout|offline` injects a per-request fault; `x-mock-latency-ms` delays a request. 429 supplies Retry-After. `failureRate` applies the configured `failWith` to a fraction of API requests. `failNextUploadAtBlock` fails one incoming new block, then clears; repeat the same block/token to resume.
- `/mock-sharepoint/view/sample` is a local frameable sample; `?blocked=1` or `spBlockFraming:true` sets framing denial headers. No target SharePoint URL is fetched.

## Seed and verification

Contoso has 22 rows because the full fixture list contains 14 files, plus five SharePoint links and three state/provider edge cases. Fabrikam has 1,200 rows, Empty Contact has none, and Legacy Notes Deal starts with 40 annotations and no document rows. Six Modified By users and varied dates are included.

Office files are ZIP-based valid packages, PDFs contain three pages, JPEG/PNG are generated original images, ZIP contains sample files, and MP4 is a real two-second test pattern. `Network Diagram.vsdx` intentionally contains placeholder bytes as specified. Existing installed ffmpeg is used instead of downloading ffmpeg-static; the generator fails clearly if it is unavailable.

Jest/supertest verifies CRUD, validation, filter/order/paging, formatted values, single-file equality, persisted bytes, ownership-aware cascade, ETags, size limits, ordered blocks, 30 MiB interrupted/resumed transfer and download byte equality, migration, config, metadata, offline/throttle faults, framing and reset. Fixtures are also decoded with Office ZIP inspection, PDF/image libraries and ffprobe.
