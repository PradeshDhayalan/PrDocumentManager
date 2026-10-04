# DMS Grid
React 16.14 + Fluent UI v9 virtual PCF document grid with a local mock Dataverse API and test harness. The approved visual prototype remains in `mockup/`. Implementation follows milestones from the uploaded brief; `CHANGELOG.md` records verified progress.

```sh
npm ci
npm run seed
npm run dev
```
Harness: http://localhost:5173 . Mock API: http://localhost:5174 . API calls from the harness use a same-origin Vite proxy.

Checks: `npm run typecheck`, `npm run lint`, `npm test`, `npm run e2e`, `npm run build`, `npm run build:pcf`.

M0 and M1 are complete: the harness renders the themed PCF shell, and the persistent mock API supports CRUD, file transfers, concurrency, migration and fault injection. UI wiring belongs to M2–M7 and is not implemented in the production control yet. The visual prototype is separate. Microsoft Graph authentication is deferred. No real tenant or Dataverse environment is connected. See `docs/spikes.md` and `docs/decisions.md`.

Sample generation needs `ffmpeg` and `ffprobe` on PATH (or `DMS_FFMPEG_PATH` / `DMS_FFPROBE_PATH`). The current environment has both. Generated sample data stays local and is ignored by git. See [mock API guide](docs/mock-api.md) for working examples.

## Opening from a cloud workspace

`localhost:5173` is only reachable on the machine running the server. This cloud session has no supported browser preview forwarding tool, so a desktop browser cannot reach the cloud listener through that address.

Run `npm run export:harness` to create `artifacts/dms-harness.zip`. Extract the ZIP on your computer, open a terminal in `dms-harness`, and run `node start.cjs` with Node 20+. Open http://localhost:5173. This prebuilt package includes the React/Fluent harness, bundled mock API and generated seed files; it needs no npm install or ffmpeg on the receiving computer. API and UI share one loopback-only port. It retains the current M0/M1 behavior; grid wiring is still pending.
