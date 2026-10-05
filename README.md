# DMS Grid
React 16.14 + Fluent UI v9 virtual PCF document grid with a local mock Dataverse API and test harness. The approved visual prototype remains in `mockup/`. Implementation follows milestones from the uploaded brief; `CHANGELOG.md` records verified progress.

```sh
npm ci
npm run dev
```
Harness: http://localhost:5173 . Mock API: http://localhost:5174 . API calls from the harness use a same-origin Vite proxy.

Checks: `npm run typecheck`, `npm run lint`, `npm test`, `npm run e2e`, `npm run build`, `npm run build:pcf`.

The control now loads real sample rows from the persistent mock API. The interactive List and Tile views include Microsoft file icons, image thumbnails, Fluent Personas, card shadows, shared selection, command overflow, search/presets, paging, filters and a details pane. Upload, duplicate detection, download, preview, copy-link, SharePoint-reference creation, metadata editing and deletion are connected to the API.

The harness provides four themes, record scenarios (including 1,200 Fabrikam documents), and an active-provider switch. `npm run dev` generates samples automatically on first use; later starts preserve your uploads and edits. `npm run seed` explicitly resets local data. Video generation uses ffmpeg when installed and an included original sample otherwise. `ffprobe` is needed for the video validation test, but not to run the app.

Microsoft Graph authentication/photos and real-tenant Dataverse validation remain deferred. Personas currently use initials. JPEG/PNG previews use the stored image bytes; Office/PDF tiles use illustrative covers, while PDF/image/text files can be opened in the preview dialog. Office rendering, SharePoint embedding and production server plugins remain outside this mock integration. See `docs/spikes.md` and `docs/decisions.md`.

To update your existing Mac clone, stop its running dev command with Ctrl+C, then:

```sh
cd /Users/pradeshdhayalan/Documents/ChatGPT/PrDocumentManagementGrid
git switch codex/dms-grid-review
git pull --ff-only
npm ci
npm run dev
```

Open http://localhost:5173 on that Mac. Both the harness and mock API must be running. Source files are in `control/DmsGrid/src/`; the harness mounts the PCF `init`/`updateView` lifecycle.

## Opening from a cloud workspace

`localhost:5173` is only reachable on the machine running the server. This cloud session has no supported browser preview forwarding tool, so a desktop browser cannot reach the cloud listener through that address.

Run `npm run export:harness` to create `artifacts/dms-harness.zip`. Extract the ZIP on your computer, open a terminal in `dms-harness`, and run `node start.cjs` with Node 20+. Open http://localhost:5173. This prebuilt package includes the React/Fluent harness, bundled mock API and generated seed files; it needs no npm install or ffmpeg on the receiving computer. API and UI share one loopback-only port. It includes the interactive grid and mock file operations.

Current harness screens: [List](docs/screenshots/list.png) · [Tiles](docs/screenshots/tiles.png).
