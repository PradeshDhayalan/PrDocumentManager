# DMS Grid
React 16.14 + Fluent UI v9 virtual PCF document grid with a local mock Dataverse API and test harness. The approved visual prototype remains in `mockup/`. Implementation follows milestones from the uploaded brief; `CHANGELOG.md` records verified progress.

```sh
npm ci
npm run seed
npm run dev
```
Harness: http://localhost:5173 . Mock API: http://localhost:5174 . API calls from the harness use a same-origin Vite proxy.

Checks: `npm run typecheck`, `npm run lint`, `npm test`, `npm run e2e`, `npm run build`, `npm run build:pcf`.

The local PCF control now implements the approved mockup: list and tile views, selection, server-side search/filter/sort/paging, details, metadata editing (including bulk updates), file uploads/downloads, previews, deletion and SharePoint links. The persistent mock API supplies local files, metadata, concurrency checks, migration and fault injection. The original visual prototype is preserved in `mockup/`. Microsoft Graph authentication is deferred. No real tenant or Dataverse environment is connected. See `docs/spikes.md` and `docs/decisions.md`.

Sample generation needs `ffmpeg` and `ffprobe` on PATH (or `DMS_FFMPEG_PATH` / `DMS_FFPROBE_PATH`). Set those variables if FFmpeg is installed outside PATH. Generated sample data stays local and is ignored by git. See [mock API guide](docs/mock-api.md) for working examples.

## Opening from a cloud workspace

`localhost:5173` is only reachable on the machine running the server. This cloud session has no supported browser preview forwarding tool, so a desktop browser cannot reach the cloud listener through that address.

Run `npm run export:harness` to create `artifacts/dms-harness.zip`. Extract the ZIP on your computer, open a terminal in `dms-harness`, and run `node start.cjs` with Node 20+. Open http://localhost:5173. This prebuilt package includes the React/Fluent harness, bundled mock API and generated seed files; it needs no npm install or ffmpeg on the receiving computer. API and UI share one loopback-only port. It includes the working local grid and mock API.

## Using the local control

Run `npm run dev` and open http://localhost:5173. The control opens directly in the Blue light theme with no harness header or theme switcher. Contoso (22 sample documents) is the default; sample test routes use `?record=fabrikam`, `?record=empty`, `?record=unsaved` and `?record=legacy`. The configured mock API provider determines whether the command bar offers Notes uploads or SharePoint references. The control uses Microsoft’s default Segoe UI font stack. The local harness loads Microsoft-hosted Segoe UI regular, semibold and bold webfonts when the installed font is unavailable (such as on macOS). This requires access to Microsoft’s font CDN.

Uploads accept multiple files or drag and drop, enforce configured size/extensions, display progress, and retry transient block failures. Downloaded files retain their original bytes and filenames. Metadata edits and deletes use ETags; a conflicting change prompts a refresh. List and tile layouts share selection. Use the More commands menu for density, additional columns and keyboard help, including search on narrow screens.

Images, PDFs, video and text have local previews. Office and unsupported formats download for viewing in their applications. Office covers are representative thumbnails. SharePoint references open externally with existing permissions; they do not upload or delete SharePoint files. Azure Blob sample rows have no connected file provider. Real Dataverse/SharePoint authentication and tenant verification remain deferred.

`npm run e2e` uses Playwright’s installed Chromium by default. Install it with `npx playwright install chromium`, or set `PLAYWRIGHT_CHROMIUM_EXECUTABLE` to an existing compatible browser.

## Custom action buttons

Set the PCF input `customActionsJson` to the configuration in [examples/custom-actions.json](examples/custom-actions.json). Buttons render inside the control, use selection limits, and move into More commands when space is limited. The local harness includes working **Mark as reviewed** and **Show selection** examples.

Register named handlers through the form JavaScript web resource in [examples/custom-actions.webresource.js](examples/custom-actions.webresource.js). The PCF raises `OnCustomAction` with selected records and a completion callback; it does not evaluate JavaScript stored in JSON. See [developer setup and event contract](docs/custom-actions.md). Custom PCF events are a Microsoft preview feature; local behavior is verified, while real model-driven form registration needs tenant verification.
