# Changelog
## Phase 2 requirements and documentation

Recorded the requested pagination, PCF configuration, custom columns, double-click editing, drag-and-drop overlay, tile pencil removal and Office thumbnail explanation in [Phase 2 requirements](docs/phase2-requirements.md). Added the requirement to update README and relevant supporting documentation with every completed change, and to record changes and validation in this changelog.

The local implementation is described in [grid configuration](docs/grid-configuration.md). Harness/mock API builds, production PCF build (119 KiB), typecheck, ESLint and the three new browser tests passed. The custom-column API test verified metadata and typed values survive a store restart. Broader checks encountered missing `ffprobe`, an unavailable CDN icon and pre-existing formatting issues. Production tenant integration remains unverified.

Validation for this documentation update: reviewed the requirements against the requested changes and checked the documentation links and whitespace. No application behavior changed.

## M0 — scaffold (complete)
Official PAC virtual dataset scaffold, strict TypeScript workspace, shared resx, four-theme harness and check scripts.

Verified: typecheck, lint, Jest (1 test), Playwright (4 themes + unsaved), harness/mock builds; PCF platform bundle 17,446 bytes. React is deduplicated at 16.14; Griffel 1.5.32 is pinned for classic JSX compatibility.

## M1 — mock API and samples (complete)
Persistent mock Dataverse CRUD/query/config/metadata API; ETag preconditions and ownership-aware server cascade; real file upload/download; resumable ordered 4 MiB blocks; fault injection; polling-driven legacy migration; reset/debug tools.

Seed: 22 Contoso rows, 1,200 Fabrikam rows, 40 legacy annotations, six users and 14 sample formats. Office packages, multi-page PDFs, images, ZIP and a two-second video are valid/openable; Visio is the specified placeholder.

Verified: nine Jest tests (including 30 MiB byte-equal download after an injected 429), Playwright themed shell/keyboard/unsaved checks, typecheck, lint, harness/mock builds, PCF bundle 17,446 bytes. Production grid/API wiring remains the next milestones.

## Working local control

Replaced the placeholder shell with the approved Fluent UI document library: list/tiles, shared selection, search, type/status filters, saved view presets, sorting, paging, details, metadata-driven single/bulk editing, additional columns, density, fullscreen and responsive command overflow. Uses the default Microsoft Segoe UI font stack.

Connected persistent mock document CRUD, HTTPS SharePoint references, file validation, single/block uploads with progress/cancellation and transient-block retry, byte-preserving downloads, local image/PDF/video/text previews, and legacy migration. Saves/deletes enforce ETags. Upload failure compensation uses server-owned document deletion; the browser does not delete annotations.

The harness exposes all sample parent records and provider modes. Real tenant integration remains subject to the existing verification spikes. Office thumbnails are representative covers; Office files open through download.

## Local control interaction and layout refinements

Added mouse drag selection in list and tile views, with Ctrl/Cmd/Shift for additive selection and Escape to cancel. Cards dedicate more space to previews, clamp names to two lines, and omit status badges. The expiry column displays dates without expiry tags. Fluent themes explicitly use the requested Microsoft default font family including BlinkMacSystemFont.

## Verified rendered typography

Browser font inspection found that macOS rendered San Francisco despite the requested Segoe UI CSS stack. The local harness now loads Microsoft's hosted Segoe UI webfonts for regular, semibold and bold weights, preserving the exact requested family stack. Verification checks the rendered platform font rather than the font-family declaration alone.

## Blue theme and standalone control surface

Removed the entire local harness header, theme/record/provider selectors, breadcrumbs and account introduction. The control uses the Blue light theme, including when the host supplies a dark theme. Sample records remain available through test URLs.

## Configurable custom actions

Added the `customActionsJson` PCF input and `OnCustomAction` event. Named handlers receive a snapshot of selected records, parent context and a completion callback. Buttons support icons, ordering, selection limits, confirmation, busy states, overflow and refresh after success. Configuration validates without evaluating source code; errors and a 30-second completion timeout appear in the control.

The local harness uses the same example JavaScript web resource intended for form OnLoad registration. Mark as reviewed persists a Description marker; Show selection displays selected names. Unit tests cover parsing, dispatch, completion, timeout and cancellation; browser coverage verifies confirmation, persistence, overflow and selection limits. Real tenant event registration remains unverified.

## Earlier remote implementation history

The following entries describe the earlier GitHub revisions. The local v2 control incorporates the subsequent UI and custom-action changes above.

## Fresh-clone setup fix
Generate PCF manifest types automatically after dependency installation so type-checking and CI work without a prior PCF build.

## Interactive mock integration
Replaced the empty shell with API-backed Fluent List/Tile document views, shared range/toggle selection and keyboard navigation, command overflow, column menus/resizing, server search/presets/paging, status/expiry badges, image thumbnails, Personas and a details pane. Added config/capability-based provider resolution, including unavailable-provider fallback.

Connected a three-worker upload queue with incremental SHA-256 duplicate checks, filename/size/extension validation, 4 MiB block transfers, progress/cancel/retry and server compensation. Connected owned-byte download, safe text/image/PDF preview, reference opening, copy links, SharePoint URL validation, bulk deletion and metadata-driven editing with ETag reload/overwrite handling.

Harness adds record/provider switches and automatic first-run sample setup without resetting existing data. Included an original MP4 fallback so running samples does not require ffmpeg. DatePicker compatibility uses platform Fluent exports in the PCF build to preserve contexts and avoid duplicated core components.

Validation: 16 Jest tests; browser checks for both selection views, themes/unsaved/empty states, 1,200-row paging, upload/duplicate/preview/byte-equal download/edit/delete, SharePoint links, narrow command overflow, built-PCF date picker and metadata conflicts. Typecheck/lint, harness/API builds and production PCF build checked. Production PCF bundle: 292,884 bytes, below the 500 KiB limit.

These checks validate the mock integration. Graph photos, tenant authorization, real action payloads/security/metadata mapping, production cascade plugins, actual Office thumbnail rendering and optional SharePoint embedding remain deferred; S1–S8 are not marked tenant-verified.

## Portable review package

Published `downloads/dms-grid-review.zip` with the interactive app, bundled mock API, generated sample files, dependency notices and Mac launcher. Defaults to port 5180, opens the browser on Mac with `--open`, and exposes the packaged revision in `build-info.json`. Verified from a fresh extracted folder without installing dependencies: 22 documents, successful upload, byte-equal download, deletion and no browser errors.

## Configurable columns and tile spacing

All built-in and metadata table columns can be shown, hidden and reordered in Edit columns. Header drag grips reorder columns; Fluent resize handles change their widths. Entity visibleColumns defines initial order, and Reset restores it. Resizing and column dragging do not trigger record drag selection. Tiles have 24px desktop gutters, more body padding and lighter shadows. File visuals now use Microsoft's official Office and file-type artwork.

Merged the remote setup and portable-review improvements with the local v2 control. Preserved automatic seed setup, manifest generation, existing review artifacts and the earlier implementation modules. Restored the pinned React 16-compatible Griffel runtime for production PCF builds.
