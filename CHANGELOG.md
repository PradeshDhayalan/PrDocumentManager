# Changelog
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
