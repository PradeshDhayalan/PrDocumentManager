# Changelog
## M0 — scaffold (complete)
Official PAC virtual dataset scaffold, strict TypeScript workspace, shared resx, four-theme harness and check scripts.

Verified: typecheck, lint, Jest (1 test), Playwright (4 themes + unsaved), harness/mock builds; PCF platform bundle 17,446 bytes. React is deduplicated at 16.14; Griffel 1.5.32 is pinned for classic JSX compatibility.

## M1 — mock API and samples (complete)
Persistent mock Dataverse CRUD/query/config/metadata API; ETag preconditions and ownership-aware server cascade; real file upload/download; resumable ordered 4 MiB blocks; fault injection; polling-driven legacy migration; reset/debug tools.

Seed: 22 Contoso rows, 1,200 Fabrikam rows, 40 legacy annotations, six users and 14 sample formats. Office packages, multi-page PDFs, images, ZIP and a two-second video are valid/openable; Visio is the specified placeholder.

Verified: nine Jest tests (including 30 MiB byte-equal download after an injected 429), Playwright themed shell/keyboard/unsaved checks, typecheck, lint, harness/mock builds, PCF bundle 17,446 bytes. Production grid/API wiring remains the next milestones.

## Fresh-clone setup fix
Generate PCF manifest types automatically after dependency installation so type-checking and CI work without a prior PCF build.

## Interactive mock integration
Replaced the empty shell with API-backed Fluent List/Tile document views, shared range/toggle selection and keyboard navigation, command overflow, column menus/resizing, server search/presets/paging, status/expiry badges, image thumbnails, Personas and a details pane. Added config/capability-based provider resolution, including unavailable-provider fallback.

Connected a three-worker upload queue with incremental SHA-256 duplicate checks, filename/size/extension validation, 4 MiB block transfers, progress/cancel/retry and server compensation. Connected owned-byte download, safe text/image/PDF preview, reference opening, copy links, SharePoint URL validation, bulk deletion and metadata-driven editing with ETag reload/overwrite handling.

Harness adds record/provider switches and automatic first-run sample setup without resetting existing data. Included an original MP4 fallback so running samples does not require ffmpeg. DatePicker compatibility uses platform Fluent exports in the PCF build to preserve contexts and avoid duplicated core components.

Validation: 16 Jest tests; browser checks for both selection views, themes/unsaved/empty states, 1,200-row paging, upload/duplicate/preview/byte-equal download/edit/delete, SharePoint links, narrow command overflow, built-PCF date picker and metadata conflicts. Typecheck/lint, harness/API builds and production PCF build checked. Production PCF bundle: 292,884 bytes, below the 500 KiB limit.

These checks validate the mock integration. Graph photos, tenant authorization, real action payloads/security/metadata mapping, production cascade plugins, actual Office thumbnail rendering and optional SharePoint embedding remain deferred; S1–S8 are not marked tenant-verified.
