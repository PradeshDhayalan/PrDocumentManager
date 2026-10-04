# Changelog
## M0 — scaffold (complete)
Official PAC virtual dataset scaffold, strict TypeScript workspace, shared resx, four-theme harness and check scripts.

Verified: typecheck, lint, Jest (1 test), Playwright (4 themes + unsaved), harness/mock builds; PCF platform bundle 17,446 bytes. React is deduplicated at 16.14; Griffel 1.5.32 is pinned for classic JSX compatibility.

## M1 — mock API and samples (complete)
Persistent mock Dataverse CRUD/query/config/metadata API; ETag preconditions and ownership-aware server cascade; real file upload/download; resumable ordered 4 MiB blocks; fault injection; polling-driven legacy migration; reset/debug tools.

Seed: 22 Contoso rows, 1,200 Fabrikam rows, 40 legacy annotations, six users and 14 sample formats. Office packages, multi-page PDFs, images, ZIP and a two-second video are valid/openable; Visio is the specified placeholder.

Verified: nine Jest tests (including 30 MiB byte-equal download after an injected 429), Playwright themed shell/keyboard/unsaved checks, typecheck, lint, harness/mock builds, PCF bundle 17,446 bytes. Production grid/API wiring remains the next milestones.
