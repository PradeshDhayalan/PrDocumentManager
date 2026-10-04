# DMS Grid visual prototype

A standalone design-review harness with React 16.14 and Fluent UI v9 9.46.2. This is a visual prototype, not milestone M0 or a deployable PCF control. Platform-library compatibility has not yet been verified against current Microsoft documentation.

## Run

```sh
npm install
npm run dev
```

Open the Vite URL. `npm run build` creates a production preview bundle.

## Review

Review 02 uses the native Fluent DataGrid, Toolbar, InlineDrawer and OverlayDrawer. The surrounding form shell has been simplified.

- List and Tiles share selected document IDs.
- Teal, Blue, Dark and High contrast themes.
- Notes upload action / SharePoint New > Link action.
- Search, document-type filters, view presets, name sorting, multi-selection, details, expansion and sample dialogs.
- Review 02 replaces the original icons with locally rendered Microsoft-authored static SVGs; provenance and the unresolved commercial asset licence are documented in `docs/third-party-notices.md`.
- Sample records are entirely local. No document or SharePoint API requests occur.

## Deliberate prototype limits

Upload, download, link creation and metadata save are simulations. Delete removes sample rows in memory; reload restores them. No real files are handled. Metadata editing uses a Fluent OverlayDrawer. Full grid keyboard navigation, paging, virtualisation, PCF packaging, resx localisation and production accessibility testing are deferred until design approval. English text is sample copy. Styles use Griffel and Fluent theme tokens; the prototype's custom Teal brand ramp defines its harness theme.

## Verification

Production preview build passed. Playwright checked shared selection across views, search and empty results, the link dialog, and absence of horizontal page overflow at 320, 480, 720, 1024 and 1440 px. All four themes rendered without browser runtime errors. This is not a full accessibility audit. With the dev server running, use `npm run check:review` and `npm run capture` (scripts use `/usr/bin/chromium`). Screenshots are in `screenshots/`.

## Review 03 — thumbnails and people

List rows, tiles, details and the preview dialog render content thumbnails. Office/PDF/text/link covers are representative sample UI previews rather than rendered file bytes. JPG previews use a real local sample image. Unavailable files retain a fallback. Fluent Personas with coloured initials appear in Modified By, tiles and details activity; this supersedes the brief's original text-only list decision at the user's request. No external profile-photo or thumbnail calls are made.

## Review 04 — card depth and Graph photo support

Tiles now use Fluent shadow4 at rest, shadow16 on hover/focus, and shadow8 with the selected outline. Motion respects the reduced-motion preference. Persona avatar images accept Blob photos from the authenticated host, with shared caching and unmount cleanup. Graph authentication is deferred by the user, so the preview still displays initials. See `docs/profile-photos.md`. Run `node test/profile-photos.mjs` for the photo contract checks.

## Review 05 — command bar and overflow

All document command-bar buttons are Fluent v9 ToolbarButton components inside a Toolbar. Fluent v9 does not expose the v8 CommandBar component. Overflow, OverflowItem, useOverflowMenu and useIsOverflowItemVisible move actions into an accessible More commands menu instead of wrapping. Dialogs and form fields retain standard Fluent components. Search remains usable from overflow through a Fluent dialog. UI code is React 16.14; no other UI component library or Fluent v8 UI package is used. The PCF control and backend remain deferred; this statement covers the visual harness.

Run `node test/command-overflow.mjs` with the dev server running to verify narrow-width search and selection actions from overflow.
