# Phase 2 requirements

Recorded on 6 October 2026. This phase covers the requested grid enhancements and the ongoing documentation requirement.

| ID | Requirement | Acceptance criteria |
| --- | --- | --- |
| P2-01 | Pagination | Display 10 documents per page by default. Provide Previous/Next navigation and allow the page size to be configured through PCF configuration. |
| P2-02 | Configurable grid | Allow displayed columns, editable attributes, filter columns, search field, page size and tile size to be configured through PCF configuration. Support small, medium and large tiles. Default the search field to Name. |
| P2-03 | Custom columns | Allow new document columns to be created and included in the grid through configuration. In the local app, persist custom column metadata and values. In Dataverse, discover published Document table columns and allow their logical names to be configured. |
| P2-04 | Double-click editing | Double-clicking a document row or tile preview opens its editable document attributes. Saving persists the changes and respects the configured editable columns. |
| P2-05 | Drag-and-drop feedback | Shade the grid background while files are dragged over it and display clear instructions confirming that files can be dropped to upload. Remove the overlay on drop, drag exit or cancellation. |
| P2-06 | Tile actions | Remove the pencil button from tile view while retaining document attribute editing through double-click and the command bar. |
| P2-07 | Office preview explanation | Document how PPT/PPTX, DOC/DOCX and Excel thumbnails are generated. Explain that current thumbnails use representative React/CSS artwork rather than actual document contents, and identify what would be needed for accurate file previews. |
| P2-08 | Documentation after every update | Update `README.md` and any relevant supporting documentation with every completed feature change, fix or configuration update. Keep behavior, defaults, setup instructions and known limitations consistent with the app. Record completed changes and validation results in `CHANGELOG.md`. Documentation is part of completing an update. |

## Implementation and verification

See [grid configuration](grid-configuration.md) for the implemented local configuration, custom columns, pagination and Office thumbnail explanation. See [the changelog](../CHANGELOG.md) for recorded progress and validation. These requirements do not establish that production Dataverse integration has been verified.

The Office preview requirement is to explain the current behavior; accurate Office document rendering has not been requested as part of this phase.
