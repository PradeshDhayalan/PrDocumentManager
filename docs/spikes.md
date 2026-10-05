# Real-environment verification
| ID | Validation | Boundary | State |
|---|---|---|---|
| S1 | Dataset regarding filter/page size and saved/unsaved form record context | PCF init adapter | Open |
| S2 | SharePoint iframe CSP, tenant framing and reliable fallback | SharePoint preview adapter | Open |
| S3 | Annotation block action payloads, token lifetime and limits | Transfer services | Open |
| S4 | Fetch Web API headers and certification; real authorisation | Dataverse client | Open |
| S5 | Attribute metadata, field security and option colours | Metadata mapper | Open |
| S6 | Polymorphic annotation parent binding/entity set resolution | Note provider | Open |
| S7 | Commercial redistribution of Microsoft file icon assets | Icon source/notices | Open |
| S8 | Exact documented/target platform versions and bundle size | Manifest/package | Microsoft Learn blocked; PAC 2.12.2 template defaults to Fluent 9.68.0; requested 9.46.2 pinned pending verification |

## M1 isolation
- `mock-api/src/api/store.ts` contains the simplified annotation block upload/download protocol and resume behavior (S3). Token lifetimes across restarts remain unverified.
- `mock-api/src/api/metadata.ts` contains simplified metadata labels/options/security flags (S5).
- `mock-api/src/api/store.ts` isolates annotation parent binding and ownership-aware pre-delete cascade (S6). No production client deletes annotations.
- No real-org assumptions were marked verified. The mock tests prove local contracts only.

## Interactive control isolation
The typed `DataverseClient`, `DocumentRepository`, provider registry, `MetadataService` and server compensation action implement the local contracts. The production bundle builds at about 286 KiB and mounts with React 16.14/platform Fluent in the browser test. This does not verify a target Dataverse tenant or the S1–S8 contracts above. SharePoint embedding remains disabled; Graph authentication/photos remain deferred.
