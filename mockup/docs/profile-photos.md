# Persona photos

Live Graph authentication is deferred by the user to the integration discussion. The visual harness makes no Graph calls by default and keeps Fluent's initials fallback. No credentials, tokens or profile photos are stored in browser storage.

`DocumentPerson` uses the Persona avatar `image` slot. `UserPhotoProvider` shares one request per user across list, tiles and details, handles missing/error photos, aborts pending requests and revokes Blob URLs when the host unmounts.

The host can supply `window.dmsUserPhotoLoader` before mount: `(dataverseUserId, signal) => Promise<Blob | null>`. This is a harness injection point, not a production PCF authentication contract. Name-to-sample-ID mapping is limited to the mockup; production records must supply the actual Dataverse user ID.

`ProfilePhotoLoaders.mjs` provides two integration adapters:

- `createGraphPhotoLoader({ resolveEntraUserId, getAccessToken })` requests `GET https://graph.microsoft.com/v1.0/users/{entraUserId}/photo/$value` with a host-supplied bearer token. A Dataverse systemuser GUID must first be mapped to the user's Entra object ID, not used interchangeably or inferred from the display name.
- `createProxyPhotoLoader({ getPhotoUrl, origin })` uses an authenticated same-origin server endpoint. The server handles user identity mapping and Graph permissions. This endpoint has not been built.

Graph consent/permissions, tenant access, host authentication and the chosen adapter remain integration decisions. Authentication and missing-photo responses (401/403/404) fall back to initials. Graph calls remain specific to profile photos; the document storage providers do not gain Graph/SharePoint access.

Run `node test/profile-photos.mjs` to verify the photo loading contracts with fake HTTP responses; it does not contact Microsoft Graph.
