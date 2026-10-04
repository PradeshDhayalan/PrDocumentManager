# Microsoft file icon assets — visual review 02

The locally rendered SVGs in `public/icons` are Microsoft-authored artwork extracted unchanged from the exported SVG render functions in `@fluentui/react-icons-northstar` version 0.71.4, published by the Microsoft Fluent UI project. Only static SVG files are used; no Northstar component code or UI library is bundled or used at runtime. These are Microsoft Office/Teams-era file icons, not a claim of exact parity with the current SharePoint icon set.

Assets: WordColorIcon, ExcelColorIcon, PowerPointColorIcon, FilesPdfColoredIcon, FilesPictureColoredIcon, FilesTextColoredIcon, FilesGenericColoredIcon.

Source: https://github.com/microsoft/fluentui/tree/master/packages/fluentui/react-icons-northstar

Copyright (c) Microsoft Corporation. The repository's MIT licence is preserved at `public/icons/MICROSOFT-LICENSE.txt`. That licence expressly says fonts and icons are also subject to https://aka.ms/fluentui-assets-license. The redirected asset licence PDF could not be retrieved because the environment blocks static2.sharepointonline.com. The separate asset conditions and paid AppSource redistribution remain unverified (SPIKE S7). These assets are used for the user's local visual review; no commercial release or external deployment has been made. Do not treat the package's MIT licence as commercial asset clearance.

The preferred `@fluentui/react-file-type-icons` package 8.18.2 was inspected. It generates CDN URLs and does not contain the actual SVG assets. Its Office CDN is blocked here; no runtime CDN dependence has been added.

UI chrome icons remain `@fluentui/react-icons` 2.0.308. All UI components remain `@fluentui/react-components` 9.46.2.

## Local sample JPG

`public/previews/site-photo.jpg` is an unchanged copy of scikit-learn's `china.jpg` sample, by danielbuechele, licensed CC BY 2.0. Photo: https://www.flickr.com/photos/danielbuechele/6061409035/ ; licence: https://creativecommons.org/licenses/by/2.0/ . The source attribution is preserved in `public/previews/PHOTO-ATTRIBUTION.txt`. It stands in for Site Photo.jpg in the local mockup.
