# Kapture Admin Portal — v0.2.0

SPFx web part for the existing Kapture SharePoint site. This release adds the Indovance department structure, Operations sub-departments, category/description/tags capture, automatic folder routing, metadata writing, video file selection, repository search and a redesigned responsive interface.

## Repository target

Default server-relative path:

`/sites/Kapture/Shared Documents/Kapture Knowledge Repository`

The application creates missing department folders on upload. Operations documents are stored under `Operations/<Sub-department>`; other departments are stored under `<Department>`. Categories are metadata, not folders.

## SharePoint prerequisites and important notes

- The signed-in uploader needs permission to add files and create folders in the repository.
- The implementation currently targets the standard SharePoint library title `Documents` (URL segment `Shared Documents`). Confirm this title in Site contents if the tenant has renamed the library.
- On the first upload, the app attempts to create five text metadata columns on the Documents library: Kapture Department, Kapture Sub-department, Kapture Category, Kapture Description and Kapture Tags. The uploader must have permission to add library columns. If the tenant does not allow this, a SharePoint administrator should create those columns with the same internal names before testing.
- The repository document list reads file items from the Documents library and filters them to the configured repository path.
- The Employee Wiki must separately be configured to read from this same repository path. This package does not modify the Employee Wiki.
- Video selection is enabled for common formats. Playback/preview behaviour in the Employee Wiki depends on its own implementation and SharePoint media support.
- Access control is enforced by SharePoint permissions, not by hiding the page or web part.

## Build

Use Node.js 18.19.0. From the project root:

```bash
npm install --no-audit --no-fund
npx gulp clean
npx gulp bundle --ship
npx gulp package-solution --ship
```

Package output: `sharepoint/solution/kapture-admin-portal.sppkg`.

The GitHub Actions workflow builds and uploads this package as an artifact. This source ZIP has not been tenant-deployed or live-tested; test first with a non-sensitive sample file and verify its folder and metadata in SharePoint.


## Version 0.2.1

- Rebranded the interface as Indovance Knowledge Hub and removed Kapture-specific connection messaging and the generic K mark.
- Added predefined multi-select tags to reduce manual metadata entry.
- Added repository analytics for total documents, department distribution, category distribution and Operations sub-department coverage.
- Analytics read the SharePoint metadata columns and follow SharePoint paging so the dashboard is not limited to the first 500 items.
- Preserved the existing solution ID, web-part ID, component identity and repository folder default.
- Analytics are only as accurate as the metadata saved on each file; legacy files without metadata remain unclassified in the breakdowns.


## v0.2.2 repository integration correction

- Uses the confirmed SharePoint document library title `Documents`.
- Corrects the SharePoint REST field-creation payload to `SP.XmlSchemaFieldCreationInformation`.
- Checks library access and existing fields before attempting creation.
- Reports missing-library, missing-folder and permission errors separately.
- Restricts repository listing and analytics to the configured `Kapture Knowledge Repository` folder and its descendants.
- Department and Operations sub-department folders remain automatically created on first upload.

The SharePoint account deploying/using the portal must have permission to create library columns on first initialisation, or a site owner must create them once.
