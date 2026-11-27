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
