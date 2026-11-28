# Kapture Admin Portal SPFx — v0.4.0

This source release upgrades the existing Kapture Admin Portal while retaining its SharePoint repository path and folder-based storage model.

## What's included

- Multi-file selection and sequential batch upload.
- Common metadata for a batch, or per-file metadata using an Excel-compatible CSV manifest.
- Downloadable manifest template (`Kapture-Bulk-Upload-Manifest.csv`). Open it in Excel, replace the example rows, save as CSV UTF-8, then import it.
- Manifest columns: `FileName`, `Department`, `SubDepartment`, `Category`, `Description`, `Tags`.
- `FileName` must exactly match the selected file name, including extension. Use one row per file. For non-Operations departments, leave `SubDepartment` blank. For Operations, use one of the configured Operations sub-departments.
- Sequential uploads, per-file progress, preservation of successful uploads, and retry of failed files only.
- Repository search, department/category filters, 10-item pagination, and an in-portal preview dialog.
- Existing `_KaptureIndex.json` metadata format and repository folder path are retained.

## Repository path

`/sites/Kapture/Shared Documents/Kapture Knowledge Repository`

Folder routing remains:

- `Department/Category/filename`
- `Operations/Operations Sub-department/Category/filename`

No new SharePoint library or metadata columns are required by this release.

## Build

Requires Node.js 18.17.1 through 18.x and npm.

```bash
npm install --no-audit --no-fund
npx gulp clean
npx gulp bundle --ship
npx gulp package-solution --ship
```

The generated package is `sharepoint/solution/kapture-admin-portal.sppkg`.

## Important validation before production

Test first with a small batch (2–3 files), including one Operations file and one non-Operations file. Confirm file routing, `_KaptureIndex.json` entries, retry behavior, filters, pagination and preview behavior in the target SharePoint tenant. The preview uses the SharePoint Office web viewer (`?web=1`); actual preview support depends on file type and tenant policies. Large-file/chunked upload is not implemented in this release; validate file-size limits with representative files before bulk migration.

The build workflow in `.github/workflows/build.yml` builds the package and uploads it as a GitHub Actions artifact.


## Version 0.4.2
- Bulk upload now collects metadata for all successful files and writes the metadata index once per batch, preserving existing index entries.
- Existing `_KaptureIndex.json` read or parse errors are surfaced instead of silently treating the index as empty.
- Existing index files are updated through SharePoint file-content PUT; first-time creation uses the folder Files API.
