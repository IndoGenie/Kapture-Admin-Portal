# Kapture Admin Portal — v0.3.0 (folder-based repository)

SPFx web part for the existing Kapture SharePoint site. This release removes all dependency on custom SharePoint library columns. It uses the existing `Documents` library and the existing folder `/sites/Kapture/Shared Documents/Kapture Knowledge Repository`.

## Folder routing

Uploads are routed automatically to:

- Standard departments: `Kapture Knowledge Repository/<Department>/<Category>/`
- Operations: `Kapture Knowledge Repository/Operations/<Sub-department>/<Category>/`

The app creates missing folders as needed. The department and category breakdowns in the admin analytics are inferred from these folder paths. The agreed department, Operations sub-department and category lists are built into the interface.

## Descriptions and suggested tags

Descriptions and selected predefined tags are stored in `_KaptureIndex.json` in the repository root, keyed by each file's server-relative URL. This is an application index, not SharePoint column metadata. The file is excluded from the document count/list. This simple index is intended for the single-admin upload workflow; simultaneous uploads by multiple admins could overwrite each other's index updates. SharePoint permissions remain the security boundary.

## Existing repository

Default path: `/sites/Kapture/Shared Documents/Kapture Knowledge Repository`

The app does not create custom columns and does not alter the Employee Wiki. Existing documents placed directly in the repository root or in unrecognised folder paths may not appear in department/category analytics until organised into the agreed folder structure.

## Build

Node.js 18.19.0:

```bash
npm install --no-audit --no-fund
npx gulp clean
npx gulp bundle --ship
npx gulp package-solution --ship
```

Package output: `sharepoint/solution/kapture-admin-portal.sppkg`.

The source must be built by GitHub Actions before deployment. A successful compilation does not by itself prove tenant upload permissions or live SharePoint behaviour; test with one non-sensitive sample file and verify its destination folder and portal listing before bulk use.
