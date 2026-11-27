# Kapture Admin Portal – SPFx source (v0.1.1)

A separate SPFx web part for an administrator-only page on the existing Kapture SharePoint site. It uses the same SharePoint site and existing tenant App Catalog as the Employee Wiki.

## Architecture
- Existing site: `https://teknovance.sharepoint.com/sites/Kapture`
- New modern page: `SitePages/Kapture-Admin.aspx` (Ajit creates the page)
- One document library: `Kapture Knowledge Repository`
- No Incoming/Published folders and no separate Knowledge Catalogue List in this initial version.
- Admin Portal uploads directly to the library root. Employee Wiki will be configured separately to read from the same library.

## SharePoint setup
1. Create the `Kapture Knowledge Repository` document library on the existing Kapture site.
2. Give the administrator group Edit/Contribute permissions and employees Read permissions.
3. Restrict the `Kapture-Admin.aspx` page to the authorised administrator group.
4. Confirm the library server-relative URL with the developer. The default configured path is `/sites/Kapture/Shared Documents/Kapture Knowledge Repository`; it can be changed in the web-part property pane.
5. Deploy this package through the existing tenant App Catalog, then add the Kapture Admin Portal web part to the new page. No new site or App Catalog is required.

## Build
Requires Node.js 18.19.0 and SPFx 1.20.0 dependencies. Run `npm install`, `gulp clean`, `gulp bundle --ship`, `gulp package-solution --ship`. Output: `sharepoint/solution/kapture-admin-portal.sppkg`.

## Security
The UI is not a security boundary. Enforce access through SharePoint permissions. Administrators need edit permission on the document library; employees should have read-only access. SPFx REST requests run as the signed-in user.

## Scope limits
This starter uploads files directly to the shared library and lists the files in its root. It does not yet implement document extraction, metadata columns, versioned article generation, or advanced management workflows.

## GitHub Actions setup
Create a separate GitHub repository for this Admin Portal (recommended, so its workflow cannot interfere with the Employee Wiki build). Extract this ZIP and upload the **contents inside the `Kapture-Admin-Portal` folder** to the root of that repository, including the `.github` folder. Push to `main` or run the workflow manually. Download the `kapture-admin-portal-sppkg` artifact from the successful workflow run and give that `.sppkg` to Ajit.
