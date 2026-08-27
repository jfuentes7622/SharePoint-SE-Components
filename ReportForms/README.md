# Report Forms

Report Forms is a read-only SharePoint Framework web part for displaying a single SharePoint list item with a configurable form-style layout. It is based on DynamicForms' field renderer and visual layout designer, but it does not create or update list items.

## Compatibility

- SharePoint Server Subscription Edition
- SharePoint Framework 1.5.1 on-premises profile
- React 15.6.2
- TypeScript 2.4
- Node.js 10.24.x and npm 6.x for the legacy SPFx build toolchain

Do not upgrade Microsoft SPFx packages independently. SharePoint Server Subscription Edition requires the 1.5.1 baseline used by this project.

## Behavior

- One Report Forms instance is connected to one SharePoint list.
- The selected item can come from an item ID, an `itemid` URL query parameter, or a dynamic-data connection to ListControl, GridControl, or Calendar.
- Multiple instances can be placed on one page and connected to the same or different dynamic-data sources.
- Runtime fields are always rendered in View mode.
- Add, Edit, Save, Cancel, defaults, record-filter, and validation property pages are not exposed.
- The visual designer remains available while editing the SharePoint page so authors can select fields and configure the report layout.
- A custom **List Control** field can place a separately configured ListControl web part inside the report layout. The complete web-part host is moved at runtime and returned to its original page position when the report unmounts.
- SharePoint list and item permissions are still enforced by SharePoint. Report permission checks use `ViewListItems`, not Add or Edit permissions.

## Configure

1. Add **Report Forms** to a modern SharePoint page.
2. Select the report list in the property pane.
3. Open the visual designer from the web part while the page is in edit mode.
4. Add and arrange the fields that should appear in the report.
5. Select a fixed item ID, use the configured URL query parameter, or connect the dynamic item source to another supported web part.

### Embed a ListControl

1. Add and configure a **List Control** web part on the same page. Give it a unique **Web part name** when the page contains more than one ListControl.
2. Open the Report Forms visual designer and add **List Control** from **Custom fields**.
3. Edit the new custom field and select the named ListControl web part.
4. Set its column span or move it to the required report position, then save the designer and publish the page.

The ListControl remains an independently configured web part. Report Forms stores its dynamic source ID in the form schema and moves that web part into the custom field when the report is displayed.

When using several Report Forms instances, configure each instance with its own list and dynamic item mapping. Page composition provides the combined report.

## Build

```powershell
npm install
npx gulp bundle --ship
npx gulp package-solution --ship
```

The deployable package is generated at `sharepoint/solution/sps-report-forms.sppkg`.

## Project Identity

- Solution ID: `1eb4368e-06af-4aef-97f3-7b42fa15f434`
- Feature ID: `b0803cb4-3860-418d-b1b4-d3f613b7c93c`
- Web part component ID: `e47c4f0e-0f0d-49d5-bcab-7a4d05742037`
- Web part alias: `ReportFormsWebPart`

## License

See the repository-level `LICENSE` file.

## Usage

1. Add **Report Forms** to a page, select the SharePoint list, and use the visual designer to arrange the fields and read-only layout.
2. Choose how the report resolves an item: fixed ID, URL query parameter, filter, or dynamic selection from Calendar, List Control, or Grid Control.
3. Optionally add an embedded List Control field for related records.
4. Test the empty-selection, valid-item, missing-item, and permission-denied states before publishing.

## Properties and common configuration

- **Data:** `listName` identifies the report list and `formSchemaJson` stores the visual designer layout.
- **Item resolution:** fixed item ID, `itemIdQueryParam`, dynamic item references, and `filterJson` support direct links and same-page selection workflows.
- **Runtime:** the web part is intentionally read-only and does not expose submit or edit transitions.
- **Embedded content:** a designer-configured List Control can show records related to the current parent item.
- **Layout and diagnostics:** full-width/fixed-width settings adapt the report to the page; dynamic diagnostics log item resolution, data loading, and embedded-control discovery.

Maintain the schema through the designer and use SharePoint permissions to secure report data. Hiding a field in the layout does not change list or item access.

## Common scenarios

- Select an incident, request, or case in List Control and show a formatted detail report beside it.
- Select a Calendar event and display a richer read-only event record.
- Open a report page directly with an item ID query parameter.
- Show a parent record with an embedded list of related actions, notes, or child items.
- Combine Report Forms with Print Control for a clean printable record.
