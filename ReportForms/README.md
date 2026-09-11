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
- **Repeat report for every list item** is off by default. When enabled, published/runtime view loads every item ID from the selected list in ascending order and renders the configured report once per item. SharePoint paging links are followed so the feature is not limited to the first response page.
- The selected item can come from an item ID, an `itemid` URL query parameter, or a dynamic-data connection to ListControl, GridControl, or Calendar. When **Use dynamic item ID as form item ID** is off, the incoming or URL value filters the configured writable target column instead of being treated as SharePoint `ID`; supported targets include text, number, Boolean, DateTime, lookup, and person columns.
- Multiple instances can be placed on one page and connected to the same or different dynamic-data sources.
- Runtime fields are always rendered in View mode.
- Add, Edit, Save, Cancel, defaults, record-filter, and validation property pages are not exposed.
- The visual designer remains available while editing the SharePoint page so authors can select fields and configure the report layout.
- Form Layout includes **Vertical space between fields (px)** from 0 to 100 pixels. It is the sole runtime gap between field wrappers in stacked and grid layouts, with no hidden minimum margin or transparent-wrapper padding. Fields with a visible background, border, conditional box style, or validation border retain padding inside that box. Horizontal grid spacing is unchanged.
- Date/Time fields support built-in date, date/time, and time-only output plus custom format strings. Custom formats can preserve default text case or force the final output to uppercase or lowercase.
- Lookup fields display their configured lookup values. Multi-value lookups render those display values as a comma-separated list rather than exposing SharePoint item IDs.
- The designer's **Conditional styling** workspace creates ordered field rules that apply background, text, border, corner, and font styles or set visibility. Rule values follow the source-field type: lookup items, configured choices, True/False, numbers, and dates use native selectors instead of free text. Date rules also support an inclusive **Between** range, and lookup rules compare selected item IDs including multi-value candidates. **Target fields** is a multi-select list, so one rule can affect several fields; hold Ctrl while clicking to select multiple fields. Use Move up/down to control precedence; the last matching rule wins for the same style or visibility property. Existing single-target rules and static values remain compatible. Hidden fields are removed from the layout so visible fields reflow without gaps.
- A custom **List Control** field can place a separately configured ListControl web part inside the report layout. The complete web-part host is moved at runtime and returned to its original page position when the report unmounts.
- SharePoint list and item permissions are still enforced by SharePoint. Report permission checks use `ViewListItems`, not Add or Edit permissions.

## Configure

1. Add **Report Forms** to a modern SharePoint page.
2. Select the report list in the property pane.
3. Open the visual designer from the web part while the page is in edit mode.
4. Add and arrange the fields that should appear in the report.
5. Select a fixed item ID, use the configured URL query parameter, or connect the dynamic item source to another supported web part.

To show the same report for the complete list, enable **Repeat report for every list item** in Basic settings. Repeat mode applies only at runtime; page edit mode continues to show one report and one visual designer. Fixed item ID, URL item ID, dynamic selection, and record-filter selection are ignored while repeating because the selected list supplies the record set.

### Navigate multiple reports

1. Deploy SPS Form Navigator 1.0.0.2 or later and add it to the same page as the reports.
2. Set a unique **Report name for navigation** on each Report Forms web part.
3. In Form Navigator, create categories and add each discovered report through **Manage forms**.
4. Publish the page and use the navigator to switch directly between reports.

Only Report Forms instances selected in the navigator are hidden or activated. All reports remain visible while the page is being edited. Navigation does not change a report's record-selection configuration, and Dynamic Forms and Report Forms can be included in the same navigator.

For a Date/Time field, choose **Custom** formatting to enter a token-based format, then select **Text case** as Default, UPPERCASE, or lowercase. The case option is applied after the complete custom date/time value is formatted.

### Embed a ListControl

1. Add and configure a **List Control** web part on the same page. Give it a unique **Web part name** when the page contains more than one ListControl.
2. Open the Report Forms visual designer and add **List Control** from **Custom fields**.
3. Edit the new custom field and select the named ListControl web part.
4. Set its column span or move it to the required report position, then save the designer and publish the page.

The ListControl remains an independently configured web part. Report Forms stores its dynamic source ID in the form schema and moves that web part into the custom field when the report is displayed. It sends the current parent value as a transient ListControl filter and republishes that configuration when the embedded control requests it after mounting. If the configured source field is unavailable on the loaded parent item but represents the same linked target used to resolve the report, Report Forms falls back to the incoming dynamic or URL value.

When using several Report Forms instances, configure each instance with its own list and dynamic item mapping. Page composition provides the combined report.

## Build

```powershell
npm install
npx gulp bundle --ship
npx gulp package-solution --ship
```

The deployable package is generated at `sharepoint/solution/sps-report-forms.sppkg`.

## Troubleshooting script-resource errors

Errors for `e47c4f0e-0f0d-49d5-bcab-7a4d05742037_<version>/SharePointDynamicFormWebPartStrings` or `/PropertyControlStrings` occur before Report Forms reads list data. They indicate that SharePoint could not load an SPFx localization asset, not that the user lacks access to the report list.

Property-pane version `1.0.0.62` and loader component version `1.0.37` identify the same outdated Report Forms package. The current package is solution `1.0.0.71` with component `1.0.46`.

1. Replace `sps-report-forms.sppkg` in the tenant App Catalog and approve the replacement/deployment.
2. Confirm the App Catalog displays solution version `1.0.0.71`.
3. Apply any available Report Forms update under the target site's **Site contents**.
4. Clear site data for the SharePoint tenant, including its service worker/cache, close all tenant tabs, and retest in a new InPrivate window.
5. In browser developer tools, inspect the failed localization `.js` request. A `404` indicates a stale manifest requesting a removed hashed filename. A `401` or `403` requires a SharePoint administrator to restore the tenant App Catalog and `ClientSideAssets` library to the organization's standard inherited/read-access configuration.

The historical `1.0.37` package contained its localization files, but their generated hashes differ from the current release. A browser retaining the old component manifest can continue requesting obsolete filenames after an App Catalog package upgrade.

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
- **Runtime:** the web part is intentionally read-only and does not expose submit or edit transitions. Date/time fields support custom patterns and final-output text case.
- **Conditional presentation:** ordered style and visibility rules react to report values; later matching rules take precedence and hidden fields reflow the layout.
- **Embedded content:** a designer-configured List Control can show records related to the current parent item using a transient filter that survives embedded-control remounts.
- **Layout and diagnostics:** full-width/fixed-width settings adapt the report to the page; dynamic diagnostics log item resolution, data loading, and embedded-control discovery.

Maintain the schema through the designer and use SharePoint permissions to secure report data. Hiding a field in the layout does not change list or item access.

## Common scenarios

- Select an incident, request, or case in List Control and show a formatted detail report beside it.
- Select a Calendar event and display a richer read-only event record.
- Open a report page directly with an item ID query parameter.
- Show a parent record with an embedded list of related actions, notes, or child items.
- Combine Report Forms with Print Control for a clean printable record.
