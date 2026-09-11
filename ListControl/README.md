# ListControl

SPFx 1.5.1 list control web part for SharePoint Server Subscription Edition.

## Purpose

- Connect to a SharePoint list
- Choose a SharePoint view
- Select, rename, resize, and reorder the displayed view columns
- Render rows using the view output from `RenderListDataAsStream`
- Apply preset filters and conditional row or column styling
- Publish selected item ID and selected mode via Dynamic Data
- Support `New`, `Edit`, `View`, and `Delete` actions

## View Columns

Selecting a list and view initializes the `viewColumns` collection from the SharePoint view's field order. Each row contains:

| Field | Description |
|---|---|
| `fieldName` | SharePoint internal field name used to read the value. |
| `displayName` | Column heading shown by ListControl. |
| `width` | Optional CSS width such as `180px`, `20%`, or `12rem`. Leave blank for automatic sizing. |

Use Move up and Move down in the collection editor to control display order. Changing the selected list or view clears and rebuilds the collection from that view, while later edits remain independent of SharePoint's view-column order.

If the collection is empty, ListControl falls back to the selected SharePoint view's fields and order.

Lookup values are normalized to their SharePoint display text. Object or JSON-shaped values containing `lookupValue`/`LookupValue` are shown as the readable lookup label rather than raw JSON.

Person/User columns, including **Modified By** (`Editor`) and **Created By** (`Author`), prefer the SharePoint display name or title across uppercase and lowercase response variants. Account name, email, and numeric ID are used only when no display label is available, preventing expanded user metadata from appearing in cells, sorting, grouping, or filters.

## Display And Actions

- `pageSize` controls client-side pagination; `0` or blank uses Automatic (50), and the maximum is 100 rows per page.
- SharePoint rows are cached progressively using paged batches (500 by default, configurable from 100 to 2000). The next batch is prefetched near the cache boundary, page totals show `+` while more rows exist, and client-only filtering, sorting, or grouping completes the cache before totals become final.
- The interactive column header follows page scrolling while the list is visible, stops at the bottom of the table, and keeps open filter dialogs anchored to their header buttons; horizontal scrolling remains synchronized with the columns.
- When the table is wider than its viewport, left/right navigation arrows appear on hover or keyboard focus and move with the visible portion of the list.
- The view selector, Refresh, Add, Edit, View, Delete, and Link to Item controls can be shown or hidden independently.
- Link to Item can navigate to the list's default display form or an `.aspx` page discovered from the current site's Site Pages and publishing Pages libraries. It passes the selected ID through a configurable query parameter, can optionally include a return URL, and can open in the same tab, a new tab, a sized new window, or a resizable in-page dialog. The dialog closes from its Close button or embedded callback; releasing a native resize handle does not dismiss it. Previously saved custom target URLs remain available in the dropdown.
- Body, header, selected-row, alternating-row, table, button, and web-part-container styles are configurable in the property pane.
- Preset filters and conditional formatting rules can be built in their property-pane designers and stored as JSON. Conditional overrides include background, foreground, font, alignment, border color/type/thickness, and square or rounded corners with a configurable radius.
- Diagnostic logging can be enabled for data loading and runtime troubleshooting.

## Integration with SharePointDynamicForm

Bind the form web part's dynamic item ID setting to this web part's `Selected Item ID` dynamic property. Dynamic Form can use that value as the form item ID or assign it to a selected parent lookup column.

Bind Dynamic Form's dynamic mode setting to `Selected Mode` so the Add, Edit, and View actions switch the connected form to `new`, `edit`, or `view` mode. When the selected value creates or filters related records instead of opening the same SharePoint item directly, configure a supported writable target column in the form; lookup/person targets use the selected numeric item ID.

When embedded in Report Forms, ListControl accepts a transient parent filter without changing its saved web-part properties. Filter target columns are hydrated by item ID even when they are absent from the selected view, and remain hidden unless they were configured for display. On mount or remount, ListControl requests the current runtime configuration so the host can immediately restore the filter.

## Commands

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.

## Usage

1. Add **SPS List Control**, select the source list and default view, and review the generated view-column collection.
2. Reorder columns, rename labels, and assign widths only where automatic sizing is insufficient.
3. Configure page size, preset filters, visible actions, item links, and conditional styles.
4. Optionally connect the selected item ID and requested New/Edit/View mode to Dynamic Forms or Report Forms.
5. Test every visible action with representative users before publishing.

## Properties and common configuration

- **Data:** `instanceName`, `listName`, `viewId`, `viewColumns`, `pageSize`, and `fetchBatchSize` define the source, identity, columns, visible page size, and progressive SharePoint cache. Blank page size means Automatic (50); fetch batches default to 500.
- **Navigation:** view-selector, item-link, target-page, item-ID parameter, return-URL, and same-tab/new-tab/new-window/dialog target properties control how users move between views and records.
- **Actions:** refresh, add, edit, view, and delete toggles simplify the toolbar for the intended workflow; permissions still come from SharePoint.
- **Rules:** filter and conditional-style designers store JSON definitions with field operators, date/current-user expressions, and row/column scopes.
- **Appearance:** body/header/selected/alternate rows, table borders, web-part surface, buttons, typography, alignment, and date/time formats are configurable.
- **Advanced:** `forceFullWidth` expands the available workspace; embedded runtime filters can use hidden support columns and recover after remounts; `enableDiagnostics` logs list, view, filter, dynamic-data, and action details.

Selecting a different list or view can rebuild the column collection. Review custom labels, ordering, and widths after changing either source setting.

## Common scenarios

- A searchable operational list with only View and Refresh actions.
- A record selector that drives Dynamic Forms in New, Edit, or View mode.
- A dashboard table that highlights status, deadlines, ownership, or exceptions.
- A paged list linked to a custom details/report page through an item ID query parameter.
- A wide, read-only results table paired with Full Width Control.