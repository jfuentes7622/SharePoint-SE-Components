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

## Display And Actions

- `pageSize` controls client-side pagination; `0` or blank displays all rows.
- The view selector, Refresh, Add, Edit, View, Delete, and Link to Item controls can be shown or hidden independently.
- Link to Item can navigate to a configured page, pass the selected ID through a configurable query parameter, and optionally include a return URL.
- Body, header, selected-row, alternating-row, table, button, and web-part-container styles are configurable in the property pane.
- Preset filters and conditional formatting rules can be built in their property-pane designers and stored as JSON.
- Diagnostic logging can be enabled for data loading and runtime troubleshooting.

## Integration with SharePointDynamicForm

Bind the form web part's dynamic item ID setting to this web part's `Selected Item ID` dynamic property.

The list control also publishes `Selected Mode` for future integration, but the current form web part primarily consumes the dynamic item ID.

## Commands

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.