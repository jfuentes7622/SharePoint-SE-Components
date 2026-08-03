# ListControl

SPFx 1.5.1 list control web part for SharePoint Server Subscription Edition.

## Purpose

- Connect to a SharePoint list
- Choose a SharePoint view
- Render rows using the view output from `RenderListDataAsStream`
- Publish selected item ID and selected mode via Dynamic Data
- Support `New`, `Edit`, `View`, and `Delete` actions

## Integration with SharePointDynamicForm

Bind the form web part's dynamic item ID setting to this web part's `Selected Item ID` dynamic property.

The list control also publishes `Selected Mode` for future integration, but the current form web part primarily consumes the dynamic item ID.

## Commands

- `npm install`
- `npm run build`
- `npm run start`