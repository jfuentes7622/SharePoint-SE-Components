# GridControl

SPFx 1.5.1 editable SharePoint grid for SharePoint Server Subscription Edition.

GridControl starts with the ListControl display, filtering, styling, and dynamic-data baseline, then adds inline row creation and editing. Select **+ Add row** to insert a draft row, or use the row-level **Edit** button for an existing row. **Save** validates and creates or updates the SharePoint item; **Cancel** discards the draft.

## Purpose

- Connect to a SharePoint list
- Choose a SharePoint view
- Design, rename, resize, and reorder editable grid columns visually
- Render rows using the view output from `RenderListDataAsStream`
- Apply preset filters and conditional row or column styling
- Publish selected item ID and selected mode via Dynamic Data
- Add and edit SharePoint items directly in the grid
- Validate required fields and deterministic Dynamic Form field rules before saving
- Select multiple rows and delete them in one operation
- Restrict read access and add/edit/delete access by SharePoint group

## Grid Designer

Select a SharePoint list in the web-part properties, then choose **Open Grid Designer**. The designer provides three working panes:

1. **Available SharePoint fields** lists fields that can be added to the grid.
2. **Grid Columns** controls displayed column order and selection. Add, move, or remove columns here.
3. **Column Settings** configures the selected column's label, description, editor, width, state, default, and validation.

Choose **Save Grid** to apply the design or **Cancel** to discard the current changes. Changing the selected list clears the saved design because SharePoint field identities are list-specific. Before a design is saved, GridControl uses fields from the selected SharePoint view as a display fallback.

Fields marked required by SharePoint remain required in Grid Designer. Their Required checkbox is read-only and cannot be cleared; GridControl also enforces the list requirement at runtime for older saved designs.

Field help is displayed on hover over the column header, cells, and editing control. The SharePoint list-column description takes precedence when present. The Grid Designer description is used only when the list column does not have a description.

The designer can configure:

- `fieldName`, `type`, `label`, and `description`
- Boolean `visible`, `required`, and `readOnly` values
- `defaultValue` and `requiredMessage`
- `gridWidth` using values such as `180px`, `20%`, or `12rem`
- `config.placeholder`, `maxLength`, `min`, `max`, `decimals`, `choices`, and DateTime `displayFormat`
- Type-aware field validation rules and Dynamic Form-compatible advanced expressions

GridControl stores the result internally as a SharePoint Dynamic Form-compatible `FormSchema`. It reads fields from `steps[].fields` and intentionally ignores form steps, containers, wizard behavior, and form layout. The stored schema remains compatible with configurations created before the visual designer was introduced; raw JSON is not exposed as a normal property-pane authoring surface.

Grid Designer exposes one **Read only** setting because read-only and disabled cells have the same result in an inline grid. Legacy schemas with `disabled: true` remain supported and are normalized to read-only when opened in Grid Designer.

## Editable Controls

| SharePoint/Dynamic Form type | Grid editor |
|---|---|
| Text / `text` | Single-line text input |
| Note / `multiline` | Resizable multiline text area, or a visual HTML editor when SharePoint Rich Text is enabled |
| Number, Currency, Integer / `number` | Numeric input with optional min, max, and decimal step |
| Boolean / `boolean` | Checkbox |
| Choice / `dropdown` | Single-select dropdown |
| MultiChoice / `multiselect` | Multi-select list |
| DateTime / `datetime` | Date, date-time, or time input |
| URL / `url` | URL input |

Grid Designer restricts the Cell control list according to the underlying SharePoint storage type. Text and Note columns can use text, multiline, or number controls; URL columns can use URL or text controls. Number, Currency, Integer, Boolean, Choice, MultiChoice, DateTime, Lookup, Person, Taxonomy, Image, and Attachment columns remain limited to their storage-compatible control family. Existing schemas containing an incompatible selection are normalized when opened, and runtime applies the same guard to older or manually edited schemas.

Read-only, hidden, computed, and unsupported complex fields do not receive an editor. Lookup, person, taxonomy, image, and attachment editing controls remain reserved for a later implementation phase.

SharePoint Note columns configured for rich text use a visual editor with bold, italic, underline, lists, alignment, and clear-formatting controls. Existing HTML is rendered visually and preserved when saved instead of being exposed as tags. Rich and plain multiline editing surfaces can be resized horizontally and vertically.

## Validation

Simple validation options follow the selected control type:

- Text, multiline, and URL controls support minimum length, maximum length, and an optional regular-expression pattern.
- Number controls support minimum and maximum values.
- Required state and its message apply to all editable controls.
- Invalid regular expressions are shown in Grid Designer and prevent the design from being saved.

Advanced validation uses the same expression grammar and `advancedValidation` schema shape as SharePoint Dynamic Form. Rules can compare fields, target a specific cell message, or show a grid-level message. Supported operators are `==`, `!=`, `>`, `>=`, `<`, `<=`, `&&`, `||`, `!`, and parentheses. Supported functions are:

- `field("InternalName")`
- `today()` and `now()`
- `daysFromToday(n)`
- `isEmpty(value)` and `hasValue(value)`
- `between(value, minimum, maximum)`

GridControl also executes DForm-compatible field rule types `required`, `minLength`, `maxLength`, `min`, `max`, `pattern`, and `custom`, including `applyWhen` expressions. Invalid expressions and patterns fail validation visibly instead of being silently ignored.

## Display And Actions

- `pageSize` controls client-side pagination; `0` or blank uses Automatic (50), and the maximum is 100 rows per page.
- SharePoint rows are cached progressively using paged batches (500 by default, configurable from 100 to 2000). The next batch is prefetched near the cache boundary, page totals show `+` while more rows exist, and client-only filtering, sorting, or grouping completes the cache before totals become final.
- The interactive column header follows page scrolling while the grid is visible, stops at the bottom of the table, and keeps open filter dialogs anchored to their header buttons; horizontal scrolling remains synchronized with the columns.
- When the table is wider than its viewport, left/right navigation arrows appear on hover or keyboard focus and move with the visible portion of the grid.
- The view selector, Refresh, Add, Delete, and Link to Item controls can be shown or hidden independently.
- The row action column containing Edit, History, or Save/Cancel can be placed at the beginning or end of the data columns. Delete-selection checkboxes remain first when enabled.
- Runtime command buttons can display text, a standard SharePoint Fabric icon with text, or an icon only. Icon-only controls retain accessible labels and hover tooltips.
- When versioning is enabled for the selected SharePoint list, the optional row-level **History** button opens the default version-history page in a padded in-page dialog. The button is hidden when list versioning is off, and its property-pane option is disabled with an explanatory note. It remains available in read-only grids because viewing history does not mutate the item; SharePoint permissions still govern access to version details.
- Link to Item can navigate to the list's default display form or an `.aspx` page discovered from the current site's Site Pages and publishing Pages libraries. It passes the selected ID through a configurable query parameter and can optionally include a return URL. Previously saved custom target URLs remain available in the dropdown.
- Body, header, selected-row, alternating-row, table, button, and web-part-container styles are configurable in the property pane.
- Preset filters and conditional formatting rules can be built in their property-pane designers and stored as JSON. Conditional overrides include background, foreground, font, alignment, border color/type/thickness, and square or rounded corners with a configurable radius.
- Diagnostic logging can be enabled for data loading and runtime troubleshooting.
- The Security settings can restrict reading to one SharePoint group and mutations to another. Edit-group members automatically receive read access. Users without read access see no list data, while read-only users see the configured edit-denial message and no mutation controls. Blank group selections defer to normal SharePoint list permissions.
- Use a row's **Edit** button to edit it inline. Save and Cancel replace the row action while editing, and only one row can be edited at a time.
- Select rows with the checkbox column, or select all rows on the current page from the header, then choose **Delete selected**. Selections persist while paging; failed deletions remain selected for retry.

## Integration with SharePointDynamicForm

Bind the form web part's dynamic item ID setting to this web part's `Selected Item ID` dynamic property. Dynamic Form can use that value as the form item ID or assign it to a selected parent lookup column.

Bind Dynamic Form's dynamic mode setting to `Selected Mode` so adding, selecting, or editing a row switches the connected form to `new`, `view`, or `edit` mode. The target form list must contain the selected parent lookup column when the item ID is used to create or filter related records instead of opening the same item directly.

## Commands

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.

## Usage

1. Add **SPS Grid Control**, select a SharePoint list and view, and open **Grid Designer**.
2. Add and order fields, then configure each column's label, width, editor, visibility, required/read-only state, default, and validation.
3. Configure page size, filters, row actions, target-page behavior, and conditional styles.
4. Test add, edit, cancel, validation, and multi-row deletion with realistic field types and permissions before publishing.
5. Optionally connect the published selected item ID/mode to Dynamic Forms or Report Forms.

## Properties and common configuration

- **Data:** `listName`, `viewId`, and fallback `viewColumns` define the source and initial columns.
- **Selected views:** Grid Control executes the selected view's native CAML filter, sort, scope, and row limit first, then loads Grid Designer field values only for the item IDs returned by that view.
- **Grid design:** `gridSchemaJson` stores the visual designer output and is authoritative once saved. Use the designer for field compatibility and validation rather than hand-editing it.
- **Behavior:** `pageSize`, `fetchBatchSize`, view selector, refresh, add/delete/history, item links, filtering, and target-page properties control grid operation. Blank page size means Automatic (50); fetch batches default to 500.
- **Editors and validation:** text, multiline, number, Boolean, choice, multi-choice, date/time, URL, lookup, and person fields receive compatible controls; list-required fields remain required.
- **Appearance:** body/header/selected-row, borders, alternate rows, buttons, date/time formats, web-part surface, and conditional-rule designers control display.
- **Advanced:** `forceFullWidth` uses the available page width. `enableDiagnostics` records metadata, REST fallback, mutation, and rendering details.

The current user needs SharePoint permissions for each enabled action. Unsupported custom field types can be displayed but should remain read-only.

## Common scenarios

- Inline maintenance of tasks, inventory, contacts, or issue records.
- Parent/child line-item editing inside Dynamic Forms after the parent item is saved.
- Bulk selection and deletion for a controlled operational list.
- Conditional highlighting of overdue, high-priority, invalid, or completed rows.
- A selector/editor paired with a read-only Report Form on the same page.