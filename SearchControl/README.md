# SPS Search Control

SPS Search Control connects to an SPS List Control or Grid Control through SPFx Dynamic Data. The target publishes its list and typed field metadata; Search Control publishes an addressed filter state that only the selected target applies.

## October 26, 2026 theme-loading fix

- npm version: `1.0.15`; SharePoint solution and feature version: `1.0.0.15`.
- Early theme notifications cache the theme without rendering or refreshing the
  property pane before the first normal render. Later theme changes retain both
  updates; search behavior and Dynamic Data connections are unchanged.
- Redeploy `sharepoint/solution/sps-search-control.sppkg`. Run
  `node tests\themeLifecycle.test.js` from the repository root.

## October 14, 2026 release

- npm version: `1.0.14`
- SharePoint solution and feature version: `1.0.0.14`
- Package: `sharepoint/solution/sps-search-control.sppkg`
- Adds **Manual styling**, **Site theme only**, and **Site theme with overrides**.
- Existing instances remain in Manual styling unless an author selects a theme mode.
- Site theme modes cover search surfaces, text, borders, buttons, selected states, and font family without changing search behavior or Dynamic Data connections.

## Search modes

When mode switching is enabled, **Simple** and **Advanced** are shown as radio buttons in the top-right corner of the control.

- **Simple** initially shows one search box, a Search button, and a **Show fields** checkbox. Select **Show fields** to display
  the optional field selector and search one selected field instead of all compatible fields.
  Clearing the search box with its built-in **X** immediately removes the active search filter.
- **Advanced** supports multiple typed conditions joined with AND or OR. The Connector and Operator controls use compact columns so that more horizontal space remains available for the Field and Value controls.

Choice and Boolean fields use fixed options, lookup and person options are loaded on demand, and date/time fields use typed browser controls.

## Buttons

The **Buttons > Button display** property controls how Search Control command buttons are rendered:

- **Icons** displays icon-only buttons and is the default.
- **Icons and names** displays an icon followed by the button name.
- **Names** displays the button name without an icon.

This setting applies to Search, Clear, Add condition, Apply filters, and Remove. Icon-only buttons retain accessible labels and hover tooltips.

## Configuration

| Property | Description |
|----------|-------------|
| Target control | Selects the SPS List Control or Grid Control that receives the published search state. |
| Search data source | Uses either the records already loaded by the target control or a server query against the complete SharePoint source list. |
| Search fields | Selects the fields shown in the simple field selector and used by **All compatible fields**. An empty selection uses all available fields. |
| Default mode | Selects Simple or Advanced as the initial search mode. |
| Allow users to switch modes | Shows or hides the top-right Simple and Advanced radio buttons. |
| Title | Sets the heading displayed in the control. |
| Search placeholder | Sets the placeholder used by the simple text search. |
| Button display | Displays command buttons as Icons, Icons and names, or Names. The default is Icons. |
| Search while typing | Automatically publishes an all-fields simple search while the user types. |
| Minimum characters | Sets the number of characters required before search-while-typing runs. |
| Typing delay | Sets the debounce delay for search-while-typing. |
| Time format | Selects 12-hour or 24-hour time entry. |
| Minute increment | Selects 1, 5, 10, or 15-minute increments. |

The web part also includes diagnostics, version information, tab exclusion, colors, border settings, typography, and optional override CSS.

Server search queries the complete source list through the connected control and pages only matching
records. The **Ends with** and **Does not contain** operators remain available in Local mode but are
not supported in Server mode because SharePoint CAML does not provide equivalent text operators.
