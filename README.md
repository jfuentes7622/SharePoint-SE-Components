# SharePoint-SE-Components

A collection of SharePoint Framework (SPFx) components, built for SharePoint SE (SPFx v1.5.1). This repository contains reusable web parts (and one application customizer extension) that can be deployed to SharePoint Online/Server and, where noted, Microsoft Teams.

All web-part property panes display the deployed solution version. Components with an **Enable diagnostics logging** setting write optional lifecycle and data-loading details to the browser console; genuine failures remain visible even when diagnostic logging is disabled.

## Components

| Component | Type | Summary | Detailed guide |
|---|---|---|---|
| [Accordion](#accordion) | Web Part | Collapsible list-driven content sections | [Accordion/README.md](Accordion/README.md) |
| [BannerClock](#bannerclock) | Application Customizer | Centrally configured world-clock banner | [BannerClock/README.md](BannerClock/README.md) |
| [Calendar](#calendar) | Web Part | Multi-source calendar with recurrence, rules, details, and dynamic data | [Calendar/README.md](Calendar/README.md) |
| [Carousel](#carousel) | Web Part | Auto-playing image carousel from a library | [Carousel/README.md](Carousel/README.md) |
| [Contact](#contact) | Web Part | Filterable staff/contact directory cards | [Contact/README.md](Contact/README.md) |
| [Dynamic Forms](#sharepointdform-dynamic-forms) | Web Part | No-code list form designer and runtime | [DynamicForms/README.md](DynamicForms/README.md) |
| [Full Width Control](#full-width-control) | Web Part | Expands the containing page section | [FullWidthControl/README.md](FullWidthControl/README.md) |
| [GridControl](#gridcontrol) | Web Part | Designer-configured editable list grid | [GridControl/README.md](GridControl/README.md) |
| [LinkButton](#linkbutton) | Web Part | Styled hyperlink button with optional image | [LinkButton/README.md](LinkButton/README.md) |
| [ListControl](#listcontrol) | Web Part | Sortable/filterable list table and item selector | [ListControl/README.md](ListControl/README.md) |
| [Marquee](#marquee) | Web Part | Static or list-driven scrolling announcements | [Marquee/README.md](Marquee/README.md) |
| [PrintControl](#print-control) | Web Part | Prints page content without SharePoint chrome | [PrintControl/README.md](PrintControl/README.md) |
| [Report Forms](#report-forms) | Web Part | Read-only report for a selected list item | [ReportForms/README.md](ReportForms/README.md) |
| [ScriptEditor](#script-editor) | Web Part | Trusted custom HTML and JavaScript host | [ScriptEditor/README.md](ScriptEditor/README.md) |
| [Tab](#tab) | Web Part | Tabbed organization for sections or web parts | [Tab/README.md](Tab/README.md) |
| [Tiles](#tiles) | Web Part | Responsive clickable navigation tiles | [Tiles/README.md](Tiles/README.md) |

---

## Accordion

Displays collapsible accordion sections populated dynamically from a SharePoint list. Each list item becomes an accordion section, with the header and body content mapped from selected list columns. Content is rendered as HTML (surrounding paragraph tags are stripped). Supports two interaction modes: single-selection (radio-button behavior, only one section open at a time) and multiple-selection (checkbox behavior, multiple sections open simultaneously).

### Properties

**Data Source & Mapping**
| Property | Label | Description |
|---|---|---|
| `listName` | List | SharePoint list to pull accordion data from (populated with non-hidden lists) |
| `itemName` | Header | List column used as the header/title of each accordion section |
| `itemContent` | Content | List column used as the body/content of each accordion section |

**Behavior**
| Property | Label | Description |
|---|---|---|
| `optionChoice` | Accordion Style | Interaction mode: "Single" (radio buttons, one section open at a time) or "Multiple" (checkboxes, several open at once) |

**Header Styling**
| Property | Label | Description |
|---|---|---|
| `headerBackgroundColor` | Header Background Color | Background color applied to each accordion section's header/label |
| `headerTextColor` | Header Text Color | Text color applied to each accordion section's header/label |
| `headerFontBold` | Header Bold Text | Toggles bold font weight on the header text |

**Content Styling**
| Property | Label | Description |
|---|---|---|
| `contentBackgroundColor` | Content Background Color | Background color applied to each accordion section's content panel |
| `contentTextColor` | Content Text Color | Text color applied to each accordion section's content panel |
| `contentFontBold` | Content Bold Text | Toggles bold font weight on the content text |

**Font Settings**
| Property | Label | Description |
|---|---|---|
| `fontFamily` | Font Family | Font family applied to both header and content text (Segoe UI, Arial, Trebuchet MS, Georgia, Times New Roman, Courier New) |
| `fontStyle` | Font Style | Font style applied to both header and content text (Normal, Italic, Oblique) |

**Advanced**
| Property | Label | Description |
|---|---|---|
| `overrideCssUrl` | Override CSS URL | Optional external stylesheet URL loaded after the web part's own CSS to override component styles |
| `enableDiagnostics` | Enable diagnostics logging | Enables browser-console details for configuration and list loading |

---

## BannerClock

An SPFx **Application Customizer** (extension), not a web part — display name **KM World Clock**. Injects a world-clock banner showing live, auto-updating time for one or more timezones/locations across the top (and optionally bottom) of classic and modern SharePoint pages. Clocks can render in digital and/or analog format. Configuration (timezones, format, header/footer content) is data-driven from a SharePoint list and JSON settings files in Site Assets, so the display can be updated centrally without redeploying code.

### Properties

**Extension**
| Property | Description |
|---|---|
| `rootUrl` | Root site URL where the central configuration files are stored, enabling tenant-wide configuration from one location |

**Display & Format** (from `settings.txt` / `ISettings`)
| Property | Description |
|---|---|
| `digitalOnly` | When `true`, shows only the digital clock; when `false`, shows both digital and analog clocks |
| `hour12` | When `true`, uses 12-hour time (AM/PM); when `false`, uses 24-hour time |
| `displayDay` | When `true`, shows the weekday name alongside the time |

**Data**
| Property | Description |
|---|---|
| `spList` | Name of the SharePoint list containing timezone/location data (requires `Title` and `Timezone` columns) |

**Page Element Placeholders** (target DOM element IDs for injection)
| Property | Description |
|---|---|
| `ClassicBannerId` | Element ID where the banner renders on classic SharePoint sites |
| `ClassicFooterId` | Element ID where the footer renders on classic SharePoint sites |
| `ModernBannerComId` | Element ID for the banner on modern communication sites |
| `ModernFooterComId` | Element ID for the footer on modern communication sites |
| `ModernBannerTeamId` | Element ID for the banner on modern team sites |
| `ModernFooterTeamId` | Element ID for the footer on modern team sites |

---

## Calendar

Displays events from multiple SharePoint Events lists or lists with calendar views in a full interactive calendar powered by the free, MIT-licensed FullCalendar library. It supports month, week, day, and agenda-style list views; SharePoint recurring events; field-based swim lanes; event filtering and conditional styling; an in-place details panel; configurable target pages; and dynamic-data connections to SharePoint Dynamic Form.

### Properties

| Property | Label | Description |
|---|---|---|
| `dataSources` | Manage data sources | Ordered Events/calendar-view lists with date mappings, source colors, and target pages |
| `listName` | Events list | Legacy single-list fallback retained for existing configurations |
| `defaultView` | Default view | `dayGridMonth`, `timeGridWeek`, `timeGridDay`, or `listWeek` |
| `showWeekends` | Show weekends | Shows or hides Saturday and Sunday |
| `calendarHeight` | Calendar height | Calendar height in pixels |
| `enableSwimlanes` | Use swim lanes | Groups events into horizontal lanes using a selected field and a 7, 14, or 30 day range |
| `filterJson` | Event filters | Conditions created in the filter designer, with static values or user/date expressions |
| `conditionalStyleJson` | Conditional event styles | Priority-based color and typography rules, optionally scoped to one source |
| `enableEventDetails` | Show event details | Opens a configurable in-place details panel when an event is selected |
| `showLinkToItem` | Open the selected event | Opens the default display form or a discovered Site Pages/publishing Pages target |
| `enableDiagnostics` | Enable diagnostics logging | Enables browser-console details for calendar and event loading |

Each data source can use its own background color. Conditional rules override source styling, and selected-event styling has the highest priority. Calendar publishes the selected SharePoint item ID and `view` mode so a Dynamic Form web part can display the selected event without leaving the page. See [Calendar/README.md](Calendar/README.md) for recurrence behavior, details-field configuration, filter expressions, conditional-style JSON, and appearance controls.

---

## Carousel

Displays a full-width, auto-playing image slider sourced from a SharePoint document or picture library. Images are sorted by an optional slide order (falling back to library item order), and support optional titles, click-through links, and date-based expiration/visibility columns — none of which are required, so the web part works against a plain library with only image files. Features auto-play with pause-on-hover, configurable slide/transition timing, adjustable dimensions, circular or square image display, and responsive image scaling.

### Properties

| Property | Label | Description |
|---|---|---|
| `carouselSlideLibrary` | Select Carousel Library | Document library that contains the carousel slide images (required) |
| `carouselWidth` | Carousel Width | Width of the carousel container in pixels (default: 500) |
| `carouselHeight` | Carousel Height | Height of the carousel container in pixels; `0` enables dynamic height based on image aspect ratio (default: 0) |
| `carouselBackgroundColor` | Carousel Background Color | Background fill color of the carousel (supports alpha transparency) |
| `carouselSlideInterval` | Slide Interval | How long each slide displays before auto-advancing, in milliseconds (default: 5000) |
| `carouselTransitionInterval` | Transition Interval | Duration of the fade/slide transition animation between slides, in milliseconds (default: 1500) |
| `imageIsCircle` | Display images in a circle | Renders slide images as circles instead of rectangles |
| `enableDiagnostics` | Enable diagnostics logging | Enables browser-console details for image loading and carousel behavior |

---

## Contact

Displays a filterable, customizable directory of personnel/contacts organized by organizational hierarchy (directorate, division, branch, group). Supports two data sources: a SharePoint list, or inline custom contact entries authored directly in the property pane. Renders contact cards with configurable photo shape (square, rounded, or circle), job title, name, email, phone, VOIP info, and biography links, with configurable header/title colors, fonts, and styling.

### Properties

**List Configuration**
| Property | Label | Description |
|---|---|---|
| `propertyCheckbox` | Use Existing List | Toggle between a SharePoint list and inline custom contact data |
| `ContactsListId` | Select a list | SharePoint list containing contact data (shown when "Use Existing List" is checked) |
| `customList` | Enter Contacts data with BIO Links and Pictures | Inline table of contacts: Job Title, Full Name and Rank, Email, Phone Number, Link to Biography, Link to Image, Display Photo, VOIP, Group, Display Order (shown when "Use Existing List" is unchecked) |

**Column Mapping** (when using a SharePoint list)
| Property | Label | Description |
|---|---|---|
| `titleField` | Select Job Title Column | List column mapped to Job Title |
| `directorateField` | Select Directorate Column | List column mapped to Directorate (used for filtering) |
| `divisionField` | Select Division Column | List column mapped to Division (used for hierarchical filtering) |

**Filtering & Selection**
| Property | Label | Description |
|---|---|---|
| `directorate` | Directorate | Filter contacts by selected directorate |
| `division` | Division | Filter contacts by division within the selected directorate |

**Appearance & Styling**
| Property | Label | Description |
|---|---|---|
| `PersonnelPanelHeaderBackColor` | Panel Header Background Color | Contact card header background color (default: `#8A1717`) |
| `PersonnelPanelHeaderTextColor` | Panel Header Text Color | Contact card header text color (default: `#FAFAFA`) |
| `headerFontFamily` | Header Font Family | Font family used for the contact card header text |
| `headerFontStyle` | Header Font Style | Font style for the header text: `normal`, `italic`, or `oblique` |
| `headerFontBold` | Bold Header Text | Renders the header text in bold |
| `headerAlignment` | Header Text Alignment | Header text alignment: `left`, `center`, or `right` |
| `headerTopCorners` | Header Top Corners | Header panel top corners: `rounded` or `squared` |
| `tileTitleColor` | Tile Title Color | Job title text color on each contact tile |
| `titleFontFamily` | Title Font Family | Font family used for the tile job title text |
| `titleFontStyle` | Title Font Style | Font style for the tile job title text |
| `titleFontBold` | Bold Title Text | Renders the tile job title text in bold |
| `tileInfoBackgroundColor` | Tile Info Background Color | Background color behind the name/email/phone info area of each contact tile |
| `imageWidth` | Image Diameter (ex:150, no px or %) | Size of contact photos, in pixels |
| `imageShape` | Image Shape | Shape of contact photos: `square`, `rounded`, or `circle` |
| `overridecss` | Override CSS URL | URL to a custom stylesheet that overrides web part styling |
| `enableDiagnostics` | Enable diagnostics logging | Enables browser-console details for contact, record, and image loading |

---

## LinkButton

Renders a styled hyperlink displayed as a button (**SPS Link Button**). In addition to text, target URL, and alignment, it supports configurable shape, colors, typography, dimensions, an optional image icon, and an external override stylesheet. Client-side validation blocks HTML injection and limits button text to 50 characters.

### Properties

| Property | Label | Description |
|---|---|---|
| `description` | Button Title | Text displayed on the button (max 50 characters; validated against HTML injection) |
| `Link` | Link | Target URL the button navigates to when clicked (requires a valid URL) |
| `Align` | Align Button | Horizontal alignment of the button: `left`, `center`, or `right` |
| `buttonShape` | Button shape | Square or rounded corners (default: `rounded`) |
| `buttonBackgroundColor` / `buttonFontColor` | Button colors | Background and text colors |
| `fontFamily` / `fontStyle` / `fontBold` / `fontSize` | Typography | Font family, style, weight, and CSS size |
| `buttonHeight` / `buttonWidth` | Dimensions | CSS height and width (defaults: `5vh` and `200px`) |
| `showIcon` | Show icon | Displays an optional image icon |
| `iconImageUrl` / `iconSize` / `iconPosition` | Icon settings | Image URL, CSS size, and left/right placement |
| `overrideCssUrl` | Override CSS URL | Optional external stylesheet URL |
| `enableDiagnostics` | Enable diagnostics logging | Enables browser-console validation and rendering details |

> Note: Teams tab icon assets exist in the `teams/` folder, but there is no functional Teams tab implementation.

---

## GridControl

An independent SPFx solution derived from ListControl for editing SharePoint items directly in a grid. It preserves list/view selection, sorting, filtering, pagination, conditional styling, target-page discovery, and dynamic-data publication, and adds **+ Add row**, row-level inline editing with Save/Cancel, and checkbox-based bulk deletion. Text, multiline, numeric, Boolean, choice, multi-choice, date/time, and URL fields use type-appropriate controls and save through SharePoint REST.

Use **Open Grid Designer** in the property pane to add and order SharePoint fields and configure each column's label, width, storage-compatible editor, visibility, required/read-only state, defaults, and validation. List-required fields cannot be made optional. Column help appears on hover using the SharePoint field description first and the designer description only as a fallback. Control choices are restricted by the underlying SharePoint field type, while text-backed fields can safely opt into text, multiline, or number controls. Validation includes contextual length/range/pattern rules plus Dynamic Form-compatible cross-field expressions and helper functions. The three-pane designer stores a SharePoint Dynamic Form-compatible schema internally while intentionally omitting form containers, steps, and wizard layout. The selected view supplies fallback display columns until a grid design is saved. Dynamic Form also recognizes GridControl as a source for selected item ID and mode.

See [GridControl/README.md](GridControl/README.md) for the copied feature baseline and build commands.

---

## ListControl

Displays SharePoint list data as a sortable, filterable table (**SPS List Control**) with support for multiple list views, pagination, and consuming/producing SPFx dynamic data. It provides toolbar actions to create/view/edit/delete items (typically paired with the [SharePointDForm](#sharepointdform-dynamic-forms) web part for the actual add/edit UI), and supports conditional formatting rules that apply colors, fonts, and alignment to rows or specific columns based on field-value conditions. The selected item ID and requested New/Edit/View mode can be exposed as dynamic data for other web parts to consume.

### Properties

**Data Source & Display**
| Property | Label | Description |
|---|---|---|
| `instanceName` | Web part name | Identifies this List Control for Dynamic Form web part connections (leave blank for default) |
| `listName` | SharePoint list | List to display |
| `viewId` | Default view | View of the list to show initially |
| `viewColumns` | View columns | Ordered collection of internal field name, display name, and optional CSS width; use Move up/Move down to set display order |
| `pageSize` | Items per page | Number of items per page (`0` or blank shows all items) |

Selecting a list or view rebuilds `viewColumns` from the SharePoint view's field order. Blank widths use automatic sizing, and an empty collection falls back to the selected view's fields.

**View & Navigation**
| Property | Label | Description |
|---|---|---|
| `showViewSelector` | Show view label and dropdown | Show/hide the view selector dropdown in the toolbar |
| `linkTargetPageUrl` | Link target page | Default list display form or an `.aspx` page discovered from Site Pages and publishing Pages; saved custom URLs remain available |
| `linkTargetIdParam` | Link target ID parameter | Query parameter name used to pass the item ID (default: `itemid`) |
| `showLinkToItem` | Show Link to Item | Enable/disable linking to item details on another page |
| `includeReturnUrlParam` | Include return URL parameter | Adds a return URL parameter so the target page can navigate back |

**Button Visibility**
| Property | Label | Description |
|---|---|---|
| `showRefresh` | Show Refresh button | Show/hide the refresh button |
| `showAdd` | Show Add button | Show/hide the new item button |
| `showEdit` | Show Edit button | Show/hide the edit item button |
| `showView` | Show View button | Show/hide the view item details button |
| `showDelete` | Show Delete button | Show/hide the delete item button |

**Body Text Styling**
| Property | Label | Description |
|---|---|---|
| `bodyTextColor` | Body text color | Text color in table data cells |
| `bodyFontFamily` | Body font family | Font family for table data cells |
| `bodyFontSize` | Body font size | Font size for table data cells (12px–20px) |
| `bodyFontStyle` | Body font style | Font style for table data cells (normal, italic, oblique) |
| `bodyTextAlign` | Body text alignment | Text alignment in table data cells |
| `bodyFontBold` | Body text bold | Bold text in table data cells |

**Selected Row Styling**
| Property | Label | Description |
|---|---|---|
| `selectedTextColor` | Selected row text color | Text color for the selected row |
| `selectedBackgroundColor` | Selected row background color | Background color for the selected row |
| `selectedFontStyle` | Selected row font style | Font style applied to the selected row |
| `selectedFontBold` | Selected row text bold | Bold text for the selected row |

**Table Header Styling**
| Property | Label | Description |
|---|---|---|
| `headerTextColor` | Header text color | Column header text color |
| `headerBackgroundColor` | Header background color | Table header row background color |
| `headerFontFamily` | Header font family | Font family for column headers |
| `headerFontSize` | Header font size | Font size for column headers |
| `headerFontStyle` | Header font style | Font style for column headers |
| `headerTextAlign` | Header text alignment | Text alignment in column headers |
| `headerFontBold` | Header text bold | Bold column header text |

**Table Structure & Borders**
| Property | Label | Description |
|---|---|---|
| `tableBackgroundColor` | Table background color | Background color of the entire table |
| `tableBorderColor` | Table border color | Table border color |
| `tableBorderWidth` | Table border width | Thickness of table borders (0px–5px) |
| `tableRowLineWidth` | Table row line width | Thickness of horizontal lines between rows (0–10px) |
| `tableCornerStyle` | Table corner style | Square or rounded corners |
| `tableCornerRadius` | Table corner radius | Radius for rounded corners (0–40px) |
| `alternateRowShading` | Alternate row shading | Enable alternating row background colors |
| `alternateRowShadingColor` | Alternate row shading color | Background color for alternating rows |

**Web Part Container Styling**
| Property | Label | Description |
|---|---|---|
| `webpartBackgroundColor` | Web part background color | Background color of the web part container |
| `webpartBorderColor` | Web part border color | Border color of the web part container |
| `webpartBorderWidth` | Web part border width | Border width of the web part container (0–20px) |
| `cornerStyle` | Corner style | Square or rounded corners for the web part |
| `cornerRadius` | Corner radius | Radius for rounded corners on the web part (0–40px) |

**Button Styling**
| Property | Label | Description |
|---|---|---|
| `buttonTextColor` | Button text color | Text color on action buttons |
| `buttonBackgroundColor` | Button background color | Background color of action buttons |
| `buttonFontFamily` | Button font family | Font family for button labels |
| `buttonFontSize` | Button font size | Font size for button labels |
| `buttonFontStyle` | Button font style | Font style for button labels |
| `buttonFontBold` | Button text bold | Bold button text |
| `buttonCornerStyle` | Button corner style | Square or rounded button corners |
| `buttonCornerRadius` | Button corner radius | Radius for rounded button corners (0–40px) |

**Preset Filtering**
| Property | Label | Description |
|---|---|---|
| `filterJson` | Preset filter JSON | JSON array of filter conditions applied before user filters |
| `filterDesignerField` | Filter field | Field selected in the filter designer UI |
| `filterDesignerOperator` | Operator | Comparison operator (`eq`, `ne`, `contains`, `notcontains`, `startswith`, `endswith`, `gt`, `ge`, `lt`, `le`) |
| `filterDesignerLogical` | Logical | Logical connector between conditions (`AND`/`OR`) |
| `filterDesignerValueType` | Value type | Static value or expression (e.g., `today`, `me`, `date(-3)`) |
| `filterDesignerValue` | Value/Expression | The filter value or expression to compare against |
| `filterDesignerSelectedIndex` | Existing conditions | Condition to load, edit, or remove |
| `filterDesignerLookupPick` | Select item | Item picker for lookup-field filter conditions |

**Conditional Styling**
| Property | Label | Description |
|---|---|---|
| `conditionalStyleJson` | Conditional style JSON | JSON array of rules that style rows/columns based on conditions |
| `conditionalStyleRuleName` | Rule name | Identifier for the conditional style rule |
| `conditionalStyleScope` | Apply style to | Scope of the rule: row or column |
| `conditionalStyleApplyField` | Target column | Column to style (required when scope is column) |
| `conditionalStyleConditionField` | Condition field | Field evaluated for the styling condition |
| `conditionalStyleConditionOperator` | Condition operator | Comparison operator (`eq`, `ne`, `contains`, `notcontains`, `startswith`, `endswith`, `gt`, `ge`, `lt`, `le`) |
| `conditionalStyleConditionLogical` | Condition logical | Logical connector for multiple conditions (`AND`/`OR`) |
| `conditionalStyleConditionValueType` | Value type | Static value or expression |
| `conditionalStyleConditionValue` | Condition value/Expression | Value or expression to compare against |
| `conditionalStyleEnabled` | Enable this rule | Enable/disable the rule without deleting it |
| `conditionalStylePriority` | Priority | Rule priority (lower number = higher priority) |
| `conditionalStyleSelectedIndex` | Existing conditions | Rule to load, edit, or remove |
| `conditionalStyleBackgroundColor` | Background color override | Row/cell background color when the condition matches |
| `conditionalStyleForegroundColor` | Foreground color override | Text color when the condition matches |
| `conditionalStyleFontFamily` | Font family override | Font family when the condition matches |
| `conditionalStyleFontSize` | Font size override | Font size when the condition matches |
| `conditionalStyleFontStyle` | Font style override | Font style when the condition matches |
| `conditionalStyleFontWeight` | Font weight override | Font weight when the condition matches |
| `conditionalStyleTextAlign` | Text align override | Text alignment when the condition matches |
| `conditionalStyleLookupPick` | Select item | Item picker for lookup-field conditions |

**Diagnostics**
| Property | Label | Description |
|---|---|---|
| `enableDiagnostics` | Enable diagnostics logging | Enables console logging for debugging and troubleshooting |

---

## Marquee

Displays a scrolling announcement ticker (**SPS Marquee**) rendered as an animated bar, typically positioned above the Suite Navigation. Ideal for broadcasting company-wide announcements or news items. Supports Microsoft Teams tab deployment.

### Properties

**Content/Text**
| Property | Label | Description |
|---|---|---|
| `description` | Scrolling Text | Announcement text to scroll (max 250 characters, no HTML tags) |
| `listName` | List | Optional SharePoint list containing dynamic announcements |
| `messageField` | Message column | Column containing message text; list values are rendered as plain text |
| `messageDuration` | Seconds to show each message | Time before advancing to the next list message (default: 8) |

**Style/Appearance**
| Property | Label | Description |
|---|---|---|
| `marqueeBackColor` | Announcements Background Color | Background color of the marquee bar (supports alpha transparency) |
| `marqueeTextColor` | Announcements Text Color | Text color of the scrolling announcement (supports alpha transparency) |
| `fontFamily` / `fontSize` / `fontStyle` / `fontBold` | Typography | Announcement font settings |
| `marqueeHeight` | Marquee height | CSS height/minimum height (default: `32px`) |
| `scrollSpeed` | Scroll speed | Seconds per complete pass; lower is faster (default: 50) |
| `scrollDirection` | Scroll direction | Left or right |

**Behavior/Status**
| Property | Label | Description |
|---|---|---|
| `marqueeActive` | Active | Enable/disable display of the marquee announcement |
| `enableDiagnostics` | Enable diagnostics logging | Enables browser-console details for message loading and rendering |

---

## SharePointDForm (Dynamic Forms)

**Dynamic Forms** is a no-code visual form builder and runtime for SharePoint lists. It supports 15+ field types (text, number, date/time, choice, multi-select, lookup, person, taxonomy, image, URL, attachment, rich text, and more), conditional field visibility/required/read-only rules, multi-step wizard forms with progress indicators, advanced validation with custom expressions, default value assignment, filtering to resolve which record to load, permission-based access control, and New/Edit/View modes. Forms can redirect on submit/cancel, support custom button styling and label positioning, and consume selected item IDs and modes from [ListControl](#listcontrol), [GridControl](#gridcontrol), or [Calendar](#calendar).

### Properties

**List & Schema Configuration**
| Property | Label | Description |
|---|---|---|
| `listName` | SharePoint list | Target list for the form |
| `mode` | Form mode | New (create), Edit (update), or View (read-only) |

**Item/Record Resolution**
| Property | Label | Description |
|---|---|---|
| `useItemId` | Use fixed item ID | Use a static item ID instead of dynamic resolution |
| `itemId` | Item ID | Fixed record ID to open directly (optional) |
| `itemIdQueryParam` | Item ID query parameter | URL query parameter to read the item ID from (default: `itemid`) |
| `dynamicPreferredSourceInstanceId` | Dynamic Item ID (from another web part) | Connects to a ListControl, GridControl, or Calendar instance to receive a record ID dynamically |

**Dynamic Data Source & Field Targeting**
| Property | Label | Description |
|---|---|---|
| `useDynamicItemIdAsItemId` | Use dynamic ID as form item ID | ON: dynamic value is used as the record ID to load/edit; OFF: dynamic value only populates a field |
| `dynamicItemTargetField` | Parent lookup column | Lookup column to populate with the selected ListControl, GridControl, or Calendar item ID |
| `dynamicItemMode` | Dynamic Mode (new/edit/view) | Form mode sourced from another web part |
| `dynamicItemModeReference` | Dynamic mode reference | Reference string for the dynamic mode source |
| `dynamicItemIdReference` | Dynamic item ID reference | Reference string for the dynamic ID source |

**Record Filtering**
| Property | Label | Description |
|---|---|---|
| `useDynamicValueAsFilter` | Use dynamic value to find record (filter mode) | ON: searches for a record where a column equals the dynamic value; OFF: dynamic value only used as a field value |
| `filterJson` | Record filter JSON | JSON array of filter conditions to find the record to edit/view (supports lookup, choice, text, number, boolean fields and dynamic values) |

**Filter Designer** (UI builder for `filterJson`)
| Property | Label | Description |
|---|---|---|
| `filterDesignerField` | Filter field | Field selector for building a filter condition |
| `filterDesignerType` | Field type | Data type of the filter field (text, number, lookup, boolean, datetime) |
| `filterDesignerOperator` | Operator | Comparison operator: `eq`, `ne`, `gt`, `ge`, `lt`, `le`, `contains`, `startswith`, `endswith` |
| `filterDesignerLogical` | Join with previous | Logical operator to chain conditions: `AND`/`OR` |
| `filterDesignerValueSource` | Value source | Dynamic value, static value, or expression |
| `filterDesignerValue` | Static value / Expression | Static value or expression (e.g., `today`, `date(-3)`, `me.email`) |
| `filterDesignerLookupPick` | Selected lookup item | Helper for selecting a lookup item ID to use in the filter |
| `filterDesignerSelectedIndex` | Existing conditions | Selects an existing filter condition to manage |

**Permissions & Access Control**
| Property | Label | Description |
|---|---|---|
| `permissionBaseLookupField` | Column to base permissions on | Lookup field that determines whether the user can access the form |
| `permissionScope` | Permission scope | List (check list permissions) or Item (check specific item permissions) |
| `permissionScopeItemId` | Permission scope item ID | Specific item ID used for item-level permission checks |
| `permissionScopeLookupPick` | Selected permission item | Helper for selecting the permission reference item |
| `permissionDeniedMessage` | Permission denied message | Form-designer message used consistently for form-list, lookup-list, lookup-item, and blocked-submit denials |

**Recent Field Controls**

- Fields are enabled by default and can be visibly disabled per field without changing configured/default values.
- Field labels support Top and Left; Boolean labels additionally support Bottom and Right.
- Boolean fields show no automatic Yes/No text. Optional **Checkbox text** can be configured beside the checkbox.
- DateTime fields support Date Only, Date and Time, and Time Only formats.
- Date and Time / Time Only fields support UTC (default) or Local browser time interpretation. Existing forms needing browser-local behavior should explicitly select Local.
- Permission-denied UI defaults to “You do not have permission to access or modify this record.” Detailed permission-source diagnostics remain console-only.

**Buttons & Submit Behavior**
| Property | Label | Description |
|---|---|---|
| `addSubmitButtonLabel` | Add button text | Submit button text in New mode (default: "Add") |
| `editSubmitButtonLabel` | Edit button text | Submit button text in Edit mode (default: "Save") |
| `submitButtonLabel` | Submit button text | Generic submit button text (default: "Submit") |
| `showCancelButton` | Show cancel button | Show/hide the cancel button |
| `cancelButtonLabel` | Cancel button text | Text for the cancel button (default: "Cancel") |
| `cancelRedirectUrl` | Redirect URL after cancel | URL to navigate to on cancel (leave empty to reset the form) |
| `submitRedirectUrl` | Redirect URL after submit | URL to navigate to after successful submission (leave empty to show a success message) |
| `onSubmitMessage` | Submit success message | Message shown after successful submission (default: "Form submitted successfully!") |

**Button Styling**
| Property | Label | Description |
|---|---|---|
| `buttonTextColor` | Button text color | Button text color (default: `#000000`) |
| `buttonBackgroundColor` | Button background color | Button background color (default: `#f0f0f0`) |
| `buttonBorderColor` | Button border color | Button border color (default: `#f0f0f0`) |
| `buttonBorderWidth` | Button border width | Border width in pixels (0–10, default: 1) |
| `buttonFontFamily` | Button font family | Font family for button text (default: `inherit`) |
| `buttonFontSize` | Button font size | Font size in pixels (10–40, default: 14) |
| `buttonFontStyle` | Button font style | Normal or Italic (default: normal) |
| `buttonFontBold` | Button font bold | Bold button text (default: false) |
| `buttonCornerStyle` | Button corner style | Square or Rounded (default: square) |
| `buttonCornerRadius` | Button corner radius (px) | Radius for rounded corners (0–40, default: 4) |

**Form Display & Layout**
| Property | Label | Description |
|---|---|---|
| `labelPosition` | Label position | Field labels appear on Top or Left (default: top) |
| `showFieldDescription` | Show field description | Display field descriptions from the form schema (default: false) |

**Advanced Validation Rules**
| Property | Label | Description |
|---|---|---|
| `advancedValidationEnabled` | Enable advanced validation rules | Master toggle for the advanced validation engine (default: false) |
| `advancedValidationJson` | Validation rules JSON | JSON array of validation rules: `[{"expression":"...","message":"...","targetField":"InternalName(optional)"}]` |

**Advanced Validation Designer** (UI builder for `advancedValidationJson`)
| Property | Label | Description |
|---|---|---|
| `validationDesignerExpressionField` | Expression field selector | Field picker for building validation expressions |
| `validationDesignerExpression` | Expression | Validation logic expression using fields, operators, and functions |
| `validationDesignerMessage` | Validation message | Error message shown when validation fails |
| `validationDesignerTargetField` | Message target field (optional) | Field where the inline validation error displays |
| `validationDesignerExpressionTemplate` | Expression helper | Helper templates: `today()`, `now()`, `daysFromToday(n)`, `isEmpty()`, `hasValue()`, `between()`, operators (`==`, `!=`, `>`, `>=`, `<`, `<=`, `&&`, `\|\|`, `!`) |
| `validationDesignerOneClickTemplate` | One-click template | Quick templates for common validations: date range, required pair, conditional required |
| `validationDesignerSelectedIndex` | Existing rules | Selects an existing validation rule to manage |

**Default Values**
| Property | Label | Description |
|---|---|---|
| `defaultValuesJson` | Default values JSON (Add mode) | JSON object mapping field internal names to default values for New mode: `{"Category":"IT","IsActive":true,"DepartmentLookup":1}` |

**Default Values Designer** (UI builder for `defaultValuesJson`)
| Property | Label | Description |
|---|---|---|
| `defaultDesignerField` | Default field | Field selector for setting a default value |
| `defaultDesignerValueType` | Default value type | String, Number, Boolean, JSON, or Expression |
| `defaultDesignerValue` | Default value | The value to set as default |
| `defaultDesignerLookupPick` | Selected lookup item | Helper for selecting a lookup item ID to use as a default |
| `defaultDesignerSelectedField` | Existing defaults | Selects an existing default entry to manage |

**Diagnostics & Development**
| Property | Label | Description |
|---|---|---|
| `enableDynamicDiagnostics` | Enable dynamic diagnostics | Toggle diagnostic console logging for dynamic data sources and runtime behavior (default: true) |
| `formSchemaJson` | Form schema | JSON representation of the entire form structure (generated by the designer; not typically hand-edited) |

**Designer Mode**
| Property | Description |
|---|---|
| `isInDesignerMode` | Internal flag indicating whether the form is in visual designer editing mode |

> Note: Many JSON-backed properties (filter, validation, defaults, conditional styles) have companion "validate" buttons/result fields in the property pane for testing JSON before saving. Expressions support field references (`field("InternalName")`), date functions (`today()`, `date(n)`), and user context (`me`, `me.email`, `me.login`, `me.id`).

---

## Tab

Creates a tabbed interface control (**SPS Tabs**) for modern SharePoint pages and Teams, grouping either page Sections or individual Web Parts into tabs. Supports custom tab shape, fonts, colors, optional tab icons/images with positioning, and custom CSS overrides. Non-active tab content is hidden in view mode while remaining visible in edit mode. Works on modern SharePoint pages, Teams tabs, and Teams personal apps (not classic SharePoint).

> This folder also contains a separate `kmMarquee` web part (**SPS Marquee**) — a distinct scrolling-announcement component with its own text/background color properties and active toggle, not a duplicate. See [Marquee](#marquee) for the equivalent component maintained in its own project.

### Properties

**Tab Content & Management**
| Property | Label | Description |
|---|---|---|
| `collectionData` | Tabs | Collection defining each tab's title, text position (left/center/right), image URL, image position (left/right), and image-only display option |

**Tab Target & Mode**
| Property | Label | Description |
|---|---|---|
| `TabType` | Tab Area Type | Groups "Sections" or "WebParts" into tabs |

**Tab Shape & Layout**
| Property | Label | Description |
|---|---|---|
| `tabShape` | Tab Shape | "Rounded Top Corners" or "Squared" |
| `tabHeight` | Tab Height (px) | Tab height in pixels, 36–160 (default: 60) |
| `autoTabWidth` | Auto Size Width | Automatically calculate tab width |
| `tabWidth` | Tab Width (px) | Tab width in pixels, 60–400 (default: 120); disabled when Auto Size Width is on |

**Font & Text Styling**
| Property | Label | Description |
|---|---|---|
| `fontFamily` | Font | Font family (Segoe UI, Arial, Calibri, Verdana, Tahoma, Trebuchet MS, Georgia, Times New Roman, Courier New; default: Segoe UI) |
| `fontStyle` | Font Style | Normal, Italic, or Oblique (default: Normal) |
| `isBold` | Bold | Bold tab text |
| `fontSize` | Font Size (px) | Font size in pixels, 10–36 (default: 14) |

**Colors**
| Property | Label | Description |
|---|---|---|
| `selectedColor` | Active Tab Color | Active/selected tab highlight color (default: `#8A1717`) |
| `disableColor` | Inactive Tab Color | Inactive/unselected tab color (default: `#393939`) |

**Advanced Styling**
| Property | Label | Description |
|---|---|---|
| `overrideCSS` | Override CSS (URL of CSS) | URL to an external CSS file for custom tab styling |
| `useGlobalCSS` | Use Global CSS for Tabs | Applies a global SharePoint CSS override from `/tabs/css/tabsoverride.css` |
| `enableDiagnostics` | Enable diagnostics logging | Enables browser-console details for tab-zone discovery and configuration |

---

## Tiles

Displays a responsive grid of clickable navigation tile cards (**SPS Tiles**). Tiles support normal and hover images, image-only modes, per-tile colors, configurable shapes and dimensions, typography, and instant/fade/directional-slide hover transitions. Shows placeholder messaging when no tiles are configured and supports Teams tab deployment.

### Properties

| Property | Label | Description |
|---|---|---|
| `collectionData` | Tile data | Ordered tile collection with title, URL, description, target, color, normal image settings, and hover image settings |
| `title` | — | Web part instance title displayed at the top of the component |
| `tileShape` | Shape | `squared`, `rounded`, or `round` |
| `backgroundColor` / `textColor` / `hoverColor` | Colors | Global tile, text, and hover colors; each tile may override its normal background |
| `hoverSameAsBackground` | Same hover color | Reuses the normal background color while hovering |
| `hoverTransition` | Mouse-over transition | `solid`, `fade`, or slide from top/bottom/left/right |
| `fontFamily` / `fontStyle` / `fontBold` / `fontSize` | Typography | Global tile text settings |
| `tileWidth` / `tileHeight` | Tile dimensions | Width and height in pixels (defaults: 140 × 140) |
| `enableDiagnostics` | Enable diagnostics logging | Enables browser-console tile rendering details |

---

## Full Width Control

**SPS Full Width Control** is a layout-only web part that expands its containing SharePoint section to the available page width. It has no configurable properties. The control displays a selectable marker while the page is being edited and collapses to an invisible, zero-height host in read mode.

### Usage and configuration

1. Place the control in the same section as the content that needs additional width.
2. Publish or preview the page and verify the section with and without the property pane visible.
3. Add one control to each section that should expand; neighboring sections remain unchanged.

Common uses include wide calendars, grids, list tables, reports, tile navigation, and carousels. See [FullWidthControl/README.md](FullWidthControl/README.md) for section behavior and limitations.

---

## Print Control

**SPS Print Control** creates a printable snapshot of the SharePoint page content area without site navigation or chrome. It automatically discovers the page canvas or accepts a custom CSS selector for a specific printable region.

### Key properties

| Property | Purpose |
|---|---|
| `contentSelector` | Optional CSS selector for the printable region |
| `buttonLabel` / `alignment` | Print command text and placement |
| Button color, font, border, and corner properties | Visual presentation of the print command |
| `enableDiagnostics` | Logs content discovery and print preparation |

Use it with Report Forms for individual records or with composed Calendar, List Control, and Grid Control report pages. Always verify browser print preview. See [PrintControl/README.md](PrintControl/README.md).

---

## Report Forms

**Report Forms** renders a SharePoint list item through the Dynamic Forms schema in a permanently read-only mode. The item can be fixed, supplied through a URL parameter, resolved by a filter, or received dynamically from Calendar, List Control, or Grid Control.

### Key properties

| Property | Purpose |
|---|---|
| `listName` | SharePoint list containing report records |
| `formSchemaJson` | Visual designer output for fields and layout |
| Item ID, query parameter, dynamic reference, and filter settings | Resolve the record displayed by the report |
| Full/fixed width and diagnostic settings | Control layout and troubleshooting |

An embedded List Control can display related records. Report Forms does not submit changes or transition into edit mode. See [ReportForms/README.md](ReportForms/README.md).

---

## Script Editor

**SPS Script Editor** hosts trusted custom HTML, CSS, and JavaScript on a page. Authors may enter code directly or load supported text/script files. Optional settings expose classic SharePoint page context and remove container padding for embedded layouts.

### Security and usage

- The manifest requires custom script and the web part executes unsanitized code.
- Restrict page editing and web-part configuration to trusted authors.
- Review scripts as code, never place secrets in page content, and use only trusted external assets.
- Test read/edit mode, navigation, and repeated rendering before publishing.

Use this control for approved legacy integrations or small internal embeds when a dedicated SPFx component is not available. See [ScriptEditor/README.md](ScriptEditor/README.md).

---

## Common configuration

- **List access:** list-backed controls require visitors to have read access to every configured list, library, view, field, lookup source, and image location. Create/update/delete actions additionally require the corresponding SharePoint permissions.
- **Internal field names:** designers and JSON-backed rules refer to SharePoint internal names. Recheck mappings after changing a list, view, or schema.
- **Designer-managed JSON:** use the Calendar, Dynamic Forms, Grid Control, and List Control designers for filters, styles, schemas, defaults, and validation. Hand-edit JSON only when migrating a reviewed advanced configuration.
- **Dynamic data:** give List Control instances meaningful names, then connect their selected item ID/mode to Dynamic Forms or Report Forms. Calendar and Grid Control can participate in the same selection workflow.
- **Width:** use a control's `forceFullWidth` property when it should expand independently. Use Full Width Control when all web parts in a shared section should expand together.
- **CSS overrides:** host override stylesheets in a stable location readable by the audience. Treat them as versioned code and avoid selectors that affect unrelated page elements.
- **Diagnostics:** enable diagnostic logging temporarily while configuring or troubleshooting, then review the browser console for source discovery, request, rule, and dynamic-data details.

## Common scenarios

### Record selection and editing

Place List Control or Grid Control beside Dynamic Forms. The selector publishes the item ID and New/Edit/View mode; Dynamic Forms loads the corresponding record and workflow.

### Read-only report workspace

Use List Control, Grid Control, or Calendar as the selector, Report Forms as the detail surface, and Print Control for a clean browser/PDF output. Add Full Width Control when the complete section needs more room.

### Event operations

Combine multiple Events lists in Calendar, apply filters and conditional styles, and connect selected events to Dynamic Forms for updates or Report Forms for read-only details.

### Portal navigation

Use Tabs to organize page sections and Tiles or Link Button for navigation. Full Width Control can expand the section into a wide navigation band.

### Communications

Use BannerClock for centrally managed page chrome, Marquee for urgent or rotating announcements, Carousel for visual campaigns, and Accordion for longer supporting information.

## Deployment checklist

1. Build and package each project independently from its folder.
2. Upload the generated `.sppkg` from `sharepoint/solution/` to the SharePoint app catalog and deploy it according to the target scope.
3. Add the app to the target site where required and activate extension features such as BannerClock at the intended scope.
4. Confirm list/library permissions, custom-script policy for Script Editor, Site Assets configuration for BannerClock, and external CSS/image accessibility.
5. Add the component to a test page, configure it, test with representative user permissions, and inspect the browser console when diagnostics are enabled.
6. Publish the page and repeat key workflows in read mode, including property-pane/full-width behavior, dynamic connections, printing, and target-page navigation.

---

## Building

Each component is an independent SPFx solution with its own `package.json`, `gulpfile.js`, and `config/` folder. From within a component's folder:

```bash
npm install
npm i -g gulp
gulp bundle --ship
gulp package-solution --ship
```

See each component's own `README.md` for any component-specific build notes.

