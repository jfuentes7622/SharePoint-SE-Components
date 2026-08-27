# SPS Calendar

SPFx 1.5.1 SharePoint events calendar powered by the free, MIT-licensed FullCalendar library.

## Features

- Combine events from multiple visible SharePoint Events lists.
- Expand SharePoint recurring-event series into occurrences for the visible range.
- Display month, week, day, or agenda-style list views.
- Show titles, start/end times, all-day status, and locations.
- Show or hide weekends.
- Filter events using SharePoint item fields and JSON conditions.
- Apply global event styling and priority-based conditional style rules.
- Group events into field-based horizontal swim lanes without premium FullCalendar plugins.
- Open an event's default SharePoint display form or a custom page URL.
- Open an animated in-place event details panel with configurable fields and ordering.
- Publish the selected event ID and `view` mode for Dynamic Form connections.
- Enable diagnostic console logging for list and event loading.

## Configuration

| Property | Description | Default |
|---|---|---|
| `dataSources` | Ordered collection of Events/calendar-view lists, date-column mappings, and target pages. | Empty |
| `listName` | Legacy single-list fallback used by existing configurations. | Empty |
| `defaultView` | `dayGridMonth`, `timeGridWeek`, `timeGridDay`, or `listWeek`. | `dayGridMonth` |
| `showWeekends` | Includes Saturday and Sunday. | Enabled |
| `calendarHeight` | Calendar height in pixels. | `600` |
| `enableSwimlanes` | Replaces the standard calendar grid with field-grouped horizontal lanes. | Disabled |
| `swimlaneFieldName` | Internal name of the SharePoint field used to create lanes. | Empty |
| `swimlaneDays` | Rolling visible range: `7`, `14`, or `30` days. | `7` |
| `swimlaneUnassignedLabel` | Lane label for events whose grouping field is empty. | `Unassigned` |
| `filterJson` | Conditions applied to SharePoint items before events are rendered. | Empty |
| `showLinkToItem` | Opens the selected event directly or adds an item link to the details panel. | Disabled |
| `enableEventDetails` | Opens an expanding details panel when an event is selected. | Disabled |
| `eventDetailsFieldsJson` | Ordered collection of SharePoint fields and display labels shown in the details panel. | Standard event fields |
| `detailsLinkPresentation` | Displays the details-panel item link as a button or linked title. | Button |
| `linkTargetPageUrl` | Legacy custom target page for single-list configurations. | Empty |
| `linkTargetIdParam` | Query parameter used for the selected item ID on custom pages. | `itemid` |
| `includeReturnUrlParam` | Adds the current URL in a `return` parameter. | Disabled |
| `conditionalStyleJson` | Data-driven event appearance rules. | Empty |
| `enableDiagnostics` | Writes calendar lifecycle and data-loading details to the browser console. | Enabled |

The dedicated **Appearance** property page separates styling for the calendar surface, web-part header, calendar date title, toolbar buttons, events, selected event, and expanded event details panel. The optional description appears directly below the editable title. It has independent text color, font family, size, style, and weight, while title alignment, header background color, and minimum height apply to the entire title/description/legend header. Set the minimum height to `0` to let the header grow automatically with its content. The data-source legend icon is anchored in the header and appears only when more than one source is configured. Selected-event controls include background, text, and border colors; font family, size, style, and weight; border width; and corner radius. The selected style overrides normal and conditional event styling in both the standard calendar and swim lanes, remains visible after the details panel closes, and moves when another event is selected. The details panel controls include overlay color and opacity; panel background, border, radius, width, padding, and shadow; independent title, field-label, and field-value typography; row dividers; and close-button colors and radius.

Use **Manage data sources** to add and order Events lists or generic lists that contain a SharePoint calendar view. Each row selects its source list, start and end DateTime columns, and target page. **Defaults** resolves to `EventDate` and `EndDate` for classic Events lists and to the start/end fields configured in a modern calendar view. Explicit column selections override those defaults. The target dropdown includes **Display List Default Form** plus `.aspx` pages discovered from the current site's Site Pages and publishing Pages libraries, so no URL needs to be typed. Events from all rows are merged into the same calendar. Filters and conditional style entries can apply to all data sources or one selected source; selecting a source loads that list's fields into the designer. Existing filter and style entries without `sourceListName` remain global. SharePoint recurrence expansion is available for Events lists using the built-in date columns; generic calendar-view lists render non-recurring items.

Each data source can define its own event background with the collection's color picker. Event text color remains global. Source backgrounds are the base event appearance; conditional style rules can override them, and selected-event colors remain the highest-priority state. When two or more sources are configured, an information icon in the web-part header opens a legend on hover or click. Each legend swatch uses the source background and global event text color to identify the source list.

Event details use SharePoint field metadata when displaying values. Rich-text multiline columns render their sanitized HTML formatting, while plain-text columns remain literal text; choosing the collection's **Text** format also forces literal text.

Recurring events automatically include a **Repeats** information row in the details panel. It describes the same daily, weekly, monthly, or yearly rule used to expand the event, including its interval, selected weekdays, occurrence count, or end date when available. One-time events do not show this row.

The current user needs read access to the selected Events list. Calendar reads the item fields so filters and style rules can use standard or custom field internal names.

Calendar queries events that overlap the currently visible month, week, day, list, or swim-lane range. Navigating to another range reloads that range from SharePoint, avoiding the previous oldest-500-item limit on large Events lists. Blank event end dates are supported, and the item Title is displayed on its calendar day.

## Event Details Panel

Enable **Show an expanding event details panel** under Event links, then select **Manage details fields**. Each collection row first selects a data source, then selects from that source's SharePoint fields. The selected event only displays detail rows assigned to its source. Rows also support a display label, label visibility, value format, and hide-when-empty behavior. Add, remove, edit, or drag rows to control their display order. Existing rows without a source remain assigned to the first data source for compatibility. When no custom collection is configured, Calendar shows Title, start time, end time, Location, Description, and Category.

Selecting an event publishes its dynamic-data value and opens the panel over the calendar. Select the X, the shaded backdrop, or press Escape to return to the calendar. Enable **Open the selected event** to add either an **Open event** button or linked event title to the panel. The target page, item-ID parameter, and return-page options are shared with direct event navigation. When the details panel is disabled, selecting an event navigates immediately instead.

## Recurring Events

SharePoint recurrence masters are expanded for the currently visible range. Calendar supports daily, weekday-only, weekly, monthly by date, monthly by weekday position, yearly by date, and yearly by weekday position patterns. Series with a fixed occurrence count or window end are bounded accordingly.

Each occurrence keeps the master event's duration, title, filtering, conditional styling, swim-lane grouping, and item-link behavior. Modified and deleted recurrence exceptions suppress the corresponding generated occurrence; modified exceptions are rendered from their own SharePoint item data.

## Swim Lanes

Enable **Use swim lanes**, then select a visible list field such as Category, Location, Assigned To, Department, or Resource. Calendar creates one horizontal lane per distinct field value. Lookup and person fields use their SharePoint display text, multi-value events can appear in more than one lane, and empty values appear in the configured Unassigned lane.

The lane toolbar moves through non-overlapping rolling date ranges and respects **Show weekends**. Multi-day events appear in each visible day they overlap. Selecting a lane event uses the same dynamic-data publication, conditional styling, and optional item-link behavior as the standard calendar view.

FullCalendar's native resource timeline is premium. This implementation provides field-based swim lanes using the existing MIT-licensed dependencies and does not require a scheduler license.

## Event Filters

Calendar has a dedicated **Event filter designer** property page. Select a SharePoint field, operator, logical join, and static or expression value, then add the condition to the generated JSON. Existing conditions can be loaded, updated, removed, or reset. Lookup fields provide a target-list item picker. Raw JSON editing and validation remain available on the same page.

Filters use `eq`, `ne`, `contains`, `notcontains`, `startswith`, `endswith`, `gt`, `ge`, `lt`, or `le`. Conditions are evaluated in order using `and` or `or`:

```json
[
	{ "field": "Category", "operator": "eq", "logical": "and", "value": "Meeting" },
	{ "field": "EventDate", "operator": "ge", "logical": "and", "valueType": "expression", "value": "today" }
]
```

Expression values support `today`, `now`, `date(n)`, `me`, `me.email`, `me.login`, `me.id`, and `me.title`. For example, `date(7)` means seven days from today.

## Conditional Styles

Calendar has a dedicated **Conditional event style designer** property page. Select a SharePoint field, operator, logical join, static or expression value, priority, colors, and typography, then add the condition to the generated JSON. Existing conditions can be loaded, updated, removed, or reordered. Reuse the same rule name across entries to build a multi-condition rule.

The raw JSON editor and validation result remain available on the same page for advanced changes. Rules are applied by ascending priority, so higher-priority matching rules override earlier values:

```json
[
	{
		"enabled": true,
		"priority": 10,
		"conditions": [
			{ "field": "Category", "operator": "eq", "logical": "and", "value": "Holiday" }
		],
		"style": {
			"backgroundColor": "#a4262c",
			"borderColor": "#751d22",
			"textColor": "#ffffff",
			"fontWeight": "bold"
		}
	}
]
```

Supported style values are `backgroundColor`, `borderColor`, `textColor` (or `color`), `fontFamily`, `fontSize`, `fontStyle`, and `fontWeight`.

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.

## Usage

1. Add **SPS Calendar** to a page and open **Manage data sources**.
2. Add one or more Events lists or lists that contain calendar views, then verify the start/end field mappings for each source.
3. Choose the default view, weekend visibility, height, and optional swim-lane grouping.
4. Use the filter, conditional-style, and event-details designers as needed; avoid hand-editing their stored JSON unless maintaining an existing advanced configuration.
5. Optionally connect the calendar's selected item and `view` mode to Dynamic Forms or Report Forms on the same page.

## Properties and common configuration

- **Data:** `dataSources` stores ordered list sources, field mappings, colors, and target pages. `listName` remains a legacy single-list fallback.
- **Behavior:** `defaultView`, `showWeekends`, `calendarHeight`, swim-lane settings, event details, target-page settings, and return-URL handling control interaction.
- **Rules:** `filterJson` and `conditionalStyleJson` are maintained by visual designers and support field comparisons plus current-user/date expressions.
- **Appearance:** calendar surface, title, toolbar, event, selected-event, details-panel, typography, border, and spacing groups can be configured independently.
- **Advanced:** `forceFullWidth` expands the web part to available page width. `enableDiagnostics` records source discovery, recurrence, filter, and rendering details.

Recurring-event expansion is intended for SharePoint Events lists. Users must have read access to every configured source and to fields used by filters, details, or style rules.

## Common scenarios

- Combine departmental Events lists into one color-coded organizational calendar.
- Use swim lanes to compare rooms, teams, resources, or event categories over 7, 14, or 30 days.
- Highlight urgent, overdue, or executive events with priority-based conditional styles.
- Show an in-page read-only report or form when a visitor selects an event.
- Link events to a custom details page while preserving a return URL to the calendar.
