# SPS Calendar

SPFx 1.5.1 SharePoint events calendar powered by the free, MIT-licensed FullCalendar library.

## Features

- Load events from visible SharePoint Events lists.
- Display month, week, day, or agenda-style list views.
- Show titles, start/end times, all-day status, and locations.
- Show or hide weekends.
- Set a fixed calendar height or allow automatic sizing.
- Open an event's SharePoint display form when the event is selected.
- Enable diagnostic console logging for list and event loading.

## Configuration

| Property | Description | Default |
|---|---|---|
| `listName` | SharePoint Events list (`BaseTemplate` 106). | Empty |
| `defaultView` | `dayGridMonth`, `timeGridWeek`, `timeGridDay`, or `listWeek`. | `dayGridMonth` |
| `showWeekends` | Includes Saturday and Sunday. | Enabled |
| `calendarHeight` | Calendar height in pixels; `0` uses automatic height. | `0` |
| `enableDiagnostics` | Writes calendar lifecycle and data-loading details to the browser console. | Enabled |

The current user needs read access to the selected Events list. The web part reads the standard `Title`, `EventDate`, `EndDate`, `fAllDayEvent`, and `Location` fields.

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.
