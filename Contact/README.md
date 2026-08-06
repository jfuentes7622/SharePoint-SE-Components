# SPS Contacts

SPFx 1.5.1 personnel directory web part for SharePoint Server Subscription Edition.

## Features

- Load contacts from a SharePoint list or maintain inline contact entries in the property pane.
- Filter contacts through directorate and division hierarchy.
- Display job title, name, email, phone, VOIP, biography link, and optional photo.
- Render photos as square, rounded, or circular images.
- Customize card headers, titles, information areas, fonts, colors, alignment, and corners.
- Load an optional override stylesheet.
- Enable diagnostic console logging for list, record, and image services.

## Data Source

Enable **Use Existing List** to select a SharePoint list and map the job title, directorate, and division columns. The remaining contact fields follow the component's contact schema, including name, email, phone number, biography link, image link, display-photo flag, VOIP, group, and display order.

Disable **Use Existing List** to enter contact records directly in the property-pane collection editor.

## Appearance

| Property | Description |
|---|---|
| `imageShape` | Contact image shape: `square`, `rounded`, or `circle`. |
| `imageWidth` | Contact image size in pixels. |
| `PersonnelPanelHeaderBackColor` / `PersonnelPanelHeaderTextColor` | Card header colors. |
| `headerFontFamily` / `headerFontStyle` / `headerFontBold` | Header typography. |
| `headerAlignment` / `headerTopCorners` | Header alignment and corner style. |
| `tileTitleColor` | Job-title text color. |
| `titleFontFamily` / `titleFontStyle` / `titleFontBold` | Job-title typography. |
| `tileInfoBackgroundColor` | Background behind contact details. |
| `overridecss` | Optional external stylesheet URL. |
| `enableDiagnostics` | Writes contact loading and rendering details to the browser console. |

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.
