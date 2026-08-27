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

Enable **Use a List** to use SharePoint data, then choose **Use an Existing List** or **Create a New List**. After a list is selected or created, the **Columns to Use** group loads its columns for every contact field mapping.

To create a compatible list on demand, enter a name under **New Contact List Name** and select **Create Contact List**. The web part creates the required columns in the current SharePoint web, selects the new list, and configures all field mappings. The editing user must have permission to manage lists.

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
