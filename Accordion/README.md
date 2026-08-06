# SPS Accordion

SPFx 1.5.1 web part that renders SharePoint list items as collapsible accordion sections.

## Features

- Map list columns to accordion headers and HTML body content.
- Allow one open section at a time or multiple open sections.
- Customize header and content colors, fonts, and bold text.
- Load an optional override stylesheet.
- Enable diagnostic console logging.

## Configuration

| Property | Description |
|---|---|
| `listName` | SharePoint list containing accordion items. |
| `itemName` | Column used as each section header. |
| `itemContent` | Column used as each section body. |
| `optionChoice` | Single-selection or multiple-selection accordion behavior. |
| `headerBackgroundColor` | Header background color. |
| `headerTextColor` | Header text color. |
| `headerFontBold` | Bold header text. |
| `contentBackgroundColor` | Expanded content background color. |
| `contentTextColor` | Expanded content text color. |
| `contentFontBold` | Bold content text. |
| `fontFamily` | Font family for headers and content. |
| `fontStyle` | Normal, italic, or oblique text. |
| `overrideCssUrl` | Optional external stylesheet loaded after component styles. |
| `enableDiagnostics` | Writes lifecycle and data-loading details to the browser console. |

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.
