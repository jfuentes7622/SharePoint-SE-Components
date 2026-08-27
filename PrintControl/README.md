# SPS Print Control

SPFx 1.5.1 web part that prints the SharePoint page content area without the site header, navigation, or page chrome.

## Features

- Add the control only to pages that should expose printing.
- Print all web parts in the enclosing SharePoint canvas, including multiple ReportForms or DynamicForms instances.
- Exclude the Print Control, page editing controls, and form action buttons from the printed output.
- Preserve current input, textarea, checkbox, and select values in the printable clone.
- Align and style the print button with configurable colors, typography, border, and corners.
- Override automatic canvas discovery with an optional CSS selector.
- Enable diagnostic console logging.

## Configuration

| Property | Description | Default |
|---|---|---|
| `buttonLabel` | Visible print button label. | `Print` |
| `alignment` | Left, center, or right alignment. | Right |
| `buttonTextColor` / `buttonBackgroundColor` / `buttonBorderColor` | Button colors. | Black / light gray / light gray |
| `buttonFontFamily` / `buttonFontSize` / `buttonFontStyle` / `buttonFontBold` | Button typography. | Inherited / 14px / normal / disabled |
| `buttonBorderWidth` | Border width in pixels. | `1` |
| `buttonCornerStyle` / `buttonCornerRadius` | Square or rounded corners. | Square / 4px |
| `contentSelector` | Optional CSS selector for the printable region. Blank uses the enclosing SharePoint canvas. | Empty |
| `enableDiagnostics` | Writes validation and rendering details to the browser console. | Enabled |

Place one SPS Print Control web part anywhere in the same page canvas as the content to print. Additional Print Control instances are automatically removed from printed output, but normally only one is needed.

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.
