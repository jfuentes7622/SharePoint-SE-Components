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

## Usage

1. Add **SPS Print Control** inside the page content that users need to print.
2. Leave the content selector blank to use automatic SharePoint canvas discovery, or provide a CSS selector for a specific report container.
3. Configure the button label, alignment, colors, typography, border, and corner style.
4. Test print preview in each supported browser and verify that site navigation and page chrome are excluded.

## Properties and common configuration

- **Target:** `contentSelector` optionally identifies the element to print. Use a stable ID or class owned by the page solution; invalid selectors fall back or fail to find content.
- **Content and placement:** `buttonLabel` supplies the command text and `alignment` positions the button within its host.
- **Appearance:** text/background/border colors, font family/size/style/weight, border width, corner style, and radius define the print button.
- **Advanced:** `enableDiagnostics` logs content discovery and print preparation details.

The control prints browser-rendered content. Page authors should verify tables, forms, overflow areas, and content loaded after initial page render in print preview.

## Common scenarios

- Print a Report Forms record without SharePoint navigation.
- Produce a clean hard copy of a Calendar, List Control, or Grid Control view.
- Target a composed report section while excluding action controls elsewhere on the page.
- Add a consistently styled print command to departmental report pages.
