# SPS Link Button

SPFx 1.5.1 web part that renders a configurable hyperlink as a styled button.

## Features

- Validate button text and destination URLs.
- Align the button left, center, or right.
- Configure shape, dimensions, colors, and typography.
- Add an optional image icon before or after the label.
- Load an optional override stylesheet.
- Enable diagnostic console logging.

## Configuration

| Property | Description | Default |
|---|---|---|
| `description` | Button label, limited to 50 characters with HTML rejected. | `Go to..` |
| `Link` | Destination URL. | `https://` |
| `Align` | Left, center, or right alignment. | Right |
| `buttonShape` | Square or rounded corners. | `rounded` |
| `buttonBackgroundColor` | Button background color. | `#8A1717` |
| `buttonFontColor` | Button text color. | `#ffffff` |
| `fontFamily` / `fontStyle` / `fontBold` | Button typography. | Inherited / normal / disabled |
| `fontSize` | CSS font size. | `1.25rem` |
| `buttonHeight` / `buttonWidth` | CSS dimensions. | `5vh` / `200px` |
| `showIcon` | Displays an optional image icon. | Disabled |
| `iconImageUrl` | Image URL used when the icon is enabled. | Empty |
| `iconSize` | CSS icon size. | `24px` |
| `iconPosition` | Places the icon on the left or right. | `left` |
| `overrideCssUrl` | Optional external stylesheet URL. | Empty |
| `enableDiagnostics` | Writes validation and rendering details to the browser console. | Enabled |

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.
