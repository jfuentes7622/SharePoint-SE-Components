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

## Usage

1. Add **SPS Link Button** to a page and enter a short button title and valid destination URL.
2. Choose the link target and horizontal alignment.
3. Configure shape, dimensions, colors, typography, and an optional image icon.
4. Test keyboard focus, the destination, and the result at narrow page widths before publishing.

## Properties and common configuration

- **Content:** `description` is the displayed label and `Link` is the destination. The label rejects HTML and is limited to 50 characters.
- **Navigation:** `linkTarget` determines whether the destination opens in the current or a new browser context; `Align` positions the button.
- **Appearance:** shape, background/text colors, font family/style/weight/size, height, and width control the standard button.
- **Icon:** `showIcon`, image URL, size, and left/right position add a visual cue without replacing the accessible text label.
- **Advanced:** `overrideCssUrl` loads an external stylesheet and `enableDiagnostics` logs validation and rendering details.

Use a URL that visitors can access and avoid fixed dimensions that clip the longest translated label.

## Common scenarios

- A primary call to action linking to a form or service portal.
- A compact download or policy button with a recognizable icon.
- Consistent previous/next navigation across a set of guidance pages.
- A branded link whose colors and typography match a departmental site.
