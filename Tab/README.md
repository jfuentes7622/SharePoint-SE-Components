# SPS Tabs

SPFx 1.5.1 web part for grouping modern SharePoint page sections or individual web parts into a tabbed interface.

## Features

- Build and reorder tabs in a property-pane collection editor.
- Target page sections or individual web parts.
- Configure each tab's title, text alignment, image, image position, and image-only mode.
- Customize tab shape, dimensions, typography, active color, and inactive color.
- Automatically size tab widths or use a fixed width.
- Load an optional custom stylesheet.
- Keep all target content visible while the page is in edit mode.
- Enable diagnostic console logging for zone discovery and configuration troubleshooting.

## Configuration

### Tabs

The `collectionData` editor defines tabs in display order. Use its Move up and Move down actions to change the order.

| Field | Description |
|---|---|
| `Title` | Required tab label. |
| `textPosition` | Label alignment: `left`, `center`, or `right`. |
| `imageUrl` | Optional image displayed in the tab. |
| `imagePosition` | Places the image to the `left` or `right` of the label. |
| `onlyImage` | Displays only the image and omits the text label. |

### Target And Appearance

| Property | Description | Default |
|---|---|---|
| `TabType` | `1` groups page sections; `2` groups individual web parts. | Sections (`1`) |
| `tabShape` | Rounded top corners or square corners. | `rounded` |
| `fontFamily` | Font family used by tab labels. | `Segoe UI` |
| `fontStyle` | Normal, italic, or oblique text. | `normal` |
| `isBold` | Uses bold tab labels. | Disabled |
| `fontSize` | Label size from 10 to 36 pixels. | `14` |
| `tabHeight` | Tab height from 36 to 160 pixels. | `60` |
| `autoTabWidth` | Sizes each tab from its content. Image-only tabs are always auto-sized. | Disabled |
| `tabWidth` | Fixed width from 60 to 400 pixels when auto sizing is disabled. | `120` |
| `selectedColor` | Active tab color, including optional alpha transparency. | `#8A1717` |
| `disableColor` | Inactive tab color, including optional alpha transparency. | `#393939` |
| `enableDiagnostics` | Writes tab discovery and configuration details to the browser console. | Enabled |

### Custom CSS

- Set `overrideCSS` to the URL of a custom stylesheet, or
- Enable `useGlobalCSS` to load `/tabs/css/tabsoverride.css` from the web application root.

The global CSS option takes precedence over `overrideCSS`.

## Requirements

- SharePoint Server Subscription Edition
- SharePoint Framework 1.5.1
- A modern SharePoint page containing the sections or web parts to group
- Node.js 10.24.x and npm 6.x for local builds

## Build And Package

From the `Tab` folder:

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.

## Disclaimer

This code is provided as-is without warranty of any kind, either express or implied.
