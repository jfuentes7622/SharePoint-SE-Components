# SPS Tabs

SPFx 1.5.1 web part for grouping modern SharePoint page sections or individual web parts into a tabbed interface.

## Features

- Build and reorder tabs in a property-pane collection editor.
- Target page sections or individual web parts.
- Configure each tab's title, text alignment, image, image position, and image-only mode.
- Customize tab shape, dimensions, typography, active/inactive colors, borders, and the tab-strip line.
- Fade images on inactive tabs without changing label or border opacity.
- Style the outer web-part border and corners independently from individual tabs.
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
| `inactiveImageFade` | Percentage opacity reduction applied only to inactive tab images; active images remain fully opaque. | `45` |
| `tabBorderColor` | Border color for individual tabs. | `#000000` |
| `tabBorderWidth` | Individual tab border width from 0 to 20 pixels. | `0` |
| `tabBorderStyle` | Individual tab border type: none, solid, dashed, dotted, or double. | `solid` |
| `tabCornerRadius` | Rounded-tab corner radius from 0 to 40 pixels. | `10` |
| `tabLineColor` | Independent color of the line below the tab strip. | `#8A1717` |
| `tabLineWidth` | Tab-strip line width from 0 to 20 pixels. | `3` |
| `webPartBorderColor` | Border color for the outer web-part wrapper. | `#cccccc` |
| `webPartBorderWidth` | Outer wrapper border width from 0 to 20 pixels. | `0` |
| `webPartBorderStyle` | Outer wrapper border type: none, solid, dashed, dotted, or double. | `solid` |
| `webPartCornerStyle` | Uses square or rounded outer wrapper corners. | `square` |
| `webPartCornerRadius` | Outer wrapper radius from 0 to 40 pixels when rounded corners are selected. | `8` |
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

## Usage

1. Add **SPS Tabs** before the content it will organize and choose Sections or Web Parts mode.
2. Add tab definitions in the same order as their target content, including optional images and text/image positioning.
3. Configure width strategy, height, shape, colors, typography, and inactive-image fade.
4. Verify every tab in edit and read modes; section targeting depends on the surrounding SharePoint canvas order.

## Properties and common configuration

- **Tabs:** `collectionData` stores ordered labels, alignment, image URLs/positions, and image-only choices.
- **Target:** `TabType` determines whether tabs control neighboring SharePoint sections or represent web-part content.
- **Layout:** auto width or fixed width, tab height, shape, border/line, and content spacing control the frame.
- **Appearance:** active/inactive colors, font settings, title, web-part border, and image fade control visual hierarchy.
- **Advanced:** external or global CSS options support shared branding. `forceFullWidth` expands the control, and `enableDiagnostics` logs target discovery and selection.

In Sections mode, tab order must match section order. Recheck the configuration after authors insert, remove, or rearrange page sections.

## Common scenarios

- Group dashboard sections into Overview, Metrics, Risks, and Actions tabs.
- Organize multiple list/report areas without forcing visitors down a long page.
- Build icon-assisted tabs for a project or department hub.
- Combine Tabs, Tiles, and Full Width Control for a wide portal navigation experience.
