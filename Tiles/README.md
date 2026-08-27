# SPS Tiles

SPFx 1.5.1 web part that displays a responsive collection of navigation tiles.

## Features

- Configure title, description, URL, and link target for each tile.
- Set per-tile normal and hover images, image position, image-only mode, and background color override.
- Customize global shape, colors, typography, width, and height.
- Choose instant, fade, or directional slide hover transitions.
- Enable diagnostic console logging.

## Tile Collection

| Field | Description |
|---|---|
| `title` | Required tile label. |
| `description` | Optional supporting text. |
| `url` | Required destination URL. |
| `target` | Opens in the current context or a new tab (`_blank`). |
| `color` | Optional tile background override. |
| `imageUrl` / `imagePosition` / `imageOnly` | Normal image, its position, and image-only mode. |
| `hoverImageUrl` / `hoverImagePosition` / `hoverImageOnly` | Image settings used while hovering. |

## Appearance

| Property | Description | Default |
|---|---|---|
| `tileShape` | `squared`, `rounded`, or `round`. | `rounded` |
| `backgroundColor` / `textColor` | Default tile background and text colors. | `#8A1717` / `#FAFAFA` |
| `hoverColor` | Hover background color. | `#6a1010` |
| `hoverSameAsBackground` | Reuses the normal background while hovering. | Enabled |
| `hoverTransition` | `solid`, `fade`, `slide-top`, `slide-bottom`, `slide-left`, or `slide-right`. | `fade` |
| `fontFamily` / `fontStyle` / `fontBold` | Tile typography. | Inherited / normal / disabled |
| `fontSize` | Text size in pixels. | `14` |
| `tileWidth` / `tileHeight` | Tile dimensions in pixels. | `140` / `140` |
| `enableDiagnostics` | Writes tile rendering details to the browser console. | Disabled |

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.

## Usage

1. Add **SPS Tiles** and use the tile collection editor to create and order destinations.
2. Give each tile a title, URL, target, and optional normal/hover image and color overrides.
3. Choose stacked, fixed-column, or responsive flow layout, then set tile dimensions and gaps.
4. Configure shape, typography, colors, hover transition, and web-part appearance; test keyboard and narrow-width behavior before publishing.

## Properties and common configuration

- **Content:** `collectionData` stores per-tile title, destination, target, colors, image URLs, image positions, opacity, and image-only/background modes.
- **Navigation:** each tile opens its URL in the configured current or new browser context.
- **Layout:** tile width/height, stacked/columns/flow mode, column count, and gap determine wrapping and density.
- **Appearance:** global and per-tile colors, shape, typography, normal/hover imagery, hover transition, title, surface, border, and padding control presentation.
- **Advanced:** `forceFullWidth` gives large tile sets more room and `enableDiagnostics` logs collection and rendering details.

Use flow layout when responsive wrapping is more important than a fixed column count. Keep titles short and ensure image URLs are readable by the page audience.

## Common scenarios

- A quick-links portal for common applications and team sites.
- A department directory with branded images and per-tile colors.
- A product, service, or resource catalog with hover imagery.
- A full-width navigation band combined with Tabs or Full Width Control.
