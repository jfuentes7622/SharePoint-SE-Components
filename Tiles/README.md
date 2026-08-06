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
