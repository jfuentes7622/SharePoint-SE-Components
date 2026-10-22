# SPS Tiles

SPFx 1.5.1 web part that displays a responsive collection of navigation tiles.

## October 22, 2026 release

- npm version: `2.0.5`
- SharePoint solution and feature version: `2.1.0.20`
- Package: `sharepoint/solution/sps-tiles.sppkg`
- Adds **Manual styling**, **Site theme only**, and **Site theme with overrides**.
- Existing instances remain in Manual styling unless an author selects a theme mode.
- Site theme modes cover web-part, title, and tile surfaces, text, borders, and font families. Per-tile color and image overrides remain intact.
- Fixes inline web-part title editing and adds a **Title** text field to the **Web part title** settings group.
- Adds independent color, thickness, style, and rounded-corner settings for the borders on each tile.

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
| `tileBorderColor` | Color for each tile's border; an unset color follows that tile's background. Configured colors follow the selected appearance mode. | Unset |
| `tileBorderWidth` | Tile border thickness, from 0 to 20 pixels. | `2` |
| `tileBorderStyle` | `none`, `solid`, `dashed`, `dotted`, or `double`. | `solid` |
| `tileCornerRadius` | Corner radius from 0 to 200 pixels when the shape is `rounded`; square and round shapes retain their shape. | `10` |
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

### Editing the web-part title

In page edit mode, enable **Show title** under **Web part title**. Change the
**Title** text field in that group, or type directly into the title above the
tiles. Inline edits update immediately and synchronize with the property pane.
Save and publish the page to retain the title for readers. Turning **Show title**
off hides the title without erasing its text; individual tile labels are unchanged.

From the repository root, run `node tests\tiles.test.js` to verify inline
editing, property-pane configuration, clearing, visibility toggles, and restored
read-mode properties. Live SharePoint save/publish verification is still required.

### Borders around individual tiles

Use the **Tile borders** property-pane group to configure the border color,
thickness, and style on every tile. These settings are separate from the outer
border under **Web part appearance**. Set thickness to **0** or style to **None**
to remove tile borders. To change the corner radius, choose **Rounded corners**
under **Appearance > Shape**, then adjust **Tile corner radius**.

Existing tiles retain their 2-pixel solid, background-colored borders until these
settings are changed. Image-only, background-image, and hover-image tiles use the
same settings without changing their configured outer dimensions. In **Site theme
only** mode, configured border colors follow the site theme and the color picker
is disabled, consistent with the other color controls.

The shared test command above also covers border property wiring, legacy defaults,
image variants, zero-width borders, styles, shape behavior, theme modes, and
restored read-mode properties.

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
