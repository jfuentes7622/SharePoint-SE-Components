# SPS Carousel

SPFx 1.5.1 image carousel backed by a SharePoint document or picture library.

## October 26, 2026 release

- npm version: `0.0.8`
- SharePoint solution and feature version: `1.0.0.41`
- Package: `sharepoint/solution/sps-carousel.sppkg`
- Adds **Manual styling**, **Site theme only**, and **Site theme with overrides**.
- Existing instances remain in Manual styling unless an author selects a theme mode.
- Site theme modes cover the carousel surface, border, title, caption, links, and font families while retaining saved manual values.
- Disables optional legacy PnP usage telemetry without changing SharePoint's CSP.
- Handles early SharePoint theme notifications without rendering before web part initialization; later theme changes still refresh the Carousel.
- Empty or missing values in the selected optional link column no longer stop all slides from loading. Unlinked images still display; populated text and hyperlink fields retain their navigation URLs and link labels.

## Features

- Automatically play library images with pause-on-hover behavior.
- Render direct SharePoint image URLs without blocking on client-side canvas conversion.
- Apply the selected SharePoint library view's filters, sorting, scope, and row limit.
- Show each image's library title, description, and link in a configurable caption.
- Use optional display-order, visibility, and expiration metadata when available.
- Configure carousel dimensions and animation timings.
- Render images as rectangles or circles.
- Scale rectangular images to fit the available image area without stretching or cropping; differing aspect ratios leave empty space. Fixed-height slides reserve space for their title and footer, while height `0` follows the image's aspect ratio.
- Enable diagnostic console logging.

## Content Security Policy

The bundled PnP property-controls 1.1.0 color picker normally initializes legacy
Application Insights at import time, requesting
`https://az416426.vo.msecnd.net/scripts/a/ai.0.js`. SharePoint's CSP can block that
optional usage-telemetry script.

The Carousel webpack configuration replaces only PnP's telemetry module with a
local no-op before bundling. The color picker and Carousel functionality remain
available, and `enableDiagnostics` still controls local console diagnostics.
No CSP allowlist changes or external telemetry scripts are required.
Rebuild and redeploy the solution package to update existing installations.

## Configuration

Appearance mode supports manual styling, strict SharePoint site-theme colors/fonts, or site-theme colors/fonts with non-empty manual overrides. Strict site-theme mode disables the controlled color and font-family fields without clearing saved values.

| Property | Description | Default |
|---|---|---|
| `carouselSlideLibrary` | Document or picture library containing slide images. | Required |
| `carouselViewId` | SharePoint library view used to filter and order slides; choose all items to bypass a view. | Library default view |
| `slideTitleField` | Selected library column used for each slide title. | `Title` when available |
| `slideDescriptionField` | Selected library column used for each slide description. | `Description` when available |
| `slideLinkField` | Selected library column used for each slide link. | `LinkTarget` or `ClickLink` when available |
| `carouselWidth` | Carousel width in pixels. | `500` |
| `carouselHeight` | Height in pixels; `0` derives height from the image aspect ratio. | `0` |
| `carouselSlideInterval` | Time each slide remains visible, in milliseconds. | `5000` |
| `carouselTransitionInterval` | Transition duration, in milliseconds. | `1500` |
| `imageIsCircle` | Crops slide images into circles. | Disabled |
| `enableDiagnostics` | Writes lifecycle and image-loading details to the browser console. | Enabled |

## Library Columns

| Column | Type | Purpose |
|---|---|---|
| Any selected text-compatible column | Text | Slide title or description. |
| Any selected hyperlink or text column | Hyperlink or text | Slide destination. A hyperlink column's description becomes the link text. |
| `SlideOrder` | Number | Optional display order. |
| `Display` | Yes/No or choice | Optional visibility control. |
| `StartDate` / `Expiration` | Date and time | Optional visibility window. |

## Build And Package

```powershell
npm install
gulp clean --ship
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.

Run `npm run test:theme-lifecycle` to verify that early theme notifications do not
render against uninitialized properties or DOM, while later notifications still
refresh the web part.

Run `npm run test:image-sizing` to verify fixed-height image areas, automatic
height, caption space, image navigation, and unchanged circular-image rendering.

Run `npm run test:record-service` to verify empty optional link values, mixed
linked/unlinked slides, custom field mappings, and propagation of library errors.

## Usage

1. Create or choose a SharePoint picture/document library and upload the slide images.
2. Add **SPS Carousel** to a page and select the library and SharePoint view.
3. Map optional title, description, and link columns when those captions or actions are required.
4. Set dimensions and timing, preview the result at common page widths, and publish.

## Properties and common configuration

- **Data:** `carouselSlideLibrary` selects the image library; title, description, and link field mappings enrich each slide but are optional.
- **Behavior:** `carouselSlideInterval` controls display time and `carouselTransitionInterval` controls transition duration, both in milliseconds.
- **Appearance:** width/height, circular-image mode, surface and border styling, caption styling, and title/description/link typography define presentation.
- **Advanced:** `forceFullWidth` uses the available page width. An override stylesheet supports site-specific refinements, and `enableDiagnostics` logs library and image-loading details.

Use consistent image aspect ratios for a stable layout. Confirm that visitors can read the library and any metadata columns used by the carousel.

## Common scenarios

- A rotating campaign or announcement banner linked to landing pages.
- A product, facility, or project showcase sourced from a maintained image library.
- Leadership messages with a title and short caption on each slide.
- A full-width visual header whose content can be updated without editing the page.
