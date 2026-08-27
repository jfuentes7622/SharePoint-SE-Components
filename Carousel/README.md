# SPS Carousel

SPFx 1.5.1 image carousel backed by a SharePoint document or picture library.

## Features

- Automatically play library images with pause-on-hover behavior.
- Show each image's library title, description, and link in a configurable caption.
- Use optional display-order, visibility, and expiration metadata when available.
- Configure carousel dimensions and animation timings.
- Render images as rectangles or circles.
- Enable diagnostic console logging.

## Configuration

| Property | Description | Default |
|---|---|---|
| `carouselSlideLibrary` | Document or picture library containing slide images. | Required |
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
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.

## Usage

1. Create or choose a SharePoint picture/document library and upload the slide images.
2. Add **SPS Carousel** to a page and select the library.
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
