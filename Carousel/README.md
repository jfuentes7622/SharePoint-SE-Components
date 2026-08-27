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
