# SPS Carousel

SPFx 1.5.1 image carousel backed by a SharePoint document or picture library.

## Features

- Automatically play library images with pause-on-hover behavior.
- Use optional title, link, display-order, and expiration metadata when available.
- Configure carousel dimensions and animation timings.
- Render images as rectangles or circles.
- Enable diagnostic console logging.

## Configuration

| Property | Description | Default |
|---|---|---|
| `carouselSlideLibrary` | Document or picture library containing slide images. | Required |
| `carouselWidth` | Carousel width in pixels. | `500` |
| `carouselHeight` | Height in pixels; `0` derives height from the image aspect ratio. | `0` |
| `carouselBackgroundColor` | Background color, including optional alpha transparency. | Theme/default |
| `carouselSlideInterval` | Time each slide remains visible, in milliseconds. | `5000` |
| `carouselTransitionInterval` | Transition duration, in milliseconds. | `1500` |
| `imageIsCircle` | Crops slide images into circles. | Disabled |
| `enableDiagnostics` | Writes lifecycle and image-loading details to the browser console. | Enabled |

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.
