# SPS Marquee

SPFx 1.5.1 scrolling-announcement web part for SharePoint Server Subscription Edition. The marquee is injected into the configured SharePoint page header area so it can span beyond the web part canvas.

## Features

- Display a static announcement or cycle through messages from a SharePoint list.
- Select the list column that contains message text.
- Configure how long each list message remains active.
- Customize colors, typography, height, speed, and scroll direction.
- Enable or disable the marquee without removing the web part.
- Render list-sourced content as plain text so stored HTML is not interpreted.
- Enable diagnostic console logging for troubleshooting.

## Configuration

### Content

| Property | Description | Default |
|---|---|---|
| `description` | Static scrolling text. The value is limited to 250 characters and HTML tags are rejected. It is also used as the fallback when a list source is not fully configured. | Empty |
| `listName` | Optional SharePoint list containing dynamic messages. | Empty |
| `messageField` | Column whose values are displayed. Available columns are loaded after selecting a list. | Empty |
| `messageDuration` | Seconds each list message is displayed before advancing to the next message. | `8` |

When both `listName` and `messageField` are selected, the web part loads up to 100 list items in item-ID order and cycles through their non-empty message values. Clear either setting to return to the static announcement.

### Appearance And Behavior

| Property | Description | Default |
|---|---|---|
| `marqueeActive` | Shows or hides the marquee. | Enabled |
| `marqueeBackColor` | Marquee background color, including optional alpha transparency. | `#333333` |
| `marqueeTextColor` | Announcement text color, including optional alpha transparency. | `#FAFAFA` |
| `fontFamily` | Announcement font family. | `Segoe UI` |
| `fontSize` | CSS font size applied to announcement text. | `13px` |
| `fontStyle` | Normal, italic, or oblique text. | `normal` |
| `fontBold` | Uses bold announcement text. | Enabled |
| `marqueeHeight` | CSS height/minimum height for the marquee bar. | `32px` |
| `scrollSpeed` | Seconds required for one complete pass; lower values move faster. | `50` |
| `scrollDirection` | Scrolls left or right. | `left` |
| `enableDiagnostics` | Writes lifecycle and data-loading details to the browser console. | Enabled |

## Requirements

- SharePoint Server Subscription Edition
- SharePoint Framework 1.5.1
- Node.js 10.24.x and npm 6.x for local builds
- Read access to the selected list and message column when list-backed mode is used

## Build And Package

From the `Marquee` folder:

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.

## Disclaimer

This code is provided as-is without warranty of any kind, either express or implied.
