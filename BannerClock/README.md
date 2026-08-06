# KM World Clock Banner

SPFx 1.5.1 Application Customizer for SharePoint Server Subscription Edition. It injects centrally managed header/footer content and a live multi-timezone clock banner into classic and modern pages.

## Features

- Render multiple clocks from a SharePoint list.
- Show digital clocks only or combine digital and analog clocks.
- Use 12-hour or 24-hour time and optionally show weekday names.
- Target classic pages, modern communication sites, and modern team sites.
- Load header HTML, footer HTML, settings, and additional assets from Site Assets.
- Point multiple sites at a central configuration location through `rootUrl`.
- Enable diagnostic console logging across the customizer, loader, and clock component.

## Configuration

### Extension Properties

| Property | Description |
|---|---|
| `rootUrl` | Optional root site/web application URL containing the shared Site Assets configuration. When omitted, the current site is used. |
| `enableDiagnostics` | Writes initialization, asset-loading, and clock-list details to the browser console. Enabled unless explicitly set to `false`. |

### Settings File

| Setting | Description |
|---|---|
| `spList` | SharePoint list containing clock locations and timezone values. |
| `digitalOnly` | Hides analog clock faces when enabled. |
| `hour12` | Uses 12-hour time with AM/PM instead of 24-hour time. |
| `displayDay` | Shows the weekday name. |
| `ClassicBannerId` / `ClassicFooterId` | Target element IDs on classic pages. |
| `ModernBannerComId` / `ModernFooterComId` | Target element IDs on modern communication sites. |
| `ModernBannerTeamId` / `ModernFooterTeamId` | Target element IDs on modern team sites. |

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.
