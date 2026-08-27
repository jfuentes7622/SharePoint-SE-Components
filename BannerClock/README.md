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

## Usage

1. Deploy the package and register the application customizer at the required web or site scope.
2. Set `rootUrl` to the site that hosts the shared configuration assets.
3. Create the configured SharePoint list with `Title` and `Timezone` columns and add one row per location.
4. Place the settings and optional header/footer assets in Site Assets, then verify classic, modern team, and communication pages that are in scope.

## Properties and common configuration

- **Extension property:** `rootUrl` is the base site used to resolve shared settings, assets, and list data.
- **Settings file:** `spList`, `digitalOnly`, `hour12`, and `displayDay` control the source list and clock format.
- **Placeholders:** the classic, communication-site, and team-site banner/footer element IDs determine where content is injected.
- **Diagnostics:** `enableDiagnostics` enables browser-console lifecycle, configuration, and data-loading messages.

This is an application customizer rather than a page web part. Administrators configure it centrally; page authors do not receive a property pane on each page.

## Common scenarios

- A global operations banner showing the local time for regional offices.
- A service desk page showing support-center time zones and current business hours.
- A centrally managed header or footer shared across many sites.
- A 12-hour executive display or 24-hour operations display controlled from one settings file.
