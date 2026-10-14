# Customer Update - October 14, 2026

This release improves data freshness, visual integration with SharePoint site themes, and deployment consistency across the SharePoint Server Subscription Edition component suite.

## What changed

### Fresher SharePoint list and field information

Grid Control, List Control, Dynamic Forms, and Report Forms now request current SharePoint metadata instead of allowing previously cached field definitions to be reused.

Customers should see the following improvements:

- Newly added Choice and MultiChoice options appear after the page or control is refreshed.
- Dynamic Forms and Report Forms refresh saved choice-field options from the current SharePoint column definition.
- List Control requests every configured display column, including columns that are not part of the selected SharePoint view.
- List Control reloads its field definitions when its saved column or grouping configuration changes.

### SharePoint site-theme support

Every web part that already offered configurable colors or font families now includes an **Appearance mode** setting:

- **Manual styling** preserves the existing configured appearance and remains the default for existing and new installations.
- **Site theme only** uses SharePoint semantic colors and typography. Theme-controlled manual settings are unavailable while this mode is active, but their saved values are retained.
- **Site theme with overrides** starts with the SharePoint site theme and allows non-empty manual settings to override individual theme values.

The update covers Accordion, Calendar, Carousel, Contacts, Dynamic Forms, Form Navigator, Grid Control, Link Button, List Control, Marquee, Print Control, Report Forms, Search Control, Tabs, and Tiles.

Existing web parts continue to use **Manual styling** unless an author deliberately selects a site-theme mode. Layout, sizing, spacing, font weight, corner styles, conditional formatting, and item-specific styling remain unchanged unless they directly depend on a themed color or font family.

### Updated deployment packages

All 15 affected solutions received new npm, solution, and feature versions and were rebuilt as production packages. Feature versions now match their corresponding solution versions.

| Component | Solution and feature version | Deployment package |
|---|---:|---|
| Accordion | 1.0.0.22 | `sps-accordion.sppkg` |
| Calendar | 1.1.0.39 | `sps-calendar.sppkg` |
| Carousel | 1.0.0.38 | `sps-carousel.sppkg` |
| Contacts | 3.1.0.20 | `sps-contacts.sppkg` |
| Dynamic Forms | 1.0.0.112 | `sps-dynamic-form.sppkg` |
| Form Navigator | 1.0.0.4 | `sps-form-navigator.sppkg` |
| Grid Control | 1.0.0.87 | `sps-grid-control.sppkg` |
| Link Button | 2.0.2.15 | `sps-linkbutton.sppkg` |
| List Control | 1.0.0.109 | `SPS-list-control.sppkg` |
| Marquee | 2.1.0.25 | `sps-marquee.sppkg` |
| Print Control | 1.0.0.17 | `sps-print-control.sppkg` |
| Report Forms | 1.0.0.83 | `sps-report-forms.sppkg` |
| Search Control | 1.0.0.14 | `sps-search-control.sppkg` |
| Tabs | 3.1.0.36 | `sps-tabs.sppkg` |
| Tiles | 2.1.0.18 | `sps-tiles.sppkg` |

## Customer impact

- No automatic visual change is applied to existing web parts.
- Existing manual styling values are retained when switching appearance modes.
- No list schema or stored SharePoint data is modified.
- Site-theme modes can be adopted one web part at a time.
- Updated packages must be uploaded to the SharePoint App Catalog and approved using the customer's normal deployment process.

## Recommended deployment checks

1. Replace each affected package in a test App Catalog.
2. Confirm that SharePoint recognizes the package as an upgrade.
3. Test one existing page before changing its Appearance mode; it should retain its current manual styling.
4. Test **Site theme only** and **Site theme with overrides** against representative light, dark, and custom themes.
5. For Grid Control, List Control, Dynamic Forms, and Report Forms, confirm that recently added choice options and configured columns appear after refresh.
6. Promote the packages through the customer's normal staging and production process.

All affected projects completed production bundling and SharePoint package generation successfully.
