# SPS Full Width Control

SPFx 1.5.1 web part that makes its containing fixed-width SharePoint section span the available page width. The section is expanded as a shared layout container, so every web part in that section participates rather than only the Full Width Control's own host.

The web part has no configurable properties. While a page is being edited, it renders a compact selectable row that identifies the control and explains that the section is full width. On a published page, it renders no visible UI and collapses its own canvas host so it adds no spacing.

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.
