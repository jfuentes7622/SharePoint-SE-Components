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

## Usage

1. Add **SPS Full Width Control** anywhere inside the fixed-width section that should expand.
2. Keep the control in the same section as the Calendar, Grid, List, form, or other content that needs the extra width.
3. Use the visible marker to select, move, or remove the control while the page is in edit mode.
4. Preview or publish the page to confirm that the marker disappears and the section spans the available page width.

## Properties and common configuration

The control has no configurable properties. It expands the shared containing section, accounts for the visible SharePoint property pane, and recalculates on layout or window changes. It reserves a small left gutter and SharePoint Server right gutter so content does not collide with navigation or page chrome.

Only the containing section is affected. Place a separate Full Width Control in each section that must expand. Removing the web part restores the original inline layout styles.

## Common scenarios

- Give a wide Calendar, List Control, or Grid Control more horizontal workspace.
- Expand a section containing a report form and related item selector.
- Build a full-width image carousel or tile navigation band.
- Keep another section on the same page at its normal constrained width.
