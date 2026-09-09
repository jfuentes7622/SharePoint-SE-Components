# SPS Form Navigator

Grouped, direct navigation for multiple SPS Dynamic Forms and SPS Report Forms web parts on the same SharePoint page.

## Usage

1. Deploy SPS Dynamic Forms 1.0.0.96 or later and/or SPS Report Forms 1.0.0.65 or later, then deploy this package.
2. Add the Dynamic Forms and Report Forms to navigate, plus one SPS Form Navigator, to the same page.
3. Give every Dynamic Form and Report Form a unique navigation name.
4. Open the Form Navigator property pane and select **Manage categories** to add and order categories.
5. Select **Manage forms**, choose each discovered form or report and its category, then set its navigation label and behavior.
6. Select a default navigation policy and configure appearance.

The navigator coordinates targets through page events. It does not move their DOM nodes. In page read mode, selecting a node activates that target and hides the other configured forms and reports. Report Forms remain read-only and continue resolving records from their own fixed ID, URL, filter, or dynamic-data configuration.

## Categories and forms

Categories and forms are managed with collection editors; no JSON editing is required. Category rows contain a unique name and an **Initially collapsed** setting. Form rows provide dropdowns for compatible Dynamic Forms and Report Forms discovered on the current page and the configured categories, plus a navigation label, required setting, and optional policy override.

Drag collection rows to set category display order and global form workflow order. Optional forms do not block later sequential nodes. Existing pages configured with the earlier JSON properties migrate those values into the collections automatically.

## Navigation policies

- **Free navigation:** every configured form is available.
- **Completed forms only:** only the current form and previously completed forms are available.
- **Sequential unlock:** all earlier required forms must be complete.
- **Locked forms open read-only:** follows sequential order, but a locked form with an existing saved item can be previewed in View mode.

When enabled, dirty-form confirmation prevents accidental loss of unsaved input. Status indicators show current, complete, available, locked, and error states.

## Appearance

The property pane configures vertical, horizontal-category, or tree layout; surface/category/form colors; active, complete, locked, and error colors; typography; bold settings; padding; category/form gaps; tree indentation; borders; and corner radius.

## Build

```powershell
npm install
npx gulp clean
npx gulp bundle --ship
npx gulp package-solution --ship
```

The package is generated at `sharepoint/solution/sps-form-navigator.sppkg`.
