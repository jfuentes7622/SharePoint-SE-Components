---
description: "SharePoint SE Components repo conventions: SPFx 1.5.1/TS 2.4.2 build verification, diagnostics logging, property-pane version banner, and icon/text command button patterns."
applyTo: "**/*.ts,**/*.tsx"
---

# SharePoint SE Components conventions

This workspace contains many independent SPFx 1.5.1 (TS 2.4.2, old tslint) web part projects
(Accordion, Calendar, Carousel, Contact, DynamicForms, GridControl, LinkButton, ListControl,
Marquee, PrintControl, ReportForms, ScriptEditor, Tab, Tiles, FullWidthControl, BannerClock).
Each has its own package.json/tsconfig/tslint.json — verify conventions per-project rather than
assuming another project's style applies. There is a reference-only sibling repo
(`SharePoint-SPO-Components`, modern SPFx ~1.18) — never port code/deps from it directly; only use
it to see what features to replicate with SE-compatible APIs.

## Build & verification
- Treat `npx gulp bundle --ship` / `npx gulp package-solution --ship` as ground truth. `get_errors`
  can report false positives (e.g. `React.createElement` overload mismatches) and is unreliable
  when `node_modules` is missing — verify with `list_dir`/`Test-Path` first if results look wrong.
- TS 2.4.2 gotchas: no `unknown` type (use `any`); tslint `no-null-keyword` (return `undefined`, not
  `null`); ES5-only `lib` (no `Number.isFinite`/`MAX_SAFE_INTEGER` — use global `isFinite()`/
  `Number.MAX_VALUE`).
- `var`/`function` vs `let`/`const`/arrow functions: some projects' tslint enforce
  `no-var-keyword`/`prefer-const`/`no-function-expression`, others (e.g. ListControl) don't. Check
  the target project's own `tslint.json` and existing file style before choosing.
- Windows EPERM/rmdir during `gulp package-solution --ship`: clear stale ReadOnly attributes on
  `sharepoint/solution/debug/` (`Get-ChildItem <path> -Recurse -Force | % { $_.Attributes = 'Normal' }; attrib -R <path> /S /D`),
  then retry.
- `@pnp/spfx-property-controls` v1.1.0/1.5.1 `PropertyFieldColorPicker` has no `debounce`/`isHidden`
  props — omit them when porting newer configs.

## Diagnostics logging (apply to every project)
- Web part property `enableDiagnostics?: boolean`, `PropertyPaneCheckbox` in a "Diagnostics" group,
  `checked: this.properties.enableDiagnostics !== false`.
- Private `logDiagnostic(message: string): void` gated by
  `if (this.properties.enableDiagnostics === false) { return; } console.log('[ClassName] ' + message);`
  (or `this.props.enableDiagnostics` in React components). Genuine failures still use ungated
  `console.error`.
- Thread `enableDiagnostics` through any ConfigData/props objects so downstream services can gate
  their own logs.

## Property-pane version banner
- `const packageSolutionConfig: any = require('../../../config/package-solution.json');` at file
  top (after imports).
- `private getWebPartVersion(): string` reads `packageSolutionConfig.solution.version`, falling
  back to `this.context.manifest.version`.
- In `getPropertyPaneConfiguration()`: set `header.description: ''` and prepend a group
  `{ groupName: 'Version: ' + this.getWebPartVersion(), groupFields: [PropertyPaneLabel('propertyPaneVersionInfo', { text: ' ' })] }`.

## Icon/text command buttons (GridControl/ListControl pattern)
- Web part property `buttonDisplayMode?: string` (`'text' | 'iconText' | 'icon'`, default `'text'`),
  exposed via `PropertyPaneDropdown` in the Buttons property-pane group.
- Component helpers:
  - `getCommandButtonClass(className?: string)`: joins the passed class, `'gc-command-button'`, and
    `'gc-command-button-icon-only'` when mode is `'icon'`.
  - `renderCommandContent(iconName: string, label: string)`: returns plain `label` for `'text'`
    mode; otherwise renders
    `<span className="gc-command-content"><i className={'ms-Icon ms-Icon--' + iconName} aria-hidden="true"></i><span className={mode === 'icon' ? 'gc-visually-hidden' : ''}>{label}</span></span>`.
- Every actionable `<button>` (toolbar Refresh/Add/Edit/View/Delete, column filter toggle, filter
  Apply/Clear/Close, date-picker prev/next/calendar, pagination Prev/Next) gets
  `className={this.getCommandButtonClass(...)}`, `title`/`aria-label` set to the same string used
  as the label, and content from `renderCommandContent(iconName, label)`.
- CSS classes `.gc-command-content`, `.gc-command-button-icon-only`, `.gc-visually-hidden`,
  `.gc-compact-icon-button` must exist in the project's CSS — copy them from GridControl.css/
  ListControl.css rather than reinventing.
- Rely on the SharePoint page's built-in Fabric/Fluent `ms-Icon` font. Do not add
  `initializeIcons()` or `@uifabric/icons` — not used anywhere in this repo and not needed since
  classic SharePoint chrome already loads the icon font.
