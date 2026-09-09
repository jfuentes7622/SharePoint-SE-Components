# SharePoint Dynamic Form

<!-- 
  Keywords: SharePoint, SPFx, Dynamic Form, Form Builder, Web Part, React, SharePoint Framework, 
  Form Engine, Visual Designer, No-code Form, SharePoint Online, TypeScript
-->

<p align="center">
  <strong>A powerful SPFx dynamic form builder for SharePoint with visual designer</strong><br>
  <em>SharePoint Server Subscription Edition (SPFx 1.5.1)</em>
</p>

<p align="center">
  <a href="#features">Features</a> •
  <a href="#getting-started">Quick Start</a> •
  <a href="#usage-guide">Usage</a> •
  <a href="#project-structure">Structure</a> •
  <a href="#faq">FAQ</a>
</p>

<p align="center">
  English
</p>

---

## Demo

[demo-recording.mp4](docs/assets/demo-recording.mp4)

---

## What is this project?

SharePoint Dynamic Form is a SharePoint Framework (SPFx) web part that provides a **no-code form builder** for SharePoint lists. Users can visually design forms with drag-and-drop fields, configure conditional logic, and render dynamic forms without writing any code.

### Use Cases

- Business process forms
- Data collection forms
- Approval request forms
- Employee onboarding forms
- Survey and feedback forms
- Multi-step wizard forms

---

## Features

### Visual Form Designer

| Feature | Description |
|---------|-------------|
| Drag-and-Drop | Add fields by clicking `+` button |
| WYSIWYG Preview | See form layout in real-time |
| Field Configuration | Configure field properties inline |
| Layout Control | Grid and stack layout options |

### Rich Field Types (15+)

| Category | Field Types |
|----------|-------------|
| Text | Text, Multiline, Rich Text |
| Number | Number, Integer |
| Date | DateTime, Date Only, Time Only |
| Selection | Dropdown, MultiSelect, Lookup, Boolean |
| People | Person (single/multiple) |
| Metadata | Taxonomy, Term Store |
| Media | Image, URL, Attachment, Document-library upload |

### Field State And Label Controls

- Every field is enabled by default. Clear **Field enabled** in the designer to keep the field visible while preventing user input.
- Disabled fields preserve their configured/default values and are also non-editable in the designer preview.
- Fields hidden by a conditional visibility rule are removed from the form grid so the remaining fields reflow without empty gaps. Hidden fields do not enforce required validation.
- Labels can be placed above or to the left of any field.
- Boolean fields additionally support labels below the checkbox or after it on the right.
- Boolean fields no longer add automatic Yes/No text. Use **Checkbox text** to provide optional text beside the checkbox.

### Date And Time Options

DateTime fields support three per-field display formats:

| Format | Input | Purpose |
|---|---|---|
| `dateOnly` | Date picker | Captures a date without a time control. |
| `dateTime` | Date and time picker | Captures both date and time. This is the default format. |
| `timeOnly` | Time picker | Captures a time and stores it using an internal date anchor. |

DateTime and Time Only fields can interpret values as **UTC (Zulu)** or **Local browser time**. UTC is the default when `timeZone` is not set. Existing forms that require the previous browser-local interpretation should select Local browser time for those fields.

### Permission Messages

The designer's **Form messages** settings include a single customizable permission-denied message. It is used consistently when access is denied based on the form list, a lookup list, or a lookup item, and when a submission is blocked by the same permission check.

The default is: **You do not have permission to access or modify this record.**

Detailed permission-source diagnostics remain in the browser console and are not exposed in the user-facing message. The warning also follows the effective form mode: New checks Add permission, while Edit and View check Edit permission.

### Conditional Logic

- **Visibility Rules**: Show/hide fields based on conditions
- **Required Rules**: Make fields required dynamically
- **Readonly Rules**: Control field editability
- **Value Actions**: Auto-set field values

The designer also provides a dedicated **Conditional styling** workspace. Each ordered rule compares a source field with a value and either applies styles to a target field or sets its visibility. Style actions can control background, text, border, corner radius, font size, and font weight. Use **Move up** and **Move down** to set precedence: matching styles are merged in rule order, and the last matching rule wins when rules set the same style or visibility property.

Required and targeted validation failures keep their existing message and add a red border around the complete field. After a failed save, validation reruns as values change, including cross-field rules. When the field becomes valid, the red validation override is removed and its configured or conditional border is rendered again.

### Multi-Step Forms

- Step-by-step wizard navigation
- Progress indicator
- Per-step validation
- Custom step titles and descriptions
- Move container left/right controls reorder wizard steps without deleting or recreating them.
- Form-level and per-container themes can independently set description font size, family, weight, and color.

### SharePoint Integration

- Direct binding to SharePoint lists and document libraries
- Automatic field type mapping
- New/Edit/View modes
- Attachment upload support
- Document-library creation uploads one primary file first and then saves its metadata fields
- Document-library Edit mode updates metadata and can replace the current file content while preserving its file name and item ID
- People picker integration
- Term store (taxonomy) support

For a document library, add the synthetic **Document** field from the designer palette. New forms require exactly one file. The file is uploaded to the library root and its resulting list-item ID is used for the metadata update. The primary document cannot be deleted from the form; use SharePoint library actions when the whole document item must be deleted. Ordinary SharePoint lists continue to use multi-file list attachments.

### Dynamic Data Integration

- Connect to a ListControl, GridControl, or Calendar web part on the same page.
- Consume the source's selected SharePoint item ID and selected mode (`new`, `edit`, or `view`).
- Use the dynamic ID as this form's SharePoint item ID to open the same record, or clear **Use dynamic item ID as form item ID** and choose a writable target column to find or create a related record.
- The target can be a supported writable text, number, Boolean, DateTime, lookup, or person column; it is not limited to lookup columns. Lookup and person targets compare or store the numeric SharePoint item ID in the appropriate format.
- When dynamic-ID-as-item-ID is off, the configured `itemid` URL query value is also applied to the selected target column instead of being consumed as the record's SharePoint `ID`.
- Calendar selections publish `view` mode; ListControl actions publish the mode associated with the selected Add, Edit, or View command.
- Dynamic Forms also publishes its most recently saved item ID and View mode, allowing another Dynamic Forms instance to consume the result.

### Redirects And Same-Page Wizards

- Submit and Cancel navigation can target a discovered Site Pages/publishing page or a manually entered same-site URL.
- Submit can append the newly created or updated item ID to the destination using a configurable query parameter such as `itemid`. Cancel can optionally append the current record ID in the same way.
- For a page-to-page process, configure the next page as the submit destination and configure its form to consume the same query parameter.
- For a same-page process, give each form a unique **Form name for linking**, then select its **Previous form** and **Next form**. A form with a previous form stays hidden in read mode until its predecessor saves successfully and passes the saved ID.
- On each form after the first, turn off **Use dynamic ID as form item ID** and select the writable **Target column** that should receive the previous form's saved ID. The target can be a text, number, lookup, person, Boolean, or DateTime column; it does not need to be a lookup.
- Wizard Back/Next controls are available only in preview or on the published page. Edit mode keeps all forms visible for configuration and does not run wizard navigation.
- A successful save advances to the configured next form when no page redirect is configured. Page redirects take precedence when both are present.
- **Back** appears on an active form whenever a previous form is linked, including while the successor is adding or editing its record. Each successor preserves its predecessor's record ID independently, so saving or editing the successor does not replace the Back destination. Back reopens the previous saved record using the ID carried by the wizard, taking precedence over the predecessor's normal dynamic ID and mode bindings. In View mode, **Next** appears when a next form is linked and passes the current record ID forward.
- Filter-resolved Edit forms save against the record that was actually loaded. An Edit submission with no valid record ID is blocked rather than creating a new item, and wizard advancement occurs only after SharePoint accepts the save and the saved record can be loaded again.

---

## Migration Status

This repository is targeted to the SPFx line supported by SharePoint Server Subscription Edition.

Current status:
- Dependency and build baselines are aligned to SPFx 1.5.1 (on-prem profile).
- Packaging and serve configuration are set for local/on-prem development flow.

## Getting Started

### Prerequisites

| Requirement | Version |
|-------------|---------|
| Node.js | 10.24.x |
| npm | 6.x |
| SharePoint Server | SharePoint Server Subscription Edition |
| SPFx target | 1.5.1 (on-prem) |

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd SharePointDForm

# Install dependencies
npm install

# Start development server
npm start
```

### Build for Production

```bash
npm run build
```

The solution package (`.sppkg`) will be generated in `sharepoint/solution/`.

---

## Usage Guide

### Step 1: Add Web Part to Page

1. Edit a SharePoint page
2. Add the "SharePoint Dynamic Form" web part
3. Configure the web part properties

### Step 2: Configure List and Mode

| Property | Description |
|----------|-------------|
| List Name | Target SharePoint list |
| Form Mode | New / Edit / View |
| Item ID | Specific item (for Edit/View) |
| Dynamic Item ID | Selected item ID from a connected ListControl, GridControl, or Calendar web part |
| Dynamic Mode | New / Edit / View mode from the connected source |
| Dynamic Target Column | Writable column populated or filtered using the connected or URL value when dynamic-ID-as-item-ID is off |

### Step 3: Design the Form

1. Enable "Designer Mode" in property pane
2. Left panel shows available SharePoint fields
3. Click `+` to add fields to the form
4. Configure field properties (label, required, visible)
5. Arrange fields in grid layout
6. For multi-step forms, use the left/right controls beside the selected container to change step order.
7. Open **Conditional styling** to add, edit, delete, or reorder field style and visibility rules.

### Step 4: Configure Buttons and Actions

| Setting | Purpose |
|---------|---------|
| Submit Button Label | Customize submit button text |
| Show Cancel Button | Add cancel option |
| Submit Redirect URL | Redirect after submit |
| Submit Success Message | Show confirmation |

### Configuration Import/Export

Export and import form configurations as JSON for:
- Backup and restore
- Environment migration
- Version control
- Template sharing

---

## Project Structure

```
src/
├── designer/                    # Form Designer Module
│   ├── index.ts                     # Module entry
│   ├── components/
│   │   ├── FormDesigner.tsx         # Main designer component
│   │   ├── DesignerCanvas.tsx       # Drag-drop canvas
│   │   ├── DesignerFieldRenderer.tsx # Field rendering in designer
│   │   ├── PropertyPanel.tsx        # Field property editor
│   │   └── ConditionBuilder.tsx     # Conditional logic builder
│   └── controls/
│       ├── FieldPalette.tsx         # Available fields panel
│       └── FieldLayout.tsx          # Field layout controls
│
├── formEngine/                  # Form Engine Core
│   ├── index.ts                     # Module entry
│   ├── core/
│   │   ├── types.ts                 # TypeScript type definitions
│   │   ├── FormStateManager.ts      # State management
│   │   └── ValidationEngine.ts      # Field validation
│   ├── components/
│   │   ├── FormRenderer.tsx         # Main form renderer
│   │   ├── FormRenderer.module.scss # Form renderer styles
│   │   ├── FieldContainer.tsx       # Field wrapper container
│   │   ├── StepRenderer.tsx         # Step container
│   │   └── FormStepper.tsx          # Navigation stepper
│   ├── fields/                      # Field Components (15+ types)
│   │   ├── index.tsx                # Field registry
│   │   ├── BaseField.tsx            # Base field component
│   │   ├── TextField.tsx
│   │   ├── MultilineField.tsx
│   │   ├── RichTextField.tsx        # Rich text (Quill editor)
│   │   ├── NumberField.tsx
│   │   ├── DateTimeField.tsx
│   │   ├── DropdownField.tsx
│   │   ├── MultiSelectField.tsx
│   │   ├── LookupField.tsx
│   │   ├── PersonField.tsx
│   │   ├── BooleanField.tsx
│   │   ├── UrlField.tsx
│   │   ├── ImageField.tsx
│   │   ├── TaxonomyField.tsx
│   │   ├── AttachmentField.tsx
│   │   ├── NewLineField.tsx
│   │   ├── PnpControlCompat.css     # PnP controls style patch
│   │   ├── RichTextField.css        # Rich text editor styles
│   │   └── AttachmentField.css      # Attachment field styles
│   ├── hooks/
│   │   └── useDebounce.ts           # Debounce hook
│   ├── utils/
│   │   └── odata/                   # OData expression engine
│   │       ├── ODataLexer.ts        # OData tokenizer
│   │       ├── ODataParser.ts       # OData parser
│   │       ├── ODataEvaluator.ts    # OData expression evaluator
│   │       └── index.ts
│   └── data/
│       └── SharePointDataSource.ts  # SharePoint API layer
│
├── templates/                   # Form Templates
│   ├── index.ts
│   └── formTemplates.ts             # Predefined form templates
│
└── webparts/sharePointDynamicForm/  # SPFx Web Part
    ├── SharePointDynamicFormWebPart.ts       # Entry point
    ├── SharePointDynamicFormWebPart.manifest.json
    ├── components/
    │   ├── SharePointDynamicForm.tsx         # Main component
    │   ├── ISharePointDynamicFormProps.ts    # Props interface
    │   └── SharePointDynamicForm.module.scss # Component styles
    ├── propertyPane/
    │   ├── PropertyPaneConfigIO.ts           # Config import/export pane
    │   └── ConfigIOControl.tsx               # Config IO React control
    ├── utils/
    │   └── configIO.ts                       # Config read/write utilities
    ├── assets/                                # Web part icons
    └── loc/                                   # Localization (en-us)
```

---

## FAQ

### General Questions

<details>
<summary><strong>What SharePoint versions are supported?</strong></summary>

This project is built on SPFx 1.5.1 for SharePoint Server Subscription Edition's on-premises SPFx profile.
</details>

<details>
<summary><strong>Is coding required to use this?</strong></summary>

No coding is required. The visual designer allows you to build forms through a user-friendly interface.
</details>

<details>
<summary><strong>Can I customize field styles?</strong></summary>

Yes, you can modify the CSS files in `src/formEngine/fields/*.css` or use Fluent UI theming.
</details>

### Technical Questions

<details>
<summary><strong>PnP controls have style issues, how to fix?</strong></summary>

This project includes compatibility patches:
- Webpack CSS hash patch: `config/spfx-customize-webpack.js`
- Fallback styles: `src/formEngine/fields/PnpControlCompat.css`

Clear browser cache after deployment.
</details>

<details>
<summary><strong>How do I add custom field types?</strong></summary>

1. Create a new component in `src/formEngine/fields/`
2. Extend `BaseField` component
3. Register in `src/formEngine/fields/index.tsx`
4. Add type definition in `src/formEngine/core/types.ts`
</details>

<details>
<summary><strong>How to handle form validation?</strong></summary>

The form engine includes a `ValidationEngine` that supports:
- Required field validation
- Min/Max length validation
- Pattern (regex) validation
- Custom validation rules
</details>

### Deployment Questions

<details>
<summary><strong>How to deploy to SharePoint?</strong></summary>

1. Run `npm run build`
2. Upload `.sppkg` from `sharepoint/solution/` to App Catalog
3. Add the app to your site
</details>

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Fields not loading | Check list permissions |
| Styles missing | Clear browser cache |
| Submit fails | Check required fields |
| Designer blank | Select a valid list |

---

## Tech Stack

| Category | Technology |
|----------|------------|
| Framework | SPFx 1.5.1, React 15.6.2 |
| UI Library | Office UI Fabric React 5.135.6 |
| SharePoint | SPHttpClient and sp-pnp-js 3.0.10 |
| Controls | @pnp/spfx-controls-react 1.1.0 |
| Build | Gulp 3 and the SPFx 1.5.1 build rig |
| Language | TypeScript 2.4.2 |

---

## Contributing

Contributions are welcome! 

---

## License

MIT License
---

<p align="center">
  <strong>SharePoint Dynamic Form</strong> - Build dynamic forms without code<br><br>
  Made with ❤️ for SharePoint developers
</p>

## Usage

1. Add **SPS Dynamic Forms** to a page, select the target SharePoint list or document library, and choose New, Edit, or View mode.
2. Open the visual form designer to add, order, and configure fields, containers, steps, validation, and conditional behavior.
3. Choose how an existing item is resolved: fixed ID, URL parameter, filter, or a dynamic connection from Calendar, List Control, or Grid Control.
4. Configure permissions, submit/cancel behavior, labels, and buttons; test each supported mode with an account that matches the intended permissions.
5. Publish only after lookup, person, attachment or document upload, date/time, validation, and redirect behavior have been exercised against the target list or library.

## Properties and common configuration

- **List and mode:** `listName` identifies the target; `mode` controls create, update, or read-only behavior.
- **Designer schema:** `formSchemaJson` stores the complete visual design. Maintain it through the designer unless an advanced migration requires reviewed JSON changes.
- **Item resolution:** fixed item ID, `itemIdQueryParam`, dynamic ID/mode references, target lookup fields, and `filterJson` support direct and connected workflows.
- **Rules and defaults:** visual designers maintain filters, default values, conditional state, required rules, and advanced cross-field validation expressions.
- **Permissions:** list- or item-level checks and optional lookup-based scope control whether a user may load or submit the form.
- **Actions and appearance:** mode-specific submit labels, cancel/redirect behavior, success messages, label position, field descriptions, and button styling control the runtime experience.
- **Layout:** `forceFullWidth` or fixed-width settings adapt the form to the page. Diagnostic settings expose dynamic-source, load, validation, and submission details in the console.

Field internal names and data types must match the SharePoint list. Client-side visibility is a user-experience rule, not a replacement for SharePoint permissions or list validation.

## Common scenarios

- A multi-step onboarding, request, or approval form with conditional sections.
- A List Control or Grid Control selection that opens the chosen record in Edit or View mode.
- A Calendar selection that displays or updates the selected event without leaving the page.
- A parent form with embedded Grid Control line items or an embedded read-only List Control.
- A URL-driven report/edit page using `itemid` plus return and post-submit navigation.
