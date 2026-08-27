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
| Media | Image, URL, Attachment |

### Field State And Label Controls

- Every field is enabled by default. Clear **Field enabled** in the designer to keep the field visible while preventing user input.
- Disabled fields preserve their configured/default values and are also non-editable in the designer preview.
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

### Multi-Step Forms

- Step-by-step wizard navigation
- Progress indicator
- Per-step validation
- Custom step titles and descriptions

### SharePoint Integration

- Direct binding to SharePoint lists
- Automatic field type mapping
- New/Edit/View modes
- Attachment upload support
- People picker integration
- Term store (taxonomy) support

### Dynamic Data Integration

- Connect to a ListControl, GridControl, or Calendar web part on the same page.
- Consume the source's selected SharePoint item ID and selected mode (`new`, `edit`, or `view`).
- Use the dynamic ID as this form's item ID to open the same record, or use it to create/filter related records through a parent lookup column.
- Parent lookup choices are limited to lookup columns from the selected form list. Single-value and multi-value lookups receive the numeric SharePoint item ID in the correct payload format.
- Calendar selections publish `view` mode; ListControl actions publish the mode associated with the selected Add, Edit, or View command.

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
| Parent Lookup Column | Lookup column populated or filtered using the connected item ID |

### Step 3: Design the Form

1. Enable "Designer Mode" in property pane
2. Left panel shows available SharePoint fields
3. Click `+` to add fields to the form
4. Configure field properties (label, required, visible)
5. Arrange fields in grid layout

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

1. Add **SPS Dynamic Forms** to a page, select the target SharePoint list, and choose New, Edit, or View mode.
2. Open the visual form designer to add, order, and configure fields, containers, steps, validation, and conditional behavior.
3. Choose how an existing item is resolved: fixed ID, URL parameter, filter, or a dynamic connection from Calendar, List Control, or Grid Control.
4. Configure permissions, submit/cancel behavior, labels, and buttons; test each supported mode with an account that matches the intended permissions.
5. Publish only after lookup, person, attachment, date/time, validation, and redirect behavior have been exercised against the target list.

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
