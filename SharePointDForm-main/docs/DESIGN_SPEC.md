# SPFx Dynamic Form Engine - Design Specification

## 1. Requirements Summary

| Dimension | Requirement |
|-----|------|
| **Driving conditions** | Field value dependencies, list item state |
| **Config storage** | Web Part Properties |
| **Complexity** | Enterprise-grade (linear steps, conditional branching, custom field types) |
| **Designer** | Visual drag-and-drop designer, embedded in Web Part edit mode |
| **Form modes** | New, Edit, View |
| **Template reuse** | Supports creating from templates, import/export |
| **Data source** | SharePoint List only |
| **Validation** | Synchronous validation, no async required |
| **Condition syntax** | SharePoint OData filter syntax (consistent with the native View Filter) |

---

## 2. Overall Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    SPFx Dynamic Form Engine                     │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  ┌──────────────┐      ┌──────────────┐    ┌─────────────────┐ │
│  │  Designer    │──────│  Renderer    │────│  Data Layer     │ │
│  │              │      │              │    │  SharePoint     │ │
│  │              │      │              │    │  DataSource     │ │
│  │ • Drag fields│      │ • Linear     │    │                 │ │
│  │ • Property   │      │   steps      │    │ • Get Items     │ │
│  │   config     │      │ • Conditional│    │ • Create Item   │ │
│  │ • Condition  │      │   display    │    │ • Update Item   │ │
│  │   builder    │      │ • Field      │    │                 │ │
│  │ • Preview /  │      │   dependency │    │                 │ │
│  │   test       │      │ • OData      │    │                 │ │
│  │              │      │   conditions │    │                 │ │
│  └──────────────┘      └──────────────┘    └─────────────────┘ │
│           │                        │                  │        │
│           └────────────────────────┼──────────────────┘        │
│                                  ▼                            │
│                    ┌──────────────────────────┐               │
│                    │   Web Part Properties    │               │
│                    │   (FormSchema JSON)      │               │
│                    └──────────────────────────┘               │
│                                                                 │
├─────────────────────────────────────────────────────────────────┤
│  Core Modules                                                    │
├─────────────────────────────────────────────────────────────────┤
│  • ODataConditionEngine  OData condition expression engine       │
│  • ValidationEngine     Synchronous validation engine             │
│  • FormStateManager     Form state management & field dependency  │
│  • StepEngine           Linear step engine                        │
│  • FieldRegistry        Field type registry                       │
│  • SharePointDataSource SharePoint data source wrapper             │
└─────────────────────────────────────────────────────────────────┘
```

---

## 3. Core Data Structures

### 3.1 Form Mode

```typescript
type FormMode = 'new' | 'edit' | 'view';
```

### 3.2 Field Type

```typescript
type FieldType =
  | 'text'           // Single-line text
  | 'multiline'      // Multi-line text
  | 'number'         // Number
  | 'datetime'       // Date/time
  | 'dropdown'       // Dropdown selection (Choice)
  | 'multiselect'    // Multi-select (MultiChoice)
  | 'lookup'         // Lookup field
  | 'person'         // Person picker
  | 'boolean';       // Yes/No
```

### 3.3 OData Filter Expression

```typescript
// Uses SharePoint OData syntax
type FilterExpression = string;
```

### 3.4 Form Configuration (FormSchema)

```typescript
interface FormSchema {
  // Basic info
  id: string;
  name: string;
  description?: string;

  // Form mode
  mode: FormMode;

  // SharePoint config
  listName: string;           // Target List name or ID
  itemId?: number;            // Item ID in edit mode (dynamic at runtime)

  // Linear steps
  steps: FormStep[];

  // Submit config
  submitButtonLabel?: string;
  showCancelButton?: boolean;
  onSubmitMessage?: string;

  // Theme (optional)
  theme?: {
    layout?: 'stack' | 'grid';
    columns?: number;
  };
}
```

### 3.5 Step Configuration (FormStep)

```typescript
interface FormStep {
  id: string;
  title: string;
  description?: string;

  // Field list
  fields: FormField[];
}
```

### 3.6 Field Configuration (FormField)

```typescript
interface FormField {
  // Basic properties
  id: string;
  type: FieldType;
  label: string;

  // Bound to SP field
  fieldName: string;          // SP internal field name

  // Display control - uses OData filter syntax
  visible?: FilterExpression;
  required?: FilterExpression;
  readOnly?: FilterExpression;

  // Field dependency
  onChange?: FieldAction[];

  // Validation rules
  validation?: ValidationRule[];

  // UI config
  config?: FieldConfig;
}
```

### 3.7 Field Configuration (FieldConfig)

```typescript
interface FieldConfig {
  // Text field
  maxLength?: number;
  placeholder?: string;

  // Number field
  min?: number;
  max?: number;
  decimals?: number;

  // Date/time
  displayFormat?: 'dateOnly' | 'dateTime';

  // Dropdown/multi-select
  choices?: string[];
  allowFillIn?: boolean;

  // Lookup field
  lookupList?: string;
  lookupField?: string;
}
```

### 3.8 Field Dependency Actions

```typescript
type FieldAction =
  | { type: 'show'; target: string; condition?: FilterExpression }
  | { type: 'hide'; target: string }
  | { type: 'set-value'; target: string; value: any }
  | { type: 'clear'; target: string }
  | { type: 'enable'; target: string }
  | { type: 'disable'; target: string };
```

### 3.9 Validation Rules

```typescript
interface ValidationRule {
  type: 'required' | 'minLength' | 'maxLength' | 'min' | 'max' | 'pattern' | 'custom';
  value?: any;
  message: string;
  applyWhen?: FilterExpression;  // When to apply this validation
}
```

---

## 4. OData Filter Syntax Reference

### 4.1 Comparison Operators

| Operator | Description | Example |
|-------|------|------|
| `eq` | Equal to | `Status eq 'Pending'` |
| `ne` | Not equal to | `Status ne 'Completed'` |
| `gt` | Greater than | `Amount gt 1000` |
| `ge` | Greater than or equal to | `Amount ge 1000` |
| `lt` | Less than | `Amount lt 10000` |
| `le` | Less than or equal to | `Amount le 10000` |

### 4.2 String Operators

| Operator | Description | Example |
|-------|------|------|
| `contains` | Contains | `contains(Title, 'urgent')` |
| `startswith` | Starts with | `startswith(Title, 'RE:')` |

### 4.3 Logical Operators

| Operator | Description | Example |
|-------|------|------|
| `and` | And | `Status eq 'Pending' and Amount gt 1000` |
| `or` | Or | `Status eq 'Pending' or Status eq 'Draft'` |
| `not` | Not | `not(Status eq 'Completed')` |

### 4.4 Null Checks

| Operator | Description | Example |
|-------|------|------|
| `eq null` | Is null | `ApprovalDate eq null` |
| `ne null` | Is not null | `ApprovalDate ne null` |

### 4.5 Expression Examples

```typescript
// Simple condition
"Status eq 'Pending'"

// Compound condition
"Status eq 'Pending' and Department eq 'IT'"

// Numeric comparison
"Amount ge 10000 and Amount lt 50000"

// String contains
"contains(Title, 'urgent')"

// Multi-value condition
"Department eq 'IT' or Department eq 'Finance'"

// Null check
"ApprovalDate ne null and RejectionReason eq null"

// Complex condition
"(Department eq 'IT' or Department eq 'Finance') and Amount gt 10000"
```

---

## 5. Project Structure

```
src/
├── formEngine/                          # Form engine core
│   ├── core/
│   │   ├── types.ts                     # All type definitions
│   │   ├── ODataConditionEngine.ts      # OData condition engine
│   │   ├── ValidationEngine.ts          # Validation engine
│   │   ├── FormStateManager.ts          # Form state management
│   │   └── StepEngine.ts                # Step engine
│   │
│   ├── data/
│   │   └── SharePointDataSource.ts       # SharePoint data source wrapper
│   │
│   ├── fields/                          # Built-in field types
│   │   ├── BaseField.tsx                # Field base class
│   │   ├── TextField.tsx
│   │   ├── MultilineField.tsx
│   │   ├── NumberField.tsx
│   │   ├── DateTimeField.tsx
│   │   ├── DropdownField.tsx
│   │   ├── MultiSelectField.tsx
│   │   ├── LookupField.tsx
│   │   ├── PersonField.tsx
│   │   └── BooleanField.tsx
│   │
│   ├── components/
│   │   ├── FormRenderer.tsx             # Main form renderer
│   │   ├── StepRenderer.tsx             # Step renderer
│   │   ├── FieldContainer.tsx           # Field container
│   │   └── FormStepper.tsx              # Step navigation
│   │
│   └── utils/
│       ├── odata/
│       │   ├── ODataLexer.ts            # Lexer
│       │   ├── ODataParser.ts           # Parser
│       │   └── ODataEvaluator.ts        # Evaluator
│       └── fieldValidator.ts            # Field validator
│
├── designer/                            # Visual designer
│   ├── components/
│   │   ├── FormDesigner.tsx             # Main designer component
│   │   ├── DesignerCanvas.tsx           # Canvas
│   │   ├── FieldPalette.tsx             # Field palette
│   │   ├── PropertyPanel.tsx            # Property panel
│   │   ├── StepEditor.tsx               # Step editor
│   │   ├── ConditionBuilder.tsx         # OData condition builder
│   │   └── PreviewPane.tsx              # Preview
│   │
│   └── controls/
│       ├── FieldPalette.tsx             # Field palette
│       └── FieldLayout.tsx              # Field layout
│
├── webparts/
│   └── dynamicForm/
│       ├── DynamicFormWebPart.ts
│       ├── components/
│       │   └── DynamicFormViewer.tsx    # Runtime form viewer
│       │
│       └── properties/
│           ├── PropertyPaneDesigner.tsx # Property panel designer
│           └── FormTemplateSelector.tsx # Template selector
│
├── templates/
│   └── formTemplates.ts                 # Predefined templates
│
└── common/
    ├── constants/
    ├── utilities/
    └── hooks/
        ├── useFormState.ts
        ├── useCondition.ts
        └── useDataSource.ts
```

---

## 6. Core Module Design

### 6.1 OData Condition Engine

```typescript
class ODataConditionEngine {
  /**
   * Evaluate an OData filter expression
   * @param expression - OData filter expression
   * @param context - current form data
   * @returns boolean
   */
  evaluate(expression: string, context: Record<string, any>): boolean;

  /**
   * Parse the expression into an AST
   */
  private parse(expression: string): ASTNode;

  /**
   * Supported operators
   */
  private readonly operators = {
    'eq': (a, b) => a === b,
    'ne': (a, b) => a !== b,
    'gt': (a, b) => a > b,
    'ge': (a, b) => a >= b,
    'lt': (a, b) => a < b,
    'le': (a, b) => a <= b,
    'and': (a, b) => a && b,
    'or': (a, b) => a || b,
    'not': (a) => !a,
    'contains': (field, value) => field?.includes(value) || false,
    'startswith': (field, prefix) => field?.startsWith(prefix) || false
  };
}
```

### 6.2 Form State Manager

```typescript
class FormStateManager {
  private state: Record<string, any>;
  private listeners: Map<string, Function[]>;

  setValue(fieldId: string, value: any): void;
  getValue(fieldId: string): any;
  subscribe(fieldId: string, callback: Function): void;
  triggerActions(actions: FieldAction[]): void;
  getAllState(): Record<string, any>;
}
```

### 6.3 SharePoint Data Source

```typescript
class SharePointDataSource {
  constructor(private context: WebPartContext) {}

  /**
   * Get all Lists (for the property pane dropdown)
   */
  getLists(): Promise<any[]>;

  /**
   * Get all fields of a List
   */
  getListFields(listName: string): Promise<SPField[]>;

  /**
   * Get a single item
   */
  getItem(listName: string, itemId: number): Promise<any>;

  /**
   * Create a new item
   */
  createItem(listName: string, item: any): Promise<any>;

  /**
   * Update an item
   */
  updateItem(listName: string, itemId: number, item: any): Promise<any>;

  /**
   * Get selectable values for a lookup field
   */
  getLookupChoices(lookupList: string, lookupField: string): Promise<any[]>;
}
```

**Behavior notes:**
- On Web Part initialization, all SharePoint lists in the current site are loaded automatically
- Displayed as a dropdown in the property pane, letting the user pick the target list
- List data loads asynchronously after Web Part initialization completes and refreshes the property pane

---

## 7. Complete Form Examples

### 7.1 Purchase Request Form

```json
{
  "id": "purchase-request",
  "name": "Purchase Request Form",
  "mode": "new",
  "listName": "PurchaseRequests",
  "submitButtonLabel": "Submit Request",
  "steps": [
    {
      "id": "step1",
      "title": "Basic Information",
      "fields": [
        {
          "id": "f1",
          "type": "text",
          "label": "Title",
          "fieldName": "Title",
          "required": "true"
        },
        {
          "id": "f2",
          "type": "dropdown",
          "label": "Category",
          "fieldName": "Category",
          "required": "true",
          "config": {
            "choices": ["Equipment", "Software", "Service", "Other"]
          },
          "onChange": [
            {
              "type": "show",
              "target": "f2_detail",
              "condition": "Category eq 'Other'"
            }
          ]
        },
        {
          "id": "f2_detail",
          "type": "text",
          "label": "Other Category Description",
          "fieldName": "OtherCategory",
          "visible": "Category eq 'Other'"
        },
        {
          "id": "f3",
          "type": "number",
          "label": "Amount",
          "fieldName": "Amount",
          "required": "true",
          "config": {
            "min": 0,
            "decimals": 2
          },
          "validation": [
            {
              "type": "min",
              "value": 0,
              "message": "Amount must be greater than 0"
            }
          ]
        },
        {
          "id": "f4",
          "type": "dropdown",
          "label": "Approval Level",
          "fieldName": "ApprovalLevel",
          "visible": "Amount ge 10000",
          "required": "Amount ge 10000",
          "config": {
            "choices": ["Department Manager", "Director", "VP"]
          }
        },
        {
          "id": "f5",
          "type": "person",
          "label": "VP Approver",
          "fieldName": "VPApprover",
          "visible": "Amount ge 50000 and ApprovalLevel eq 'VP'",
          "required": "Amount ge 50000 and ApprovalLevel eq 'VP'"
        },
        {
          "id": "f6",
          "type": "multiline",
          "label": "Description",
          "fieldName": "Description",
          "required": "Category eq 'Service'"
        },
        {
          "id": "f7",
          "type": "datetime",
          "label": "Expected Date",
          "fieldName": "ExpectedDate",
          "config": {
            "displayFormat": "dateOnly"
          }
        },
        {
          "id": "f8",
          "type": "boolean",
          "label": "Urgent",
          "fieldName": "IsUrgent"
        }
      ]
    }
  ]
}
```

### 7.2 Employee Information Form

```json
{
  "id": "employee-form",
  "name": "Employee Information Form",
  "mode": "new",
  "listName": "Employees",
  "steps": [
    {
      "id": "step1",
      "title": "Basic Information",
      "fields": [
        {
          "id": "f1",
          "type": "text",
          "label": "Name",
          "fieldName": "Title",
          "required": "true"
        },
        {
          "id": "f2",
          "type": "dropdown",
          "label": "Department",
          "fieldName": "Department",
          "required": "true",
          "config": {
            "choices": ["Engineering", "HR", "Finance", "Marketing"]
          }
        },
        {
          "id": "f3",
          "type": "dropdown",
          "label": "Tech Role",
          "fieldName": "TechRole",
          "visible": "Department eq 'Engineering'",
          "config": {
            "choices": ["Frontend Engineer", "Backend Engineer", "QA Engineer", "DevOps"]
          }
        },
        {
          "id": "f4",
          "type": "person",
          "label": "Manager",
          "fieldName": "Manager",
          "required": "true"
        },
        {
          "id": "f5",
          "type": "datetime",
          "label": "Start Date",
          "fieldName": "StartDate",
          "required": "true",
          "config": {
            "displayFormat": "dateOnly"
          }
        }
      ]
    }
  ]
}
```

---

## 8. Designer UI Design

### 8.1 Condition Builder

```
┌─────────────────────────────────────────────────────────────┐
│  Field display condition                       [+ Add condition]  │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌────────────────────────────────────────────────────────┐ │
│  │  [Field ▼]  [Operator ▼]  [Value____________________]  [Delete]│ │
│  │  Department  eq         'IT'                          │ │
│  │                                                        │ │
│  │  [and ▼] [Field ▼] [Operator ▼] [Value______________] [Delete]│ │
│  │          Status    eq         'Pending'               │ │
│  └────────────────────────────────────────────────────────┘ │
│                                                              │
│  OData expression preview:                                            │
│  ┌────────────────────────────────────────────────────────┐ │
│  │ Department eq 'IT' and Status eq 'Pending'             │ │
│  └────────────────────────────────────────────────────────┘ │
│                                                              │
│  [Test condition]                                                  │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

### 8.2 Designer Main Interface

```
┌─────────────────────────────────────────────────────────────────┐
│  Form Designer                                              Save │
├──────────────┬────────────────────────────┬─────────────────────┤
│              │                             │                     │
│  Field Palette │      Design Canvas         │   Property Panel    │
│              │                             │                     │
│ ┌──────────┐ │  ┌───────────────────────┐ │ ┌─────────────────┐ │
│ │ Basic Fields│ │  │ Step 1: Basic Info    │ │ │ Field Properties │ │
│ ├──────────┤ │  │                       │ │ ├─────────────────┤ │
│ │[Drag] Text│ │  │ ┌───────────────────┐ │ │ │ Title           │ │
│ │[Drag] Multi│ │  │ │ Title:            │ │ │ │ [Title________] │ │
│ │  line     │ │  │ │ [________________] │ │ │ │                 │ │
│ │[Drag] Number│ │  │ └───────────────────┘ │ │ │ Field Name      │ │
│ │[Drag] Date│ │  │                       │ │ │ [Title________] │ │
│ ├──────────┤ │  │ ┌───────────────────┐ │ │ │                 │ │
│ │ Choice Fields│ │  │ │ Department: *     │ │ │ │ Display Condition│ │
│ ├──────────┤ │  │ │ [Department ▼]    │ │ │ │ [+ Add condition]│ │
│ │[Drag]Dropdn│ │  │ └───────────────────┘ │ │ │                 │ │
│ │[Drag]Multi-│ │  │                       │ │ │ Required Condition│ │
│ │  select   │ │  │ [+ Add field]          │ │ │ [+ Add condition]│ │
│ │[Drag]Person│ │  │                       │ │ │                 │ │
│ ├──────────┤ │  └───────────────────────┘ │ │ Validation Rules │ │
│ │ Advanced  │ │                             │ │ [+ Add rule]     │ │
│ │ Fields    │ │  ┌─ Step Navigation ─────┐ │ │                 │ │
│ ├──────────┤ │  │ [Step 1] [Step 2] [+]  │ │ │ Field Dependency │ │
│ │[Drag]Lookup│ │  └───────────────────────┘ │ │ [+ Add action]   │ │
│ │[Drag]Yes/No│ │                             │ └─────────────────┘ │
│ └──────────┘ │                             │                     │
├──────────────┴─────────────────────────────┴─────────────────────┤
│  [Preview form]                                                    │
└─────────────────────────────────────────────────────────────────┘
```

---

## 9. Implementation Plan

### Phase 1: Core Engine
- [x] Create type definition file
- [x] Implement OData condition engine
- [x] Implement form state manager
- [x] Implement validation engine

### Phase 2: Data Layer & Fields
- [x] Implement SharePoint data source
- [x] Implement automatic SharePoint list retrieval
- [x] Implement built-in field type components
- [x] Implement field container component

### Phase 3: Form Rendering
- [x] Implement form renderer
- [x] Implement step renderer
- [x] Implement step navigation component

### Phase 4: Designer
- [x] Implement main designer component
- [x] Implement field palette
- [x] Implement property panel
- [x] Implement OData condition builder
- [x] Implement preview functionality

### Phase 5: Web Part Integration
- [x] Implement property pane integration
- [x] Implement SharePoint list dropdown selection
- [x] Implement form viewer
- [x] Support New/Edit/View modes

### Phase 6: Templates & Export
- [x] Create predefined templates
- [x] Implement import/export functionality

---

## 10. Tech Stack

- **Framework**: SPFx 1.21 + React 17
- **UI Library**: Fluent UI 8
- **Language**: TypeScript 5.3
- **State Management**: Custom (FormStateManager)
- **Build**: Gulp

---

