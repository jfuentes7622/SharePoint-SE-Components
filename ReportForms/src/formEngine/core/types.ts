/**
 * SPFx Dynamic Form Engine - Core Type Definitions
 */

// ============================================================================
// Basic Types
// ============================================================================

export type FormMode = 'new' | 'edit' | 'view';

export type FieldType =
  | 'text'
  | 'multiline'
  | 'richtext'
  | 'customimage'
  | 'divider'
  | 'number'
  | 'datetime'
  | 'dropdown'
  | 'multiselect'
  | 'lookup'
  | 'person'
  | 'boolean'
  | 'image'
  | 'url'
  | 'taxonomy'
  | 'attachment'
  | 'listcontrol'
  | 'newline';

export type FilterExpression = string;

// ============================================================================
// Field Value Union
// ============================================================================

/** Represents the possible value types across all form fields. */
export type FieldValue =
  | string
  | number
  | boolean
  | null
  | undefined
  | Date
  | Array<string>
  | Array<number>
  | Record<string, any>
  | Array<Record<string, any>>;

// ============================================================================
// Form Schema
// ============================================================================

export interface FormSchema {
  id: string;
  name: string;
  description?: string;
  nameAlignment?: 'left' | 'center' | 'right';
  descriptionAlignment?: 'left' | 'center' | 'right';
  logoUrl?: string;
  logoAltText?: string;
  showTitle?: boolean; // default true
  mode: FormMode;
  listName: string;
  itemId?: number;
  steps: FormStep[];
  showFieldDescription?: boolean;
  submitButtonLabel?: string;
  showCancelButton?: boolean;
  cancelButtonLabel?: string;
  cancelRedirectUrl?: string;
  onSubmitMessage?: string;
  permissionDeniedMessage?: string;
  submitRedirectUrl?: string;
  submitRedirectDelayMs?: number;
  theme?: FormTheme;
  advancedValidation?: AdvancedValidationConfig;
}

export interface AdvancedValidationConfig {
  enabled?: boolean;
  rules?: AdvancedValidationRule[];
}

export interface AdvancedValidationRule {
  id?: string;
  expression: string;
  message: string;
  targetField?: string; // field id or SharePoint internal name for inline error mapping
}

export interface FormTheme {
  layout?: 'stack' | 'grid';
  columns?: number;
  labelPosition?: 'top' | 'left';
  // Font settings for form title/name
  fontSize?: number; // in pixels
  fontFamily?: string;
  fontWeight?: 'normal' | 'bold' | '500' | '600' | '700';
  color?: string;
  // Font settings for form description
  descriptionFontSize?: number; // in pixels
  descriptionFontFamily?: string;
  descriptionFontWeight?: 'normal' | 'bold' | '500' | '600' | '700';
  descriptionColor?: string;
  // Background and border settings for container
  backgroundColor?: string;
  borderColor?: string;
  borderStyle?: 'square' | 'rounded'; // square (no radius) or rounded (8px radius)
  borderRadius?: number; // in pixels, used when borderStyle is rounded
  borderWidth?: number; // in pixels
}

// ============================================================================
// Step & Field
// ============================================================================

export interface FormStep {
  id: string;
  title: string;
  description?: string;
  titleAlignment?: 'left' | 'center' | 'right';
  descriptionAlignment?: 'left' | 'center' | 'right';
  showTitle?: boolean; // default true
  fields: (FormField | null)[]; // allow null placeholders to preserve correct grid layout positions
  visible?: boolean; // whether the step is visible (e.g. for welcome-page scenarios)
  theme?: FormTheme; // step-level layout config, overrides the global config
}

export interface FormField {
  id: string;
  type: FieldType;
  label: string;
  fieldName: string;
  description?: string;
  visible?: boolean | FilterExpression;
  required?: boolean | FilterExpression;
  requiredMessage?: string;
  readOnly?: boolean | FilterExpression;
  disabled?: boolean;
  onChange?: FieldAction[];
  validation?: ValidationRule[];
  config?: FieldConfig;
  columnSpan?: number; // number of columns the field spans, defaults to 1 (full row or based on grid column count)
  defaultValue?: FieldValue;
  startNewRow?: boolean; // whether to start a new row in the grid layout
  labelPosition?: 'top' | 'bottom' | 'left' | 'right'; // override global labelPosition from theme
  // Font settings for field label
  labelFontSize?: number; // in pixels
  labelFontFamily?: string;
  labelFontWeight?: 'normal' | 'bold' | '500' | '600' | '700';
  labelColor?: string;
  // Font settings for field input/value
  inputFontSize?: number; // in pixels
  inputFontFamily?: string;
  inputFontWeight?: 'normal' | 'bold' | '500' | '600' | '700';
  inputColor?: string;
  // Background and border settings for field
  fieldBackgroundColor?: string;
  fieldBorderColor?: string;
  fieldBorderStyle?: 'square' | 'rounded'; // square (no radius) or rounded (8px radius)
  fieldBorderRadius?: number; // in pixels, used when fieldBorderStyle is rounded
  fieldBorderWidth?: number; // in pixels
}

export interface FieldConfig {
  maxLength?: number;
  placeholder?: string;
  helpText?: string;
  isRichText?: boolean;
  min?: number;
  max?: number;
  decimals?: number;
  displayFormat?: 'dateOnly' | 'dateTime' | 'timeOnly';
  timeZone?: 'UTC' | 'local'; // applies to dateTime/timeOnly display formats; defaults to 'UTC' (Zulu)
  choices?: string[];
  choiceDisplay?: 'dropdown' | 'radio' | 'checkboxes';
  exclusiveChoiceValue?: string;
  allowFillIn?: boolean;
  lookupList?: string;
  lookupField?: string;
  allowMultiple?: boolean;
  termSetId?: string;
  listName?: string;
  listId?: string;
  itemId?: number;
  accepts?: string;
  allowAttachmentDelete?: boolean;
  booleanText?: string;
  listControlSourceId?: string;
  listControlSourceName?: string;
  listControlSourceListName?: string;
  listControlFilterSourceField?: string;
  listControlFilterTargetField?: string;
  listControlFilterOperator?: string;
  imageUrl?: string;
  imageAltText?: string;
  imageWidth?: number;
  imageHeight?: number;
  imageFit?: 'contain' | 'cover';
  imageAlignment?: 'left' | 'center' | 'right';
  dividerColor?: string;
  dividerThickness?: number;
  dividerStyle?: 'solid' | 'dashed' | 'dotted';
  dividerSpacing?: number;
  dividerOrientation?: 'horizontal' | 'vertical';
  dividerLength?: number;
}

export type FieldAction =
  | { type: 'show'; target: string; condition?: FilterExpression }
  | { type: 'hide'; target: string }
  | { type: 'set-value'; target: string; value: FieldValue }
  | { type: 'clear'; target: string }
  | { type: 'enable'; target: string }
  | { type: 'disable'; target: string };

export interface ValidationRule {
  type: 'required' | 'minLength' | 'maxLength' | 'min' | 'max' | 'pattern' | 'custom';
  value?: FieldValue;
  message: string;
  applyWhen?: FilterExpression;
}

export interface ValidationResult {
  valid: boolean;
  errors: Record<string, string[]>;
}

// ============================================================================
// Form State
// ============================================================================

export interface FieldState {
  value: FieldValue;
  touched: boolean;
  dirty: boolean;
  visible: boolean;
  required: boolean;
  readOnly: boolean;
  disabled: boolean;
  valid: boolean;
  errors: string[];
}

export interface FormState {
  fields: Record<string, FieldState>;
  currentStep: number;
  isSubmitting: boolean;
  isValid: boolean;
}

// ============================================================================
// OData AST
// ============================================================================

export type ASTNode =
  | BinaryOpNode
  | UnaryOpNode
  | FunctionCallNode
  | FieldNode
  | ValueNode
  | GroupNode;

export interface BinaryOpNode {
  type: 'BinaryOp';
  operator: 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le' | 'and' | 'or';
  left: ASTNode;
  right: ASTNode;
}

export interface UnaryOpNode {
  type: 'UnaryOp';
  operator: 'not';
  operand: ASTNode;
}

export interface FunctionCallNode {
  type: 'FunctionCall';
  name: 'contains' | 'startswith';
  args: ASTNode[];
}

export interface FieldNode {
  type: 'Field';
  name: string;
}

export interface ValueNode {
  type: 'Value';
  value: string | number | boolean | null;
}

export interface GroupNode {
  type: 'Group';
  expression: ASTNode;
}

// ============================================================================
// SharePoint Types
// ============================================================================

export enum SPFieldType {
  Text = 'Text',
  Note = 'Note',
  Number = 'Number',
  Integer = 'Integer',
  DateTime = 'DateTime',
  Choice = 'Choice',
  MultiChoice = 'MultiChoice',
  Lookup = 'Lookup',
  User = 'User',
  UserMulti = 'UserMulti',
  Boolean = 'Boolean',
  URL = 'URL',
  Hyperlink = 'Hyperlink',
  Image = 'Image',
  Calculated = 'Calculated',
  Taxonomy = 'Taxonomy',
  TaxonomyMulti = 'TaxonomyMulti',
  Attachments = 'Attachments',
}

export interface SPFieldInfo {
  id: string;
  internalName: string;
  title: string;
  description?: string;
  type: SPFieldType;
  richText?: boolean;
  required: boolean;
  readOnly: boolean;
  choices?: string[];
  lookupList?: string;
  lookupField?: string;
  allowMultipleValues?: boolean;
  maxLength?: number;
  textField?: string;
  termSetId?: string;
  displayFormat?: 'dateOnly' | 'dateTime';
}

// ============================================================================
// Context Types
// ============================================================================

export interface DataContextValue {
  getListFields: (listName: string) => Promise<SPFieldInfo[]>;
  getItem: (listName: string, itemId: number) => Promise<Record<string, any>>;
  createItem: (listName: string, item: Record<string, any>) => Promise<Record<string, any>>;
  updateItem: (listName: string, itemId: number, item: Record<string, any>) => Promise<Record<string, any>>;
  getLookupChoices: (lookupList: string, lookupField: string) => Promise<Array<{ Id: number; Title?: string; [key: string]: any }>>;
}
