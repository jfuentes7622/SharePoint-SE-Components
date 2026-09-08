import * as React from 'react';
import {
  DefaultButton,
  MessageBar,
  MessageBarType,
  PrimaryButton,
} from 'office-ui-fabric-react';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import styles from './SharePointDynamicForm.module.scss';
import { FormDesigner } from './FormDesigner';
import { RichTextEditor } from './RichTextEditor';
import { ListControlHost } from './ListControlHost';
import { FormField, FormMode, FormSchema, FieldValue, FieldConfig, AdvancedValidationRule } from '../../../formEngine/core/types';
import * as strings from 'SharePointDynamicFormWebPartStrings';

type SPFxContext = any;

export interface SharePointDynamicFormContainerProps {
  isInDesignerMode: boolean;
  isDesignerAvailable: boolean;
  onToggleDesignerMode: () => void;
  formSchemaJson: string;
  listName: string;
  mode: FormMode;
  dynamicItemId?: number;
  linkedFieldTarget?: string;
  permissionBaseLookupField?: string;
  permissionScope?: 'list' | 'item';
  permissionScopeItemId?: string;
  linkedFieldValue?: any;
  useDynamicValueAsFilter?: boolean;
  filterJson?: string;
  useItemId: boolean;
  itemId: number;
  itemIdQueryParam?: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  context: SPFxContext;
  onSaveSchema: (schema: FormSchema) => void;
  labelPosition?: 'top' | 'left';
  isPageEditMode: boolean;
  showFieldDescription?: boolean;
  containerWidth?: number;
  isFullWidth?: boolean;
  submitButtonLabel?: string;
  addSubmitButtonLabel?: string;
  editSubmitButtonLabel?: string;
  showCancelButton?: boolean;
  cancelButtonLabel?: string;
  cancelRedirectUrl?: string;
  submitRedirectUrl?: string;
  onSubmitMessage?: string;
  buttonTextColor?: string;
  buttonBackgroundColor?: string;
  buttonBorderColor?: string;
  buttonBorderWidth?: number | string;
  buttonFontFamily?: string;
  buttonFontSize?: number | string;
  buttonFontStyle?: string;
  buttonFontBold?: boolean;
  buttonCornerStyle?: 'square' | 'rounded';
  buttonCornerRadius?: number;
  defaultValuesJson?: string;
  advancedValidationEnabled?: boolean;
  advancedValidationJson?: string;
  dynamicDiagnostics?: string[];
  dynamicItemReference?: string;
  dynamicModeReference?: string;
  dynamicPreferredSourceInstanceId?: string;
  enableDynamicDiagnostics?: boolean;
}

interface ILookupOptionItem {
  Id: number;
  Title?: string;
  Email?: string;
}

interface IAttachmentItem {
  fileName: string;
  serverRelativeUrl?: string;
  length?: number;
  timeCreated?: string;
  timeLastModified?: string;
}

interface ISelectOption {
  key: string;
  text: string;
}

interface IFieldErrorMap {
  [fieldId: string]: string;
}

interface IValueMap {
  [fieldId: string]: FieldValue;
}

interface ILookupMap {
  [fieldId: string]: ILookupOptionItem[];
}

interface SharePointDynamicFormContainerState {
  schema: FormSchema | null;
  values: IValueMap;
  loadedItem: any;
  lookupOptions: ILookupMap;
  resolvedItemId: number;
  isViewEditing: boolean;
  loading: boolean;
  error: string | null;
  fieldErrors: IFieldErrorMap;
  isSubmitting: boolean;
  submitError: string | null;
  submitSuccess: string | null;
  canAddRecords: boolean;
  canEditRecords: boolean;
  permissionStatusMessage: string | null;
}

interface ILookupPermissionFieldMetadata {
  internalName: string;
  typeAsString: string;
  lookupList: string;
}

type FilterConditionOperator = 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le' | 'contains' | 'notcontains' | 'startswith' | 'endswith';
type FilterConditionLogical = 'and' | 'or';

interface IFilterCondition {
  field: string;
  type?: string;
  value: any;
  valueType?: string;
  operator?: FilterConditionOperator;
  logical?: FilterConditionLogical;
}

var LIST_CONTROL_REFRESH_EVENT = 'spse:listcontrol-refresh';

function buildBlankSchema(listName: string, mode: FormMode): FormSchema {
  return {
    id: 'blank-form',
    name: 'Form',
    mode: mode,
    listName: listName,
    theme: {
      layout: 'stack',
      columns: 1,
    },
    steps: [
      {
        id: 'step1',
        title: strings.DesignerStepDefaultTitle.replace('{0}', '1'),
        fields: [],
      },
    ],
  };
}

function isGuid(value: string): boolean {
  // Match GUIDs with or without curly braces: {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx} or xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
  return /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i.test(value);
}

function toPositiveInt(value: any): number {
  var parsed = parseInt(String(value), 10);
  return !isNaN(parsed) && parsed > 0 ? parsed : 0;
}

// Fallback used whenever a level (web part/form/container/field) does not customize its own font,
// so styling never silently cascades in from a different (unrelated) level's explicit setting.
var DEFAULT_THEME_FONT_FAMILY = "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif";

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

function normalizeFilterOperator(value: any): FilterConditionOperator {
  var op = String(value || '').trim().toLowerCase();
  if (op === 'ne' || op === 'gt' || op === 'ge' || op === 'lt' || op === 'le' || op === 'contains' || op === 'notcontains' || op === 'startswith' || op === 'endswith') {
    return op as FilterConditionOperator;
  }

  return 'eq';
}

function normalizeFilterLogical(value: any): FilterConditionLogical {
  return String(value || '').trim().toLowerCase() === 'or' ? 'or' : 'and';
}

function getFilterTypeFromSharePointType(typeAsString: string): string {
  var normalizedType = String(typeAsString || '').toLowerCase();
  if (normalizedType === 'lookup' || normalizedType === 'lookupmulti' || normalizedType === 'user' || normalizedType === 'usermulti') {
    return 'lookup';
  }
  if (normalizedType === 'number' || normalizedType === 'currency' || normalizedType === 'integer' || normalizedType === 'counter') {
    return 'number';
  }
  if (normalizedType === 'boolean') {
    return 'boolean';
  }
  if (normalizedType === 'datetime') {
    return 'datetime';
  }
  return 'text';
}

function buildODataFilterClause(fieldName: string, fieldType: string, value: any, operator?: FilterConditionOperator): string {
  if (value === undefined || value === null || value === '') { return ''; }
  var op = normalizeFilterOperator(operator);
  if (Array.isArray(value)) {
    var arrayClauses = value.map(function(arrayValue: any) {
      return buildODataFilterClause(fieldName, fieldType, arrayValue, op === 'ne' ? 'eq' : (op === 'notcontains' ? 'contains' : op));
    }).filter(function(arrayClause: string) { return !!arrayClause; });
    if (arrayClauses.length === 0) { return ''; }
    if (op === 'ne' || op === 'notcontains') {
      return arrayClauses.map(function(arrayClause: string) { return 'not (' + arrayClause + ')'; }).join(' and ');
    }
    return arrayClauses.map(function(arrayClause: string) { return '(' + arrayClause + ')'; }).join(' or ');
  }
  var clause = '';
  switch (fieldType) {
    case 'lookup':
    case 'person': {
      if (op === 'contains' || op === 'notcontains' || op === 'startswith' || op === 'endswith') {
        console.warn('[ReportFormsWebPart] Unsupported lookup operator for field ' + fieldName + ': ' + op + '. Falling back to eq.');
        op = 'eq';
      }
      var numVal = parseInt(String(value), 10);
      if (isNaN(numVal)) { 
        console.warn('[ReportFormsWebPart] Invalid lookup value for field ' + fieldName + ': ' + value);
        return ''; 
      }
      clause = fieldName + 'Id ' + op + ' ' + numVal;
      console.log('[ReportFormsWebPart] Built lookup filter: ' + clause + ' (field=' + fieldName + ', value=' + value + ')');
      return clause;
    }
    case 'number': {
      if (op === 'contains' || op === 'notcontains' || op === 'startswith' || op === 'endswith') {
        console.warn('[ReportFormsWebPart] Unsupported number operator for field ' + fieldName + ': ' + op + '. Falling back to eq.');
        op = 'eq';
      }
      var n = Number(value);
      if (isNaN(n)) { 
        console.warn('[ReportFormsWebPart] Invalid number value for field ' + fieldName + ': ' + value);
        return ''; 
      }
      clause = fieldName + ' ' + op + ' ' + n;
      console.log('[ReportFormsWebPart] Built number filter: ' + clause);
      return clause;
    }
    case 'boolean': {
      var bv = value === true || value === 'true' || value === '1';
      if (op !== 'eq' && op !== 'ne') {
        console.warn('[ReportFormsWebPart] Unsupported boolean operator for field ' + fieldName + ': ' + op + '. Falling back to eq.');
        op = 'eq';
      }
      clause = fieldName + ' ' + op + ' ' + (bv ? 'true' : 'false');
      console.log('[ReportFormsWebPart] Built boolean filter: ' + clause);
      return clause;
    }
    case 'datetime': {
      if (op === 'contains' || op === 'notcontains' || op === 'startswith' || op === 'endswith') {
        op = 'eq';
      }
      var parsedDate = new Date(String(value));
      if (isNaN(parsedDate.getTime())) {
        return '';
      }
      return fieldName + ' ' + op + " datetime'" + parsedDate.toISOString() + "'";
    }
    default: // choice, text, multiline, etc.
      if (op === 'contains' || op === 'notcontains' || op === 'startswith' || op === 'endswith') {
        var textFunction = op === 'notcontains' ? 'contains' : op;
        clause = textFunction + '(' + fieldName + ", '" + escapeODataText(String(value)) + "')";
        if (op === 'notcontains') {
          clause = 'not (' + clause + ')';
        }
      } else {
        clause = fieldName + ' ' + op + " '" + escapeODataText(String(value)) + "'";
      }
      console.log('[ReportFormsWebPart] Built text/choice filter: ' + clause + ' (type=' + fieldType + ')');
      return clause;
  }
}

function formatExpressionLocalDate(value: Date): string {
  var month = String(value.getMonth() + 1);
  var day = String(value.getDate());
  return String(value.getFullYear()) + '-' + (month.length < 2 ? '0' + month : month) + '-' + (day.length < 2 ? '0' + day : day);
}

function padTwoDigits(value: number): string {
  var text = String(value);
  return text.length < 2 ? '0' + text : text;
}

function applyDateFormatPattern(pattern: string, tokenValues: { [token: string]: string }): string {
  return pattern.replace(/YYYY|YY|MMMM|MMM|MM|M|DD|D|dddd|ddd|HH|H|hh|h|mm|ss|tt/g, function(token: string): string {
    return tokenValues[token] !== undefined ? tokenValues[token] : token;
  });
}

function applyTextCase(text: string, textCase: string): string {
  if (textCase === 'upper') { return text.toUpperCase(); }
  if (textCase === 'lower') { return text.toLowerCase(); }
  return text;
}

function resolveSafeExpression(expression: string, context: any): any {
  var normalized = String(expression || '').trim().toLowerCase();
  var pageContext = context && context.pageContext;
  var user = pageContext && pageContext.user;
  var legacyContext = pageContext && pageContext.legacyPageContext;
  if (normalized === 'today') { return formatExpressionLocalDate(new Date()); }
  if (normalized === 'now') { return new Date().toISOString(); }
  var dateMatch = /^date\(\s*([+-]?\d+)\s*\)$/.exec(normalized);
  if (dateMatch) {
    var dateValue = new Date();
    dateValue.setDate(dateValue.getDate() + parseInt(dateMatch[1], 10));
    return formatExpressionLocalDate(dateValue);
  }
  if (normalized === 'me.email') { return String(user && user.email || ''); }
  if (normalized === 'me.login') { return String(user && user.loginName || ''); }
  if (normalized === 'me.id') { return legacyContext && legacyContext.userId || ''; }
  if (normalized === 'me') { return String(user && (user.displayName || user.loginName) || ''); }
  return expression;
}

function parseFilterJsonToExpression(filterJson: string, dynamicValue: any, fields: FormField[], context: any): string {
  var expression = '';
  try {
    var conditions = JSON.parse(filterJson) as IFilterCondition[];
    if (!Array.isArray(conditions)) { 
      console.warn('[ReportFormsWebPart] filterJson is not an array: ' + filterJson);
      return expression; 
    }
    console.log('[ReportFormsWebPart] Parsing filterJson with ' + conditions.length + ' conditions');
    
    for (var i = 0; i < conditions.length; i++) {
      var cond = conditions[i];
      if (!cond || !cond.field) { 
        console.warn('[ReportFormsWebPart] Skipping invalid condition at index ' + i + ': missing field');
        continue; 
      }
      
      var val = cond.value === 'dynamic'
        ? dynamicValue
        : (String(cond.valueType || '').toLowerCase() === 'expression' ? resolveSafeExpression(String(cond.value || ''), context) : cond.value);
      var op = normalizeFilterOperator(cond.operator);
      var logical = normalizeFilterLogical(cond.logical);
      console.log('[ReportFormsWebPart] Condition ' + i + ': field=' + cond.field + ', type=' + (cond.type || '(auto)') + ', value=' + val);
      
      var condType: string = cond.type || '';
      // Resolve type from schema fields if not explicitly provided
      if (!condType) {
        for (var j = 0; j < fields.length; j++) {
          if (fields[j].fieldName && fields[j].fieldName.toLowerCase() === String(cond.field).toLowerCase()) {
            condType = fields[j].type;
            console.log('[ReportFormsWebPart] Resolved type from schema: ' + condType);
            break;
          }
        }
      }
      
      var clause = buildODataFilterClause(String(cond.field), condType || 'text', val, op);
      if (clause) {
        if (!expression) {
          expression = clause;
        } else {
          expression = '(' + expression + ') ' + logical + ' (' + clause + ')';
        }
        console.log('[ReportFormsWebPart] Added filter clause: ' + clause);
      } else {
        console.warn('[ReportFormsWebPart] Failed to build clause for condition ' + i);
      }
    }
  } catch (parseError) {
    console.error('[ReportFormsWebPart] Error parsing filterJson: ', parseError);
    if (parseError && parseError.message) {
      console.error('[ReportFormsWebPart] Parse error message: ' + parseError.message);
    }
  }
  return expression;
}

function getAllFields(schema: FormSchema): FormField[] {
  var result: FormField[] = [];
  for (var i = 0; i < schema.steps.length; i += 1) {
    var fields = schema.steps[i].fields;
    for (var j = 0; j < fields.length; j += 1) {
      if (fields[j]) {
        result.push(fields[j] as FormField);
      }
    }
  }
  return result;
}

interface IAdvancedToken {
  type: 'identifier' | 'number' | 'string' | 'operator' | 'paren' | 'comma' | 'eof';
  value: string;
}

function isValueEmpty(value: any): boolean {
  return value === undefined || value === null || value === '' || (Array.isArray(value) && value.length === 0);
}

function toComparableValue(value: any): any {
  if (value instanceof Date) {
    return value.getTime();
  }

  if (typeof value === 'string') {
    var trimmed = value.trim();
    if (!trimmed) {
      return trimmed;
    }

    if (!isNaN(Number(trimmed))) {
      return Number(trimmed);
    }

    var parsedDate = new Date(trimmed);
    if (!isNaN(parsedDate.getTime())) {
      return parsedDate.getTime();
    }

    return trimmed.toLowerCase();
  }

  return value;
}

function compareAdvancedValues(left: any, right: any): number {
  var l = toComparableValue(left);
  var r = toComparableValue(right);

  if (l === r) {
    return 0;
  }

  if (l > r) {
    return 1;
  }

  if (l < r) {
    return -1;
  }

  return 0;
}

function tokenizeAdvancedExpression(expression: string): IAdvancedToken[] {
  var tokens: IAdvancedToken[] = [];
  var i = 0;

  while (i < expression.length) {
    var ch = expression.charAt(i);

    if (/\s/.test(ch)) {
      i += 1;
      continue;
    }

    var two = expression.substr(i, 2);
    if (two === '&&' || two === '||' || two === '==' || two === '!=' || two === '>=' || two === '<=') {
      tokens.push({ type: 'operator', value: two });
      i += 2;
      continue;
    }

    if (ch === '>' || ch === '<' || ch === '!') {
      tokens.push({ type: 'operator', value: ch });
      i += 1;
      continue;
    }

    if (ch === '(' || ch === ')') {
      tokens.push({ type: 'paren', value: ch });
      i += 1;
      continue;
    }

    if (ch === ',') {
      tokens.push({ type: 'comma', value: ch });
      i += 1;
      continue;
    }

    if (ch === '"' || ch === '\'') {
      var quote = ch;
      i += 1;
      var strVal = '';
      while (i < expression.length) {
        var s = expression.charAt(i);
        if (s === '\\') {
          if (i + 1 < expression.length) {
            strVal += expression.charAt(i + 1);
            i += 2;
            continue;
          }
          break;
        }
        if (s === quote) {
          break;
        }
        strVal += s;
        i += 1;
      }
      if (i >= expression.length || expression.charAt(i) !== quote) {
        throw new Error('Unterminated string literal in expression.');
      }
      i += 1;
      tokens.push({ type: 'string', value: strVal });
      continue;
    }

    if (/\d/.test(ch)) {
      var num = ch;
      i += 1;
      while (i < expression.length && /[\d.]/.test(expression.charAt(i))) {
        num += expression.charAt(i);
        i += 1;
      }
      tokens.push({ type: 'number', value: num });
      continue;
    }

    if (/[A-Za-z_]/.test(ch)) {
      var ident = ch;
      i += 1;
      while (i < expression.length && /[A-Za-z0-9_]/.test(expression.charAt(i))) {
        ident += expression.charAt(i);
        i += 1;
      }
      tokens.push({ type: 'identifier', value: ident });
      continue;
    }

    throw new Error('Unsupported token in expression near: ' + expression.substr(i, 10));
  }

  tokens.push({ type: 'eof', value: '' });
  return tokens;
}

function evaluateAdvancedExpression(expression: string, fields: FormField[], values: IValueMap): any {
  var tokens = tokenizeAdvancedExpression(expression);
  var index = 0;

  var byKey: { [key: string]: FormField } = {};
  for (var i = 0; i < fields.length; i += 1) {
    var f = fields[i];
    if (f.id) { byKey[f.id.toLowerCase()] = f; }
    if (f.fieldName) { byKey[f.fieldName.toLowerCase()] = f; }
    if (f.label) { byKey[f.label.toLowerCase()] = f; }
  }

  var peek = function(): IAdvancedToken {
    return tokens[index] || { type: 'eof', value: '' };
  };

  var consume = function(expectedType?: string, expectedValue?: string): IAdvancedToken {
    var current = peek();
    if (expectedType && current.type !== expectedType) {
      throw new Error('Unexpected token: expected ' + expectedType + ' but found ' + current.type + '.');
    }
    if (expectedValue && current.value !== expectedValue) {
      throw new Error('Unexpected token: expected "' + expectedValue + '" but found "' + current.value + '".');
    }
    index += 1;
    return current;
  };

  var resolveFieldValue = function(fieldRef: any): any {
    var normalized = String(fieldRef === undefined || fieldRef === null ? '' : fieldRef).trim().toLowerCase();
    if (!normalized) {
      return undefined;
    }

    var found = byKey[normalized];
    if (!found) {
      return undefined;
    }

    return values[found.id];
  };

  var parseExpression = function(): any {
    return parseOr();
  };

  var parseOr = function(): any {
    var left = parseAnd();
    while (peek().type === 'operator' && peek().value === '||') {
      consume('operator', '||');
      var right = parseAnd();
      left = !!left || !!right;
    }
    return left;
  };

  var parseAnd = function(): any {
    var left = parseComparison();
    while (peek().type === 'operator' && peek().value === '&&') {
      consume('operator', '&&');
      var right = parseComparison();
      left = !!left && !!right;
    }
    return left;
  };

  var parseComparison = function(): any {
    var left = parseUnary();
    if (peek().type === 'operator' && (peek().value === '==' || peek().value === '!=' || peek().value === '>' || peek().value === '>=' || peek().value === '<' || peek().value === '<=')) {
      var op = consume('operator').value;
      var right = parseUnary();
      var cmp = compareAdvancedValues(left, right);
      if (op === '==') { return cmp === 0; }
      if (op === '!=') { return cmp !== 0; }
      if (op === '>') { return cmp > 0; }
      if (op === '>=') { return cmp >= 0; }
      if (op === '<') { return cmp < 0; }
      if (op === '<=') { return cmp <= 0; }
    }
    return left;
  };

  var parseUnary = function(): any {
    if (peek().type === 'operator' && peek().value === '!') {
      consume('operator', '!');
      return !parseUnary();
    }
    return parsePrimary();
  };

  var evaluateFunction = function(name: string, args: any[]): any {
    var fn = name.toLowerCase();
    if (fn === 'field') {
      return args.length > 0 ? resolveFieldValue(args[0]) : undefined;
    }
    if (fn === 'today') {
      var t = new Date();
      t.setHours(0, 0, 0, 0);
      return t;
    }
    if (fn === 'now') {
      return new Date();
    }
    if (fn === 'daysfromtoday') {
      var d = new Date();
      d.setHours(0, 0, 0, 0);
      var delta = Number(args[0] || 0);
      d.setDate(d.getDate() + (isNaN(delta) ? 0 : delta));
      return d;
    }
    if (fn === 'isempty') {
      return isValueEmpty(args[0]);
    }
    if (fn === 'hasvalue') {
      return !isValueEmpty(args[0]);
    }
    if (fn === 'between') {
      if (args.length < 3) {
        return false;
      }
      var c1 = compareAdvancedValues(args[0], args[1]);
      var c2 = compareAdvancedValues(args[0], args[2]);
      return c1 >= 0 && c2 <= 0;
    }

    throw new Error('Unsupported function: ' + name);
  };

  var parsePrimary = function(): any {
    var token = peek();

    if (token.type === 'paren' && token.value === '(') {
      consume('paren', '(');
      var exprVal = parseExpression();
      consume('paren', ')');
      return exprVal;
    }

    if (token.type === 'number') {
      consume('number');
      return Number(token.value);
    }

    if (token.type === 'string') {
      consume('string');
      return token.value;
    }

    if (token.type === 'identifier') {
      consume('identifier');
      var identName = token.value;
      var lowerName = identName.toLowerCase();

      if (lowerName === 'true') {
        return true;
      }
      if (lowerName === 'false') {
        return false;
      }
      if (lowerName === 'null') {
        return null;
      }

      if (peek().type === 'paren' && peek().value === '(') {
        consume('paren', '(');
        var args: any[] = [];
        if (!(peek().type === 'paren' && peek().value === ')')) {
          while (true) {
            args.push(parseExpression());
            if (peek().type === 'comma') {
              consume('comma', ',');
              continue;
            }
            break;
          }
        }
        consume('paren', ')');
        return evaluateFunction(identName, args);
      }

      return resolveFieldValue(identName);
    }

    throw new Error('Unexpected token in expression: ' + token.type + ' ' + token.value);
  };

  var result = parseExpression();
  if (peek().type !== 'eof') {
    throw new Error('Unexpected trailing tokens in expression.');
  }

  return result;
}

function parseDefaultValuesJson(raw: string | undefined): Record<string, any> {
  if (!raw) {
    return {};
  }

  try {
    var parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
      return parsed as Record<string, any>;
    }
  } catch (_error) {
    // Ignore invalid JSON and fall back to schema defaults.
  }

  return {};
}

function resolveDefaultValue(value: any, context: any): any {
  if (value && typeof value === 'object' && !Array.isArray(value) && typeof value.$expression === 'string') {
    return resolveSafeExpression(value.$expression, context);
  }
  return value;
}

function parseAdvancedValidationRulesJson(raw: string | undefined): AdvancedValidationRule[] | null {
  var source = String(raw || '').trim();
  if (!source) {
    return [];
  }

  try {
    var parsed = JSON.parse(source);
    if (!Array.isArray(parsed)) {
      return null;
    }

    var rules: AdvancedValidationRule[] = [];
    for (var i = 0; i < parsed.length; i += 1) {
      var item = parsed[i];
      if (!item || typeof item !== 'object') {
        return null;
      }

      var expression = item.expression === undefined || item.expression === null ? '' : String(item.expression).trim();
      var message = item.message === undefined || item.message === null ? '' : String(item.message).trim();
      var targetField = item.targetField === undefined || item.targetField === null ? '' : String(item.targetField).trim();

      if (!expression || !message) {
        return null;
      }

      rules.push({
        id: item.id ? String(item.id) : undefined,
        expression: expression,
        message: message,
        targetField: targetField || undefined,
      });
    }

    return rules;
  } catch (_error) {
    return null;
  }
}

function normalizeDateTimeLocalValue(value: any, timeZone?: string): string {
  if (value === undefined || value === null) {
    return '';
  }
  var text = String(value).trim();
  if (!text) {
    return '';
  }
  var zone = timeZone === 'local' ? 'local' : 'UTC';

  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(text)) {
    return text;
  }

  var parsed = new Date(text);
  if (isNaN(parsed.getTime())) {
    return text;
  }

  var pad = function(num: number): string { return num < 10 ? '0' + String(num) : String(num); };
  if (zone === 'local') {
    return parsed.getFullYear() + '-' + pad(parsed.getMonth() + 1) + '-' + pad(parsed.getDate()) + 'T' + pad(parsed.getHours()) + ':' + pad(parsed.getMinutes());
  }
  return parsed.getUTCFullYear() + '-' + pad(parsed.getUTCMonth() + 1) + '-' + pad(parsed.getUTCDate()) + 'T' + pad(parsed.getUTCHours()) + ':' + pad(parsed.getUTCMinutes());
}

function normalizeDateOnlyValue(value: any, timeZone?: string): string {
  if (value === undefined || value === null) {
    return '';
  }
  var text = String(value).trim();
  if (!text) {
    return '';
  }
  var zone = timeZone === 'local' ? 'local' : 'UTC';

  if (zone === 'UTC') {
    var isoDate = /^(\d{4}-\d{2}-\d{2})/.exec(text);
    if (isoDate) {
      return isoDate[1];
    }
  }

  var parsed = new Date(text);
  if (isNaN(parsed.getTime())) {
    return text;
  }
  var pad = function(num: number): string { return num < 10 ? '0' + String(num) : String(num); };
  if (zone === 'local') {
    return parsed.getFullYear() + '-' + pad(parsed.getMonth() + 1) + '-' + pad(parsed.getDate());
  }
  return parsed.getUTCFullYear() + '-' + pad(parsed.getUTCMonth() + 1) + '-' + pad(parsed.getUTCDate());
}

function normalizeTimeOnlyValue(value: any, timeZone?: string): string {
  if (value === undefined || value === null) {
    return '';
  }
  var text = String(value).trim();
  if (!text) {
    return '';
  }
  var zone = timeZone === 'local' ? 'local' : 'UTC';

  if (/^\d{2}:\d{2}$/.test(text)) {
    return text;
  }

  var parsed = new Date(text);
  if (isNaN(parsed.getTime())) {
    return text;
  }
  var pad = function(num: number): string { return num < 10 ? '0' + String(num) : String(num); };
  return zone === 'local'
    ? pad(parsed.getHours()) + ':' + pad(parsed.getMinutes())
    : pad(parsed.getUTCHours()) + ':' + pad(parsed.getUTCMinutes());
}

/** Builds the ISO datetime string sent back to SharePoint for a datetime field, honoring its configured format/zone. */
function buildDateTimePayloadValue(value: FieldValue, config?: FieldConfig): string | null {
  if (!value) {
    return null;
  }
  var text = String(value);
  var zone = config && config.timeZone === 'local' ? 'local' : 'UTC';
  var format = config && config.displayFormat;

  if (format === 'timeOnly') {
    var timeMatch = /^(\d{2}):(\d{2})/.exec(text);
    if (!timeMatch) {
      return null;
    }
    // anchor time-only values to a fixed reference date since SharePoint DateTime columns always store a full date
    var isoAnchor = '1970-01-01T' + timeMatch[1] + ':' + timeMatch[2] + ':00' + (zone === 'UTC' ? 'Z' : '');
    return new Date(isoAnchor).toISOString();
  }

  if (format === 'dateOnly') {
    var dateMatch = /^(\d{4}-\d{2}-\d{2})/.exec(text);
    if (!dateMatch) {
      return null;
    }
    return new Date(dateMatch[1] + 'T00:00:00' + (zone === 'UTC' ? 'Z' : '')).toISOString();
  }

  var isoDateTime = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(text) ? text + ':00' + (zone === 'UTC' ? 'Z' : '') : text;
  var parsedDateTime = new Date(isoDateTime);
  return isNaN(parsedDateTime.getTime()) ? null : parsedDateTime.toISOString();
}

function getLookupIdValue(value: any): string {
  if (value === undefined || value === null || value === '') {
    return '';
  }

  if (Array.isArray(value)) {
    return value.length > 0 ? getLookupIdValue(value[0]) : '';
  }

  if (typeof value === 'object') {
    var lookupCandidate = value.selectedItemId !== undefined ? value.selectedItemId
      : value.lookupId !== undefined ? value.lookupId
        : value.LookupId !== undefined ? value.LookupId
          : value.Id !== undefined ? value.Id
            : value.ID !== undefined ? value.ID
              : value.id;
    return getLookupIdValue(lookupCandidate);
  }

  var lookupIdText = String(value).trim();
  if (!/^\d+$/.test(lookupIdText)) {
    return '';
  }
  var parsedLookupId = parseInt(lookupIdText, 10);
  return parsedLookupId > 0 ? String(parsedLookupId) : '';
}

function getListControlFilterValue(field: FormField | null, value: any): any {
  if (value === undefined || value === null) {
    return '';
  }
  if (field && (field.type === 'lookup' || field.type === 'person')) {
    return getLookupIdValue(value);
  }
  if (Array.isArray(value)) {
    return value.map(function(entry) { return String(entry); }).join(';');
  }
  if (typeof value === 'object') {
    if (value.Url !== undefined) {
      return String(value.Url);
    }
    if (value.url !== undefined) {
      return String(value.url);
    }
    return '';
  }
  return value;
}

function getItemFieldValue(item: any, fieldName: string): any {
  if (!item || !fieldName) { return undefined; }
  if (item[fieldName] !== undefined && item[fieldName] !== null) { return item[fieldName]; }
  var normalizedName = String(fieldName).toLowerCase();
  var keys = Object.keys(item);
  for (var i = 0; i < keys.length; i += 1) {
    if (keys[i].toLowerCase() === normalizedName && item[keys[i]] !== undefined && item[keys[i]] !== null) { return item[keys[i]]; }
  }
  var idName = normalizedName + 'id';
  for (var j = 0; j < keys.length; j += 1) {
    if (keys[j].toLowerCase() === idName) { return item[keys[j]]; }
  }
  return undefined;
}

function normalizeDefaultForField(field: FormField, value: any): FieldValue {
  if (value === undefined || value === null) {
    return '';
  }

  switch (field.type) {
    case 'number':
      if (value === '') {
        return '';
      }
      var parsedNumber = Number(value);
      return isNaN(parsedNumber) ? '' : parsedNumber;
    case 'boolean':
      if (typeof value === 'boolean') {
        return value;
      }
      var boolText = String(value).trim().toLowerCase();
      return boolText === 'true' || boolText === '1' || boolText === 'yes';
    case 'datetime':
      if (field.config && field.config.displayFormat === 'dateOnly') {
        return normalizeDateOnlyValue(value, field.config.timeZone);
      }
      if (field.config && field.config.displayFormat === 'timeOnly') {
        return normalizeTimeOnlyValue(value, field.config.timeZone);
      }
      return normalizeDateTimeLocalValue(value, field.config && field.config.timeZone);
    case 'multiselect':
      if (Array.isArray(value)) {
        return value.map(function(entry) { return String(entry); });
      }
      return String(value)
        .split(/[;,]/)
        .map(function(entry) { return entry.trim(); })
        .filter(function(entry) { return !!entry; });
    case 'lookup':
      var lookupId = getLookupIdValue(value);
      return field.config && field.config.allowMultiple === true
        ? (lookupId ? [lookupId] : [])
        : lookupId;
    default:
      return String(value);
  }
}

// Blank/empty representation used when a field's raw value is null/undefined on a newly loaded
// record, so navigating records never leaves a previous record's value showing for a blank field.
function getBlankValueForField(field: FormField): FieldValue {
  switch (field.type) {
    case 'multiselect':
      return [];
    case 'lookup':
    case 'person':
      return field.config && field.config.allowMultiple === true ? [] : '';
    case 'boolean':
      return false;
    default:
      return '';
  }
}

function decodeHtmlEntities(value: string): string {
  if (!value || value.indexOf('&') < 0 || typeof document === 'undefined') {
    return value;
  }
  var decoder = document.createElement('textarea');
  decoder.innerHTML = value;
  return decoder.value;
}

function isRichTextMultilineField(field: FormField): boolean {
  return field.type === 'multiline' && !!(field.config && field.config.isRichText === true);
}

function hasUsableValue(value: any): boolean {
  if (value === undefined || value === null) {
    return false;
  }
  if (typeof value === 'string') {
    return value.trim() !== '';
  }
  return true;
}

function hasTextValue(value: string | undefined): boolean {
  return !!(value && value.trim());
}

function getStepColumns(schema: FormSchema, step: any): number {
  var theme = step && step.theme ? step.theme : schema.theme;
  var columns = theme && theme.columns ? theme.columns : 1;
  return columns > 1 ? columns : 1;
}

function isGridLayout(schema: FormSchema, step: any): boolean {
  var theme = step && step.theme ? step.theme : schema.theme;
  return !!(theme && theme.layout === 'grid' && getStepColumns(schema, step) > 1);
}

export class SharePointDynamicFormContainer extends React.Component<SharePointDynamicFormContainerProps, SharePointDynamicFormContainerState> {
  private _isMounted: boolean;
  private _lastDynamicConsoleSnapshot: string;
  private _listFieldNameLookup: { [lowerInternalName: string]: boolean };
  private _listFieldInternalNameLookup: { [lowerInternalName: string]: string };
  private _listFieldTypeLookup: { [lowerInternalName: string]: string };
  private _urlFieldDescriptions: { [fieldId: string]: string };
  private _existingAttachments: { [fieldId: string]: IAttachmentItem[] };
  private _attachmentsMarkedForDelete: { [fieldId: string]: IAttachmentItem[] };
  private _attachmentInputRefs: { [fieldId: string]: HTMLInputElement | null };

  public constructor(props: SharePointDynamicFormContainerProps) {
    super(props);
    this._isMounted = false;
    this._lastDynamicConsoleSnapshot = '';
    this._listFieldNameLookup = {};
    this._listFieldInternalNameLookup = {};
    this._listFieldTypeLookup = {};
    this._urlFieldDescriptions = {};
    this._existingAttachments = {};
    this._attachmentsMarkedForDelete = {};
    this._attachmentInputRefs = {};
    this.state = {
      schema: null,
      values: {},
      loadedItem: null,
      lookupOptions: {},
      resolvedItemId: 0,
      isViewEditing: false,
      loading: true,
      error: null,
      fieldErrors: {},
      isSubmitting: false,
      submitError: null,
      submitSuccess: null,
      canAddRecords: true,
      canEditRecords: true,
      permissionStatusMessage: null,
    };

    this.handleSubmit = this.handleSubmit.bind(this);
    this.handleCancel = this.handleCancel.bind(this);
    this.handleDesignerChange = this.handleDesignerChange.bind(this);
    this.handleDesignerSave = this.handleDesignerSave.bind(this);
  }

  public componentDidMount(): void {
    this._isMounted = true;
    console.log('[ReportFormsWebPart] Component mounted. Props:', {
      linkedFieldTarget: this.props.linkedFieldTarget,
      linkedFieldValue: this.props.linkedFieldValue,
      useDynamicValueAsFilter: this.props.useDynamicValueAsFilter,
      filterJson: this.props.filterJson ? '(set)' : '(empty)',
      listName: this.props.listName,
      mode: this.props.mode,
      dynamicItemId: this.props.dynamicItemId,
    });
    this.loadData();
  }

  public componentDidUpdate(prevProps: SharePointDynamicFormContainerProps): void {
    var requiresFullReload = (
      prevProps.formSchemaJson !== this.props.formSchemaJson ||
      prevProps.listName !== this.props.listName ||
      prevProps.mode !== this.props.mode ||
      prevProps.useDynamicValueAsFilter !== this.props.useDynamicValueAsFilter ||
      prevProps.filterJson !== this.props.filterJson ||
      prevProps.itemId !== this.props.itemId ||
      prevProps.itemIdQueryParam !== this.props.itemIdQueryParam ||
      prevProps.useItemId !== this.props.useItemId ||
      prevProps.defaultValuesJson !== this.props.defaultValuesJson ||
      prevProps.advancedValidationJson !== this.props.advancedValidationJson ||
      prevProps.advancedValidationEnabled !== this.props.advancedValidationEnabled ||
      prevProps.permissionBaseLookupField !== this.props.permissionBaseLookupField ||
      prevProps.permissionScope !== this.props.permissionScope ||
      prevProps.permissionScopeItemId !== this.props.permissionScopeItemId
    );

    if (requiresFullReload) {
      console.log('[ReportFormsWebPart] Props changed, reloading data');
      this.loadData();
      return;
    }

    var selectionPropsChanged =
      prevProps.dynamicItemId !== this.props.dynamicItemId ||
      prevProps.linkedFieldTarget !== this.props.linkedFieldTarget ||
      prevProps.linkedFieldValue !== this.props.linkedFieldValue;

    if (selectionPropsChanged) {
      console.log('[ReportFormsWebPart] Selection-related props changed, refreshing selection values');
      this.refreshSelectionValues();
    }
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  private setSafeState(nextState: Partial<SharePointDynamicFormContainerState>): void {
    if (this._isMounted) {
      this.setState(nextState as SharePointDynamicFormContainerState);
    }
  }

  private applyLinkedFieldValue(values: IValueMap, fields: FormField[]): void {
    var linkedFieldTarget = (this.props.linkedFieldTarget || '').trim().toLowerCase();
    if (!linkedFieldTarget || !hasUsableValue(this.props.linkedFieldValue)) {
      return;
    }

    for (var i = 0; i < fields.length; i += 1) {
      var candidate = fields[i];
      var matchesById = candidate.id && candidate.id.toLowerCase() === linkedFieldTarget;
      var matchesByFieldName = candidate.fieldName && candidate.fieldName.toLowerCase() === linkedFieldTarget;
      if (matchesById || matchesByFieldName) {
        values[candidate.id] = normalizeDefaultForField(candidate, this.props.linkedFieldValue);
        return;
      }
    }
  }

  private async refreshSelectionValues(): Promise<void> {
    var schema = this.state.schema;
    if (!schema) {
      this.loadData();
      return;
    }

    var resolvedItemId = this.getResolvedItemId(schema);
    var fields = getAllFields(schema);

    if (resolvedItemId <= 0 || (this.props.mode !== 'edit' && this.props.mode !== 'view')) {
      this.loadData();
      return;
    }

    try {
      var listName = schema.listName || this.props.listName;
      var item = await this.loadItem(listName, resolvedItemId);
      if (!item) {
        console.log('[ReportFormsWebPart] refreshSelectionValues: item not found for resolvedItemId=' + resolvedItemId + ', falling back to full loadData');
        this.loadData();
        return;
      }

      var values = this.mapItemToValues(schema, item, this.state.values || {});
      await this.loadAttachmentMetadata(listName, resolvedItemId, fields);
      this.applyLinkedFieldValue(values, fields);

      this.setSafeState({
        values: values,
        loadedItem: item,
        resolvedItemId: resolvedItemId,
        isViewEditing: false,
        loading: false,
        error: null,
      });
      console.log('[ReportFormsWebPart] refreshSelectionValues SUCCESS - resolvedItemId=' + resolvedItemId);
    } catch (selectionUpdateError) {
      console.error('[ReportFormsWebPart] refreshSelectionValues ERROR, falling back to full loadData: ', selectionUpdateError);
      this.loadData();
    }
  }

  private getEffectiveMode(): FormMode {
    return 'view';
  }

  private getUrlQueryValue(paramName: string): string {
    if (!paramName || typeof window === 'undefined' || !window.location || !window.location.search) {
      return '';
    }

    var requestedKey = String(paramName).trim().toLowerCase();
    if (!requestedKey) {
      return '';
    }

    var query = window.location.search.replace(/^\?/, '').split('&');
    for (var i = 0; i < query.length; i += 1) {
      var segment = String(query[i] || '');
      if (!segment) {
        continue;
      }

      var separatorIndex = segment.indexOf('=');
      if (separatorIndex < 0) {
        continue;
      }

      var rawKey = segment.substring(0, separatorIndex);
      var rawValue = segment.substring(separatorIndex + 1);
      var key = decodeURIComponent(rawKey.replace(/\+/g, '%20')).toLowerCase();
      if (key === requestedKey) {
        return decodeURIComponent(rawValue.replace(/\+/g, '%20'));
      }
    }

    return '';
  }

  private getReturnUrlQueryValue(): string {
    var returnUrl = String(this.getUrlQueryValue('return') || '').trim();
    if (!returnUrl) {
      returnUrl = String(this.getUrlQueryValue('returnurl') || '').trim();
    }
    if (!returnUrl) {
      returnUrl = String(this.getUrlQueryValue('source') || '').trim();
    }

    // Some integrations double-encode query values. Decode up to two extra passes.
    for (var i = 0; i < 2; i += 1) {
      if (returnUrl.indexOf('%') < 0) {
        break;
      }
      try {
        var decoded = decodeURIComponent(returnUrl);
        if (!decoded || decoded === returnUrl) {
          break;
        }
        returnUrl = decoded;
      } catch (_decodeError) {
        break;
      }
    }

    return returnUrl;
  }

  private getPreferredRedirectUrl(redirectUrl: string | undefined): string {
    var configuredRedirect = String(redirectUrl || '').trim();
    if (configuredRedirect) {
      return configuredRedirect;
    }

    return String(this.getReturnUrlQueryValue() || '').trim();
  }

  private isSafeRedirectUrl(targetUrl: string): boolean {
    if (!targetUrl || typeof window === 'undefined' || !window.location) {
      return false;
    }

    if (targetUrl.charAt(0) === '/') {
      return true;
    }

    try {
      var parsedTarget = new URL(targetUrl, window.location.href);
      return parsedTarget.origin === window.location.origin;
    } catch (_parseError) {
      return false;
    }
  }

  private tryRedirect(targetUrl: string): boolean {
    if (!this.isSafeRedirectUrl(targetUrl) || typeof window === 'undefined') {
      return false;
    }

    window.location.assign(targetUrl);
    return true;
  }

  private getResolvedItemId(schema?: FormSchema | null): number {
    var fromDynamic = toPositiveInt(this.props.dynamicItemId);
    var hasDynamicBinding = !!String(this.props.dynamicItemReference || '').trim();
    var fromContext = 0;
    var pageContext = this.props.context && this.props.context.pageContext;
    var targetListName = String((schema && schema.listName) || this.props.listName || '').trim().toLowerCase();
    var contextListName = pageContext && pageContext.list && pageContext.list.title
      ? String(pageContext.list.title).trim().toLowerCase()
      : '';
    if (targetListName && targetListName === contextListName && pageContext.listItem && pageContext.listItem.id) {
      fromContext = toPositiveInt(pageContext.listItem.id);
    }

    var requestedKey = (this.props.itemIdQueryParam || 'itemid').trim();
    if (!requestedKey) {
      requestedKey = 'itemid';
    }
    var fromUrl = toPositiveInt(this.getUrlQueryValue(requestedKey));

    var fromProps = toPositiveInt(this.props.itemId);
    var fromSchema = schema ? toPositiveInt(schema.itemId) : 0;

    if (this.props.mode === 'edit' || this.props.mode === 'view' || this.props.useItemId) {
      if (this.props.useDynamicValueAsFilter && String(this.props.linkedFieldTarget || '').trim()) {
        return fromProps || fromSchema || 0;
      }
      // If a dynamic source is bound, prefer live dynamic ID and do not fall back
      // to stale URL/context item IDs from the hosting page.
      if (hasDynamicBinding) {
        return fromDynamic || fromProps || fromSchema || 0;
      }
      return fromDynamic || fromUrl || fromProps || fromSchema || fromContext || 0;
    }

    return fromDynamic || fromProps || fromSchema || 0;
  }

  private getWebUrl(): string {
    return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
  }

  private cloneHeaders(source?: { [key: string]: string }): { [key: string]: string } {
    var cloned: { [key: string]: string } = {};
    if (!source) {
      return cloned;
    }

    for (var key in source) {
      if (Object.prototype.hasOwnProperty.call(source, key)) {
        cloned[key] = source[key];
      }
    }

    return cloned;
  }

  private async getWithAcceptFallback(url: string): Promise<any> {
    var response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=verbose'
        }
      });
    }

    if (!response.ok) {
      response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=minimalmetadata'
        }
      });
    }

    if (!response.ok) {
      response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      });
    }

    return response;
  }

  private async postWithAcceptFallback(url: string, payload: any, baseHeaders?: { [key: string]: string }): Promise<any> {
    var headers = this.cloneHeaders(baseHeaders);
    if (!headers['Content-Type']) {
      headers['Content-Type'] = 'application/json; charset=utf-8';
    }

    var body = JSON.stringify(payload || {});
    var response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: headers,
      body: body,
    });

    if (!response.ok) {
      var verboseHeaders = this.cloneHeaders(baseHeaders);
      verboseHeaders.Accept = 'application/json;odata=verbose';
      verboseHeaders['Content-Type'] = 'application/json;odata=verbose';
      response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: verboseHeaders,
        body: body,
      });
    }

    if (!response.ok) {
      var minimalHeaders = this.cloneHeaders(baseHeaders);
      minimalHeaders.Accept = 'application/json;odata=minimalmetadata';
      minimalHeaders['Content-Type'] = 'application/json;odata=minimalmetadata';
      response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: minimalHeaders,
        body: body,
      });
    }

    if (!response.ok) {
      var noMetadataHeaders = this.cloneHeaders(baseHeaders);
      noMetadataHeaders.Accept = 'application/json;odata=nometadata';
      noMetadataHeaders['Content-Type'] = 'application/json;odata=nometadata';
      response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: noMetadataHeaders,
        body: body,
      });
    }

    return response;
  }

  private async postRawWithAcceptFallback(url: string, body: any, baseHeaders?: { [key: string]: string }): Promise<any> {
    var acceptCandidates = [
      'application/json;odata=verbose',
      'application/json;odata=minimalmetadata',
      'application/json;odata=nometadata',
      'application/json',
      '*/*'
    ];

    var response: any = null;
    for (var i = 0; i < acceptCandidates.length; i += 1) {
      var headers = this.cloneHeaders(baseHeaders);
      headers.Accept = acceptCandidates[i];
      response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: headers,
        body: body,
      });

      if (response.ok) {
        return response;
      }
    }

    // Final attempt with no explicit Accept header.
    var noAcceptHeaders = this.cloneHeaders(baseHeaders);
    if (Object.prototype.hasOwnProperty.call(noAcceptHeaders, 'Accept')) {
      delete noAcceptHeaders.Accept;
    }
    response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: noAcceptHeaders,
      body: body,
    });

    return response;
  }

  private logDiagnostic(message: string): void {
    if (this.props.enableDynamicDiagnostics !== false) {
      console.log('[ReportFormsWebPart] ' + message);
    }
  }

  private logDynamicDiagnosticsToConsole(summary: string): void {
    if (this.props.enableDynamicDiagnostics === false) {
      this._lastDynamicConsoleSnapshot = '';
      return;
    }

    var entries = this.props.dynamicDiagnostics || [];
    var snapshot = summary + '||' + entries.join('|');
    if (snapshot === this._lastDynamicConsoleSnapshot) {
      return;
    }

    this._lastDynamicConsoleSnapshot = snapshot;
    console.log('[ReportFormsWebPart] ' + summary);
    for (var i = 0; i < entries.length; i += 1) {
      console.log('[ReportFormsWebPart] Dynamic diagnostics: ' + entries[i]);
    }
  }

  private buildListApiPathByLookupReference(lookupList: string): string {
    var cleaned = String(lookupList || '').trim().replace(/^\{/, '').replace(/\}$/, '');
    if (!cleaned) {
      return '';
    }

    if (isGuid(cleaned)) {
      return "/_api/web/lists(guid'" + cleaned + "')";
    }

    return "/_api/web/lists/getByTitle('" + escapeODataText(cleaned) + "')";
  }

  private extractPermissionFlags(raw: any): { canAdd: boolean; canEdit: boolean } {
    var source = raw || {};
    var base = source.EffectiveBasePermissions || (source.d && source.d.EffectiveBasePermissions);
    if (!base) {
      return { canAdd: false, canEdit: false };
    }

    var effectivePermissions = new SPPermission(base);
    return {
      canAdd: effectivePermissions.hasPermission(SPPermission.addListItems),
      canEdit: effectivePermissions.hasPermission(SPPermission.viewListItems),
    };
  }

  private async getListPermissionFlagsByTitle(listTitle: string): Promise<{ canAdd: boolean; canEdit: boolean }> {
    var response = await this.getWithAcceptFallback(
      this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(listTitle) + "')?$select=EffectiveBasePermissions"
    );
    if (!response.ok) {
      throw new Error('Failed to load list permissions for "' + listTitle + '". Status: ' + response.status);
    }
    return this.extractPermissionFlags(await response.json());
  }

  private async getListPermissionFlagsByLookupRef(lookupList: string): Promise<{ canAdd: boolean; canEdit: boolean }> {
    var listPath = this.buildListApiPathByLookupReference(lookupList);
    if (!listPath) {
      throw new Error('Lookup list reference is empty.');
    }

    var response = await this.getWithAcceptFallback(
      this.getWebUrl() + listPath + '?$select=EffectiveBasePermissions'
    );
    if (!response.ok) {
      throw new Error('Failed to load lookup list permissions. Status: ' + response.status);
    }

    return this.extractPermissionFlags(await response.json());
  }

  private async getLookupItemPermissionFlags(lookupList: string, lookupItemId: number): Promise<{ canAdd: boolean; canEdit: boolean }> {
    var listPath = this.buildListApiPathByLookupReference(lookupList);
    if (!listPath || lookupItemId <= 0) {
      throw new Error('Cannot evaluate lookup item permissions because list or item ID is invalid.');
    }

    var response = await this.getWithAcceptFallback(
      this.getWebUrl() + listPath + '/items(' + lookupItemId + ')?$select=EffectiveBasePermissions'
    );
    if (!response.ok) {
      throw new Error('Failed to load lookup item permissions. Status: ' + response.status);
    }

    return this.extractPermissionFlags(await response.json());
  }

  private async getLookupPermissionFieldMetadata(listName: string, internalName: string): Promise<ILookupPermissionFieldMetadata | null> {
    if (!listName || !internalName) {
      return null;
    }

    var endpoint = this.getWebUrl()
      + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/fields/getByInternalNameOrTitle('"
      + escapeODataText(internalName)
      + "')?$select=InternalName,TypeAsString,LookupList";

    var response = await this.getWithAcceptFallback(endpoint);
    if (!response.ok) {
      this.logDiagnostic('Permission metadata lookup failed for column "' + internalName + '". Status=' + response.status);
      return null;
    }

    var data = await response.json();
    var field = data && data.d ? data.d : data;
    if (!field) {
      return null;
    }

    var typeAsString = String(field.TypeAsString || '');
    var lookupList = String(field.LookupList || '');
    if (!lookupList || (typeAsString.toLowerCase() !== 'lookup' && typeAsString.toLowerCase() !== 'lookupmulti')) {
      return null;
    }

    return {
      internalName: String(field.InternalName || internalName),
      typeAsString: typeAsString,
      lookupList: lookupList,
    };
  }

  private toLookupItemId(value: any): number {
    if (Array.isArray(value)) {
      for (var i = 0; i < value.length; i += 1) {
        var fromArray = this.toLookupItemId(value[i]);
        if (fromArray > 0) {
          return fromArray;
        }
      }
      return 0;
    }

    if (value && typeof value === 'object') {
      var idCandidate = (value as any).Id || (value as any).ID || (value as any).id || (value as any).lookupId;
      if (idCandidate !== undefined && idCandidate !== null && idCandidate !== '') {
        return this.toLookupItemId(idCandidate);
      }

      var results = (value as any).results || (value as any).value;
      if (results !== undefined) {
        return this.toLookupItemId(results);
      }
    }

    return toPositiveInt(value);
  }

  private tryResolvePermissionLookupItemId(
    metadata: ILookupPermissionFieldMetadata,
    values: IValueMap,
    schema: FormSchema,
    loadedItem: any
  ): number {
    if (loadedItem) {
      var loadedItemCandidate = loadedItem[metadata.internalName + 'Id'] !== undefined
        ? loadedItem[metadata.internalName + 'Id']
        : loadedItem[metadata.internalName];
      var fromLoadedItem = this.toLookupItemId(loadedItemCandidate);
      if (fromLoadedItem > 0) {
        return fromLoadedItem;
      }
    }

    var fields = getAllFields(schema);
    for (var i = 0; i < fields.length; i += 1) {
      var field = fields[i];
      if (!field || !field.fieldName) {
        continue;
      }

      if (String(field.fieldName).toLowerCase() !== String(metadata.internalName).toLowerCase()) {
        continue;
      }

      var candidate = values[field.id];
      var fromValues = this.toLookupItemId(candidate);
      if (fromValues > 0) {
        return fromValues;
      }
    }

    return 0;
  }

  private async evaluateButtonPermissions(
    schema: FormSchema,
    values: IValueMap,
    loadedItem: any
  ): Promise<{ canAdd: boolean; canEdit: boolean; message: string }> {
    var baseListName = String(schema.listName || this.props.listName || '').trim();
    var selectedPermissionColumn = String(this.props.permissionBaseLookupField || '').trim();
    var selectedScope = this.props.permissionScope === 'item' ? 'item' : 'list';

    if (!baseListName) {
      return {
        canAdd: false,
        canEdit: false,
        message: '',
      };
    }

    if (!selectedPermissionColumn) {
      var defaultListPermissions = await this.getListPermissionFlagsByTitle(baseListName);
      return {
        canAdd: defaultListPermissions.canAdd,
        canEdit: defaultListPermissions.canEdit,
        message: 'Permission gate source=default-list("' + baseListName + '"): canAdd=' + defaultListPermissions.canAdd + ', canEdit=' + defaultListPermissions.canEdit,
      };
    }

    var lookupFieldMetadata = await this.getLookupPermissionFieldMetadata(baseListName, selectedPermissionColumn);
    if (!lookupFieldMetadata) {
      var fallbackListPermissions = await this.getListPermissionFlagsByTitle(baseListName);
      return {
        canAdd: fallbackListPermissions.canAdd,
        canEdit: fallbackListPermissions.canEdit,
        message: 'Permission column "' + selectedPermissionColumn + '" is not a valid lookup. Fallback source=default-list("' + baseListName + '"): canAdd=' + fallbackListPermissions.canAdd + ', canEdit=' + fallbackListPermissions.canEdit,
      };
    }

    if (selectedScope === 'item') {
      var fixedPermissionItemId = parseInt(String(this.props.permissionScopeItemId || ''), 10);
      if (!isNaN(fixedPermissionItemId) && fixedPermissionItemId > 0) {
        var fixedItemPermissions = await this.getLookupItemPermissionFlags(lookupFieldMetadata.lookupList, fixedPermissionItemId);
        return {
          canAdd: fixedItemPermissions.canAdd,
          canEdit: fixedItemPermissions.canEdit,
          message: 'Permission gate source=fixed-lookup-item(column="' + selectedPermissionColumn + '", itemId=' + fixedPermissionItemId + '): canAdd=' + fixedItemPermissions.canAdd + ', canEdit=' + fixedItemPermissions.canEdit,
        };
      }

      var lookupItemId = this.tryResolvePermissionLookupItemId(lookupFieldMetadata, values, schema, loadedItem);
      if (lookupItemId > 0) {
        var lookupItemPermissions = await this.getLookupItemPermissionFlags(lookupFieldMetadata.lookupList, lookupItemId);
        return {
          canAdd: lookupItemPermissions.canAdd,
          canEdit: lookupItemPermissions.canEdit,
          message: 'Permission gate source=lookup-item(column="' + selectedPermissionColumn + '", itemId=' + lookupItemId + '): canAdd=' + lookupItemPermissions.canAdd + ', canEdit=' + lookupItemPermissions.canEdit,
        };
      }

      var lookupListFallbackPermissions = await this.getListPermissionFlagsByLookupRef(lookupFieldMetadata.lookupList);
      return {
        canAdd: lookupListFallbackPermissions.canAdd,
        canEdit: lookupListFallbackPermissions.canEdit,
        message: 'Permission scope=item but lookup item is not resolved for column "' + selectedPermissionColumn + '". Fallback source=lookup-list: canAdd=' + lookupListFallbackPermissions.canAdd + ', canEdit=' + lookupListFallbackPermissions.canEdit,
      };
    }

    var lookupListPermissions = await this.getListPermissionFlagsByLookupRef(lookupFieldMetadata.lookupList);
    return {
      canAdd: lookupListPermissions.canAdd,
      canEdit: lookupListPermissions.canEdit,
      message: 'Permission gate source=lookup-list(column="' + selectedPermissionColumn + '"): canAdd=' + lookupListPermissions.canAdd + ', canEdit=' + lookupListPermissions.canEdit,
    };
  }

  private getPermissionDeniedMessage(schema?: FormSchema): string {
    var configured = String(schema && schema.permissionDeniedMessage || '').trim();
    return configured || strings.PermissionDeniedDefault;
  }

  private isPermissionDeniedError(error: any): boolean {
    if (!error) {
      return false;
    }
    if (error.isPermissionDenied === true || error.status === 401 || error.status === 403) {
      return true;
    }
    var message = String(error.message || error || '').toLowerCase();
    return /(^|\D)(401|403)(\D|$)/.test(message)
      || message.indexOf('access denied') >= 0
      || message.indexOf('permission denied') >= 0
      || message.indexOf('does not have permission') >= 0;
  }

  private renderDescriptionIcon(description: string): JSX.Element | null {
    if (!description) {
      return null;
    }
    console.log('[ReportFormsWebPart] renderDescriptionIcon - Rendering icon for description: "' + description + '"');
    return (
      <span 
        title={description}
        style={{ cursor: 'pointer', marginLeft: '4px', fontSize: '14px', color: '#0078d4', display: 'inline-block' }}
      >
        ℹ️
      </span>
    );
  }

  private getChoiceDisplayMode(field: FormField): 'dropdown' | 'radio' | 'checkboxes' {
    var configured = String(field.config && (field.config as any).choiceDisplay || '').trim().toLowerCase();
    if (field.type === 'multiselect') {
      return 'checkboxes';
    }
    if (configured === 'radio' || configured === 'checkboxes') {
      return configured as 'radio' | 'checkboxes';
    }
    return 'dropdown';
  }

  private getExclusiveChoiceValue(field: FormField): string {
    return String(field.config && (field.config as any).exclusiveChoiceValue || '').trim();
  }

  private updateExclusiveMultiChoiceSelection(field: FormField, currentValues: string[], optionValue: string, checked: boolean): string[] {
    var exclusiveChoiceValue = this.getExclusiveChoiceValue(field);
    var nextValues = currentValues.slice(0);
    var existingIndex = nextValues.indexOf(optionValue);
    var exclusiveIndex = exclusiveChoiceValue ? nextValues.indexOf(exclusiveChoiceValue) : -1;

    if (checked) {
      if (exclusiveChoiceValue && optionValue === exclusiveChoiceValue) {
        return [exclusiveChoiceValue];
      }

      if (exclusiveIndex >= 0) {
        nextValues.splice(exclusiveIndex, 1);
      }

      if (existingIndex < 0) {
        nextValues.push(optionValue);
      }
      return nextValues;
    }

    if (existingIndex >= 0) {
      nextValues.splice(existingIndex, 1);
    }

    return nextValues;
  }

  private isExclusiveChoiceLocked(field: FormField, selectedValues: string[]): boolean {
    var exclusiveChoiceValue = this.getExclusiveChoiceValue(field);
    return !!exclusiveChoiceValue && selectedValues.indexOf(exclusiveChoiceValue) >= 0;
  }

  private renderFieldHelpAndError(description: string, errorMessage: string, labelPosition?: string): JSX.Element | null {
    if (!description && !errorMessage) {
      return null;
    }
    console.log('[ReportFormsWebPart] renderFieldHelpAndError - description: "' + description + '", errorMessage: "' + errorMessage + '"');
    // Offset error message when label is to the left (align under input, not under label)
    var errorStyle: React.CSSProperties = { marginTop: '4px', fontSize: '12px', color: '#d32f2f', fontWeight: 'bold' };
    if (labelPosition === 'left') {
      errorStyle.marginLeft = '136px'; // 120px label width + 16px gap
    }
    return (
      <div>
        {errorMessage && <div style={errorStyle}>❌ {errorMessage}</div>}
      </div>
    );
  }

  private buildSchema(): FormSchema | null {
    if (!this.props.formSchemaJson) {
      return buildBlankSchema(this.props.listName, this.props.mode);
    }

    try {
      var parsed = JSON.parse(this.props.formSchemaJson) as FormSchema;
      if (!parsed || !parsed.steps) {
        return null;
      }
      if (!parsed.listName) {
        parsed.listName = this.props.listName;
      }
      if (!parsed.mode) {
        parsed.mode = this.props.mode;
      }

      var advancedRules = parseAdvancedValidationRulesJson(this.props.advancedValidationJson);
      if (advancedRules !== null && (String(this.props.advancedValidationJson || '').trim() || this.props.advancedValidationEnabled !== undefined)) {
        parsed.advancedValidation = parsed.advancedValidation || {};
        parsed.advancedValidation.rules = advancedRules;
        parsed.advancedValidation.enabled = this.props.advancedValidationEnabled === true && advancedRules.length > 0;
      }

      return parsed;
    } catch (error) {
      return null;
    }
  }

  private async loadData(): Promise<void> {
    var schema = this.buildSchema();
    if (!schema) {
      this.setSafeState({
        schema: null,
        values: {},
        lookupOptions: {},
        resolvedItemId: 0,
        isViewEditing: false,
        loading: false,
        error: strings.LoadConfigFailed,
        fieldErrors: {},
        isSubmitting: false,
        submitError: null,
        submitSuccess: null,
        canAddRecords: false,
        canEditRecords: false,
        permissionStatusMessage: strings.PermissionDeniedDefault,
      });
      return;
    }

    var configuredListName = String(schema.listName || this.props.listName || '').trim();
    if (!configuredListName) {
      this.setSafeState({
        schema: schema,
        values: {},
        loadedItem: null,
        lookupOptions: {},
        resolvedItemId: 0,
        isViewEditing: false,
        loading: false,
        error: null,
        fieldErrors: {},
        isSubmitting: false,
        submitError: null,
        submitSuccess: null,
        canAddRecords: false,
        canEditRecords: false,
        permissionStatusMessage: null,
      });
      return;
    }

    this.setSafeState({
      loading: true,
      error: null,
      submitError: null,
      submitSuccess: null,
      fieldErrors: {},
    });

    try {
      var values: IValueMap = {};
      var lookupOptions: ILookupMap = {};
      var loadedItem: any = null;
      this._urlFieldDescriptions = {};
      this._existingAttachments = {};
      this._attachmentsMarkedForDelete = {};
      
      // Load field descriptions from SharePoint list metadata
      await this.loadFieldDescriptions(schema);
      
      var resolvedItemId = this.getResolvedItemId(schema);
      console.log('[ReportFormsWebPart] Initial resolvedItemId from getResolvedItemId: ' + resolvedItemId);
      
      // If no item ID resolved yet, try to find a matching record via filter
      if (resolvedItemId <= 0) {
        console.log('[ReportFormsWebPart] No item ID resolved, attempting filter-based resolution...');
        resolvedItemId = await this.resolveFilteredItemId(schema);
        console.log('[ReportFormsWebPart] Filter resolution returned: ' + resolvedItemId);
      }
      var fields = getAllFields(schema);
      var webPartDefaults = parseDefaultValuesJson(this.props.defaultValuesJson);

      console.log('[ReportFormsWebPart] Building field values and loading lookup options...');
      console.log('[ReportFormsWebPart] Total fields in schema: ' + fields.length);

      for (var i = 0; i < fields.length; i += 1) {
        var field = fields[i];
        console.log('[ReportFormsWebPart] Field: ' + field.fieldName + ' (ID: ' + field.id + ', Type: ' + field.type + ')');
        var isMultiLookup = field.type === 'lookup' && !!(field.config && field.config.allowMultiple === true);
        var isMultiPerson = field.type === 'person' && !!(field.config && field.config.allowMultiple === true);
        
        if (field.defaultValue !== undefined) {
          values[field.id] = field.defaultValue;
        } else if (field.type === 'multiselect' || isMultiLookup || isMultiPerson || field.type === 'attachment') {
          values[field.id] = [];
        } else if (field.type === 'boolean') {
          values[field.id] = false;
        } else {
          values[field.id] = '';
        }

        if (field.type === 'lookup') {
          console.log('[ReportFormsWebPart] LOOKUP FIELD FOUND: ' + field.fieldName);
          console.log('[ReportFormsWebPart]   - field.config exists? ' + (!!field.config));
          if (field.config) {
            console.log('[ReportFormsWebPart]   - field.config.lookupList: ' + (field.config.lookupList || '(NOT SET)'));
            console.log('[ReportFormsWebPart]   - field.config.lookupField: ' + (field.config.lookupField || '(NOT SET, will default to "Title")'));
          } else {
            console.log('[ReportFormsWebPart]   - ERROR: field.config is NULL! Lookup options will not load!');
          }
          
          if (field.config && field.config.lookupList) {
            console.log('[ReportFormsWebPart] ✓ LOADING lookup options for field: ' + field.fieldName + ' from list: ' + field.config.lookupList);
            try {
              lookupOptions[field.id] = await this.loadLookupOptions(field.config.lookupList, field.config.lookupField || 'Title');
              console.log('[ReportFormsWebPart] ✓ Loaded ' + lookupOptions[field.id].length + ' lookup options for field: ' + field.fieldName);
            } catch (lookupError) {
              console.error('[ReportFormsWebPart] ✗ ERROR loading lookup options for ' + field.fieldName + ': ', lookupError);
              lookupOptions[field.id] = [];
            }
          } else {
            console.log('[ReportFormsWebPart] ✗ SKIPPING lookup load: field.config or lookupList not set for ' + field.fieldName);
            lookupOptions[field.id] = [];
          }
        } else if (field.type === 'person') {
          try {
            lookupOptions[field.id] = await this.loadPersonOptions();
            console.log('[ReportFormsWebPart] ✓ Loaded ' + lookupOptions[field.id].length + ' person options for field: ' + field.fieldName);
          } catch (personError) {
            console.error('[ReportFormsWebPart] ✗ ERROR loading person options for ' + field.fieldName + ': ', personError);
            lookupOptions[field.id] = [];
          }
        }
      }

      if (resolvedItemId <= 0) {
        for (var j = 0; j < fields.length; j += 1) {
          var candidateField = fields[j];
          if (Object.prototype.hasOwnProperty.call(webPartDefaults, candidateField.id)) {
            values[candidateField.id] = normalizeDefaultForField(candidateField, resolveDefaultValue(webPartDefaults[candidateField.id], this.props.context));
            continue;
          }

          if (Object.prototype.hasOwnProperty.call(webPartDefaults, candidateField.fieldName)) {
            values[candidateField.id] = normalizeDefaultForField(candidateField, resolveDefaultValue(webPartDefaults[candidateField.fieldName], this.props.context));
          }
        }
      }

      if (resolvedItemId > 0 && (this.props.mode === 'edit' || this.props.mode === 'view')) {
        var itemListName = schema.listName || this.props.listName;
        loadedItem = await this.loadItem(itemListName, resolvedItemId);
        if (loadedItem) {
          values = this.mapItemToValues(schema, loadedItem, values);
          await this.loadAttachmentMetadata(itemListName, resolvedItemId, fields);
        } else {
          // Item no longer exists; continue as a new form instead of failing load.
          resolvedItemId = 0;
          for (var defaultIndex = 0; defaultIndex < fields.length; defaultIndex += 1) {
            var defaultCandidateField = fields[defaultIndex];
            if (Object.prototype.hasOwnProperty.call(webPartDefaults, defaultCandidateField.id)) {
              values[defaultCandidateField.id] = normalizeDefaultForField(defaultCandidateField, resolveDefaultValue(webPartDefaults[defaultCandidateField.id], this.props.context));
              continue;
            }

            if (Object.prototype.hasOwnProperty.call(webPartDefaults, defaultCandidateField.fieldName)) {
              values[defaultCandidateField.id] = normalizeDefaultForField(defaultCandidateField, resolveDefaultValue(webPartDefaults[defaultCandidateField.fieldName], this.props.context));
            }
          }
        }
      }

      this.applyLinkedFieldValue(values, fields);

      var permissionResult = await this.evaluateButtonPermissions(schema, values, loadedItem);
      this.logDiagnostic(permissionResult.message);

      this.setSafeState({
        schema: schema,
        values: values,
        loadedItem: loadedItem,
        lookupOptions: lookupOptions,
        resolvedItemId: resolvedItemId,
        isViewEditing: false,
        loading: false,
        error: null,
        canAddRecords: permissionResult.canAdd,
        canEditRecords: permissionResult.canEdit,
        permissionStatusMessage: this.getPermissionDeniedMessage(schema),
      });
      console.log('[ReportFormsWebPart] Form data loaded. Resolved Item ID: ' + resolvedItemId + ', Values: ', values);
    } catch (error) {
      var readPermissionDenied = this.isPermissionDeniedError(error);
      var displayError = readPermissionDenied
        ? this.getPermissionDeniedMessage(schema)
        : (error && error.message ? error.message : strings.LoadConfigFailed);
      var permissionFailureMessage = readPermissionDenied
        ? 'Read permission was denied. Showing the configured permission denied message.'
        : 'Permission evaluation failed. Disabling Add/Edit. Details: ' + (error && error.message ? error.message : strings.CommonUnknownError);
      this.logDiagnostic(permissionFailureMessage);
      this.setSafeState({
        schema: schema,
        values: {},
        loadedItem: null,
        lookupOptions: {},
        resolvedItemId: 0,
        isViewEditing: false,
        loading: false,
        error: displayError,
        canAddRecords: false,
        canEditRecords: false,
        permissionStatusMessage: this.getPermissionDeniedMessage(schema),
      });
    }
  }

  private async loadItem(listName: string, itemId: number): Promise<any | null> {
    var itemUrl = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items(" + itemId + ")";
    var response = await this.getWithAcceptFallback(itemUrl + "?$select=*");
    if (response.status === 404) {
      this.logDiagnostic('Requested item ' + itemId + ' was not found. Falling back to new-item state.');
      return null;
    }
    if (!response.ok) {
      var loadError: any = new Error(strings.LoadErrorNoItemIdFromUrl + ' Status: ' + response.status);
      loadError.status = response.status;
      loadError.isPermissionDenied = response.status === 401 || response.status === 403;
      throw loadError;
    }
    return response.json();
  }

  private async loadAttachmentMetadata(listName: string, itemId: number, fields: FormField[]): Promise<void> {
    var attachmentFields = fields.filter(function(field) { return field.type === 'attachment'; });
    if (attachmentFields.length === 0) {
      return;
    }

    var endpoint = this.getWebUrl()
      + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items(" + itemId + ")/AttachmentFiles?$select=FileName,ServerRelativeUrl";

    try {
      var response = await this.getWithAcceptFallback(endpoint);
      if (!response.ok) {
        this.logDiagnostic('Attachment list was unavailable for item ' + itemId + '. The report will continue without attachments.');
        return;
      }

      var payload = await response.json();
      var responseData = payload && payload.d ? payload.d : payload;
      var rawAttachments = responseData && responseData.results ? responseData.results : (responseData && responseData.value ? responseData.value : responseData);
      var attachments: IAttachmentItem[] = [];
      if (Array.isArray(rawAttachments)) {
        for (var attachmentIndex = 0; attachmentIndex < rawAttachments.length; attachmentIndex += 1) {
          var rawAttachment = rawAttachments[attachmentIndex] || {};
          var attachment: IAttachmentItem = {
            fileName: rawAttachment.FileName || '',
            serverRelativeUrl: rawAttachment.ServerRelativeUrl || undefined,
          };

          if (attachment.serverRelativeUrl) {
            try {
              var fileEndpoint = this.getWebUrl()
                + "/_api/web/GetFileByServerRelativeUrl('" + escapeODataText(attachment.serverRelativeUrl) + "')?$select=Name,Length,TimeCreated,TimeLastModified,ServerRelativeUrl";
              var fileResponse = await this.getWithAcceptFallback(fileEndpoint);
              if (fileResponse.ok) {
                var filePayload = await fileResponse.json();
                var fileData = filePayload && filePayload.d ? filePayload.d : filePayload;
                attachment.fileName = fileData.Name || attachment.fileName;
                attachment.serverRelativeUrl = fileData.ServerRelativeUrl || attachment.serverRelativeUrl;
                attachment.length = fileData.Length === undefined ? undefined : parseInt(String(fileData.Length), 10);
                attachment.timeCreated = fileData.TimeCreated || undefined;
                attachment.timeLastModified = fileData.TimeLastModified || undefined;
              }
            } catch (_metadataError) {
              // Keep the attachment name and link when optional file metadata is unavailable.
            }
          }

          attachments.push(attachment);
        }
      }

      for (var fieldIndex = 0; fieldIndex < attachmentFields.length; fieldIndex += 1) {
        this._existingAttachments[attachmentFields[fieldIndex].id] = attachments.slice(0);
      }
    } catch (_attachmentError) {
      this.logDiagnostic('Attachment metadata could not be loaded for item ' + itemId + '. The report will continue without attachments.');
    }
  }

  private async resolveFilteredItemId(schema: FormSchema): Promise<number> {
    var filterExpressions: string[] = [];
    var fields = getAllFields(schema);
    var requestedKey = (this.props.itemIdQueryParam || 'itemid').trim();
    if (!requestedKey) {
      requestedKey = 'itemid';
    }
    var urlDynamicValue = this.getUrlQueryValue(requestedKey);
    var preferredDynamicValue: any = hasTextValue(urlDynamicValue) ? urlDynamicValue : this.props.linkedFieldValue;

    console.log('[ReportFormsWebPart] FILTER RESOLUTION STARTING. useDynamicValueAsFilter=' + this.props.useDynamicValueAsFilter + ', linkedFieldTarget=' + (this.props.linkedFieldTarget || '(none)') + ', linkedFieldValue=' + this.props.linkedFieldValue + ', urlDynamicValue=' + (urlDynamicValue || '(none)') + ', filterJson=' + (this.props.filterJson ? '(set)' : '(none)'));

    // Dynamic filter: use linkedFieldTarget + linkedFieldValue to find the record
    if (this.props.useDynamicValueAsFilter && this.props.linkedFieldTarget && hasUsableValue(preferredDynamicValue)) {
      console.log('[ReportFormsWebPart] FILTER ENABLED: Building dynamic filter from linkedFieldTarget');
      var targetLower = (this.props.linkedFieldTarget || '').trim().toLowerCase();
      console.log('[ReportFormsWebPart] Looking for field matching target (case-insensitive): "' + targetLower + '"');
      for (var i = 0; i < fields.length; i++) {
        var f = fields[i];
        var fieldIdMatch = f.id && f.id.toLowerCase() === targetLower;
        var fieldNameMatch = f.fieldName && f.fieldName.toLowerCase() === targetLower;
        console.log('[ReportFormsWebPart]   Checking field: id=' + (f.id || '(none)') + ', fieldName=' + (f.fieldName || '(none)') + ', idMatch=' + fieldIdMatch + ', nameMatch=' + fieldNameMatch);
        if (fieldIdMatch || fieldNameMatch) {
          console.log('[ReportFormsWebPart] FOUND MATCHING FIELD: ' + f.fieldName + ' (type=' + f.type + ')');
          var clause = buildODataFilterClause(f.fieldName, f.type, preferredDynamicValue);
          if (clause) { 
            filterExpressions.push(clause);
            console.log('[ReportFormsWebPart] Built filter clause: ' + clause);
          } else {
            console.warn('[ReportFormsWebPart] FAILED to build filter clause for field: ' + f.fieldName);
          }
          break;
        }
      }
      if (filterExpressions.length === 0) {
        var listFieldInternalName = this._listFieldInternalNameLookup[targetLower];
        var listFieldType = this._listFieldTypeLookup[targetLower];
        if (listFieldInternalName) {
          var metadataFilterType = getFilterTypeFromSharePointType(listFieldType);
          var metadataFilterValue = metadataFilterType === 'lookup' ? getLookupIdValue(preferredDynamicValue) : preferredDynamicValue;
          var listMetadataClause = buildODataFilterClause(listFieldInternalName, metadataFilterType, metadataFilterValue);
          if (listMetadataClause) {
            filterExpressions.push(listMetadataClause);
            console.log('[ReportFormsWebPart] Built filter from bound-list metadata: ' + listMetadataClause);
          }
        }
      }
      if (filterExpressions.length === 0) {
        console.warn('[ReportFormsWebPart] NO MATCHING FIELD FOUND for linkedFieldTarget: ' + this.props.linkedFieldTarget);
      }
    } else {
      if (!this.props.useDynamicValueAsFilter) {
        console.log('[ReportFormsWebPart] FILTER DISABLED: useDynamicValueAsFilter is FALSE. Enable the toggle to filter by linked field value.');
      } else if (!this.props.linkedFieldTarget) {
        console.log('[ReportFormsWebPart] FILTER SKIPPED: linkedFieldTarget is not set');
      } else {
        console.log('[ReportFormsWebPart] FILTER SKIPPED: dynamic filter value is not usable. Value: ' + preferredDynamicValue);
      }
    }

    // Static filter from filterJson property
    if (this.props.filterJson) {
      console.log('[ReportFormsWebPart] STATIC FILTER: Processing filterJson');
      var staticExpression = parseFilterJsonToExpression(this.props.filterJson, preferredDynamicValue, fields, this.props.context);
      if (staticExpression) {
        filterExpressions.push(staticExpression);
      }
      console.log('[ReportFormsWebPart] Static filter expression=' + (staticExpression || '(empty)') + '. Total expressions: ' + filterExpressions.length);
    }

    if (filterExpressions.length === 0) { 
      console.log('[ReportFormsWebPart] NO FILTER CLAUSES. Returning 0 (no record to load).');
      return 0; 
    }

    var listName = schema.listName || this.props.listName;
    var filterStr = filterExpressions.map(function(expression: string) {
      return '(' + expression + ')';
    }).join(' and ');
    console.log('[ReportFormsWebPart] EXECUTING FILTER. List: ' + listName + ', Filter: ' + filterStr);

    var endpoint = this.getWebUrl()
      + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items?$filter="
      + encodeURIComponent(filterStr) + '&$select=Id&$top=1';

    console.log('[ReportFormsWebPart] REST CALL: ' + endpoint);

    try {
      var response = await this.getWithAcceptFallback(endpoint);
      console.log('[ReportFormsWebPart] REST RESPONSE: Status ' + response.status);
      
      if (!response.ok) {
        console.error('[ReportFormsWebPart] REST ERROR: Status ' + response.status);
        var errorText = '';
        try {
          errorText = await response.text();
          console.error('[ReportFormsWebPart] Response body: ' + errorText);
        } catch (_textError) {
          console.error('[ReportFormsWebPart] Could not read error response');
        }
        return this.resolveLookupNavigationFallback(listName, filterStr);
      }
      
      var data = await response.json();
      var items = (data && data.value) ? data.value : [];
      if (!items || items.length === 0) {
        items = data && data.d && data.d.results ? data.d.results : [];
      }
      console.log('[ReportFormsWebPart] FILTER RESULTS: ' + items.length + ' item(s) found');
      
      if (items.length > 0 && items[0].Id) {
        var resolvedId = toPositiveInt(items[0].Id);
        console.log('[ReportFormsWebPart] FILTER SUCCESS: Resolved to item ID ' + resolvedId);
        return resolvedId;
      } else {
        console.log('[ReportFormsWebPart] FILTER NO MATCH: Filter executed but no items found');
        return this.resolveLookupNavigationFallback(listName, filterStr);
      }
    } catch (filterError) {
      console.error('[ReportFormsWebPart] FILTER ERROR: ', filterError);
    }
    console.log('[ReportFormsWebPart] FILTER RESOLUTION FAILED: Returning 0');
    return 0;
  }

  private async resolveLookupNavigationFallback(listName: string, filterStr: string): Promise<number> {
    var fallbackFilter = filterStr;
    var expandFields: string[] = [];
    var fieldKey: string;
    for (fieldKey in this._listFieldTypeLookup) {
      if (!Object.prototype.hasOwnProperty.call(this._listFieldTypeLookup, fieldKey)) { continue; }
      var fieldType = this._listFieldTypeLookup[fieldKey];
      if (fieldType !== 'lookup' && fieldType !== 'lookupmulti') { continue; }
      var internalName = this._listFieldInternalNameLookup[fieldKey];
      if (!internalName) { continue; }
      var escapedName = internalName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      var idFilterPattern = new RegExp('(^|[\\s(])' + escapedName + 'Id(?=\\s+(?:eq|ne|gt|ge|lt|le)\\s+)', 'g');
      var nextFilter = fallbackFilter.replace(idFilterPattern, '$1' + internalName + '/Id');
      if (nextFilter !== fallbackFilter) {
        fallbackFilter = nextFilter;
        if (expandFields.indexOf(internalName) < 0) { expandFields.push(internalName); }
      }
    }

    if (expandFields.length === 0 || fallbackFilter === filterStr) { return 0; }

    var fallbackEndpoint = this.getWebUrl()
      + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items?$filter="
      + encodeURIComponent(fallbackFilter)
      + '&$select=Id&$expand=' + encodeURIComponent(expandFields.join(',')) + '&$top=1';
    console.log('[ReportFormsWebPart] RETRYING LOOKUP FILTER: ' + fallbackFilter);

    try {
      var response = await this.getWithAcceptFallback(fallbackEndpoint);
      if (!response.ok) {
        console.error('[ReportFormsWebPart] LOOKUP FILTER RETRY ERROR: Status ' + response.status);
        return 0;
      }
      var data = await response.json();
      var items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      if (items && items.length > 0 && items[0].Id) {
        var resolvedId = toPositiveInt(items[0].Id);
        console.log('[ReportFormsWebPart] LOOKUP FILTER RETRY SUCCESS: Resolved to item ID ' + resolvedId);
        return resolvedId;
      }
    } catch (fallbackError) {
      console.error('[ReportFormsWebPart] LOOKUP FILTER RETRY FAILED: ', fallbackError);
    }
    return 0;
  }

  private async loadLookupOptions(lookupList: string, lookupField: string): Promise<ILookupOptionItem[]> {
    console.log('[ReportFormsWebPart] loadLookupOptions START - List: ' + lookupList + ', Field: ' + lookupField);
    
    var selector = encodeURIComponent('Id,' + lookupField);
    
    // Strip curly braces from GUID if present
    var cleanLookupList = lookupList.replace(/^\{|\}$/g, '');
    
    var endpoint = isGuid(lookupList)
      ? this.getWebUrl() + "/_api/web/lists(guid'" + cleanLookupList + "')/items?$select=" + selector + '&$top=5000'
      : this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(lookupList) + "')/items?$select=" + selector + '&$top=5000';
    
    console.log('[ReportFormsWebPart] Is GUID? ' + isGuid(lookupList) + ', Clean List: ' + cleanLookupList);
    console.log('[ReportFormsWebPart] REST Endpoint: ' + endpoint);
    
    try {
      var response = await this.getWithAcceptFallback(endpoint);
      
      console.log('[ReportFormsWebPart] REST Response Status: ' + response.status + ' (' + response.statusText + ')');
      
      if (!response.ok) {
        console.error('[ReportFormsWebPart] REST call failed! Status: ' + response.status);
        return [];
      }
      
      var data = await response.json();
      var items = data && data.value ? data.value : [];
      if (!items || items.length === 0) {
        items = data && data.d && data.d.results ? data.d.results : [];
      }
      console.log('[ReportFormsWebPart] REST returned ' + items.length + ' items from lookup list');
      
      var mappedItems = items.map(function(item: any) {
        return {
          Id: item.Id,
          Title: item[lookupField] || item.Title || String(item.Id),
        };
      });
      
      console.log('[ReportFormsWebPart] loadLookupOptions SUCCESS - Returning ' + mappedItems.length + ' items');
      return mappedItems;
    } catch (error) {
      console.error('[ReportFormsWebPart] loadLookupOptions ERROR: ', error);
      return [];
    }
  }

  private async loadPersonOptions(): Promise<ILookupOptionItem[]> {
    var endpoint = this.getWebUrl() + '/_api/web/siteusers?$select=Id,Title,Email,PrincipalType&$top=5000';

    try {
      var response = await this.getWithAcceptFallback(endpoint);
      if (!response.ok) {
        this.logDiagnostic('Failed to load person options. Status: ' + response.status + ' (' + response.statusText + ')');
        return [];
      }

      var data = await response.json();
      var users = data && data.value ? data.value : [];
      if (!users || users.length === 0) {
        users = data && data.d && data.d.results ? data.d.results : [];
      }

      return users
        .filter(function(user: any) {
          return user && user.Id > 0 && user.PrincipalType !== undefined && ((user.PrincipalType & 1) === 1);
        })
        .map(function(user: any) {
          return {
            Id: user.Id,
            Title: user.Title || user.Email || String(user.Id),
            Email: user.Email || undefined,
          };
        });
    } catch (error) {
      this.logDiagnostic('Failed to load person options: ' + (error && error.message ? error.message : 'Unknown error'));
      return [];
    }
  }

  private async logSubmitFailure(response: any, action: string, listName: string, payload: any): Promise<void> {
    var details = '';
    try {
      details = await response.text();
    } catch (_error) {
      details = '';
    }

    this.logDiagnostic(
      action + ' failed for list "' + listName + '". Status=' + response.status + ' ' + response.statusText
      + '; Payload=' + JSON.stringify(payload)
      + (details ? '; Response=' + details : '')
    );
  }

  private getPendingAttachments(fieldId: string): File[] {
    var value = this.state.values[fieldId];
    return Array.isArray(value) ? value as File[] : [];
  }

  private getAttachmentsMarkedForDelete(fieldId: string): IAttachmentItem[] {
    return this._attachmentsMarkedForDelete[fieldId] || [];
  }

  private isAttachmentDeleteAllowed(field: FormField): boolean {
    return !field.config || field.config.allowAttachmentDelete !== false;
  }

  private forceAttachmentFieldRefresh(fieldId: string): void {
    var values: IValueMap = {};
    var key: string;
    for (key in this.state.values) {
      if (Object.prototype.hasOwnProperty.call(this.state.values, key)) {
        values[key] = this.state.values[key];
      }
    }
    this.setState({ values: values });
  }

  private toggleAttachmentMarkedForDelete(fieldId: string, attachment: IAttachmentItem): void {
    var current = this.getAttachmentsMarkedForDelete(fieldId).slice(0);
    var existingIndex = -1;
    for (var i = 0; i < current.length; i += 1) {
      if (current[i].fileName === attachment.fileName) {
        existingIndex = i;
        break;
      }
    }

    if (existingIndex >= 0) {
      current.splice(existingIndex, 1);
    } else {
      current.push(attachment);
    }

    this._attachmentsMarkedForDelete[fieldId] = current;
    this.forceAttachmentFieldRefresh(fieldId);
  }

  private isAttachmentMarkedForDelete(fieldId: string, attachment: IAttachmentItem): boolean {
    var current = this.getAttachmentsMarkedForDelete(fieldId);
    for (var i = 0; i < current.length; i += 1) {
      if (current[i].fileName === attachment.fileName) {
        return true;
      }
    }
    return false;
  }

  private clearAttachmentInput(fieldId: string): void {
    var input = this._attachmentInputRefs[fieldId];
    if (input) {
      input.value = '';
    }
  }

  private buildAttachmentFailureMessage(fileName: string, errorMessage: string): string {
    this.logDiagnostic(
      this.buildAttachmentFailureMessageInternal(fileName, errorMessage)
    );
    return strings.FormSubmitFailedDefault;
  }

  private buildAttachmentFailureMessageInternal(fileName: string, errorMessage: string): string {
    return strings.UploadAttachmentFailedTemplate
      .replace('{0}', fileName)
      .replace('{1}', errorMessage || strings.FormSubmitFailedDefault);
  }

  private async attachmentExists(listName: string, itemId: number, fileName: string): Promise<boolean> {
    var endpoint = this.getWebUrl()
      + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items(" + itemId + ")/AttachmentFiles/getByFileName('"
      + encodeURIComponent(fileName).replace(/'/g, '%27')
      + "')";

    try {
      var response = await this.getWithAcceptFallback(endpoint);
      return !!response.ok;
    } catch (_error) {
      return false;
    }
  }

  private async attachmentExistsFromCollection(listName: string, itemId: number, fileName: string): Promise<boolean> {
    var endpoint = this.getWebUrl()
      + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items(" + itemId + ")?$select=AttachmentFiles/FileName&$expand=AttachmentFiles";

    try {
      var response = await this.getWithAcceptFallback(endpoint);
      if (!response.ok) {
        return false;
      }

      var payload = await response.json();
      var d = payload && payload.d ? payload.d : payload;
      var raw = d && d.AttachmentFiles;
      var items = raw && raw.results ? raw.results : raw;
      if (!Array.isArray(items)) {
        return false;
      }

      for (var i = 0; i < items.length; i += 1) {
        if (items[i] && items[i].FileName === fileName) {
          return true;
        }
      }
    } catch (_error) {
      return false;
    }

    return false;
  }

  private async waitForAttachmentVisibility(listName: string, itemId: number, fileName: string): Promise<boolean> {
    if (await this.attachmentExists(listName, itemId, fileName)) {
      return true;
    }

    // SharePoint can return non-OK while the file is still being committed.
    for (var attempt = 0; attempt < 3; attempt += 1) {
      await new Promise(function(resolve) { return setTimeout(resolve, 350 * (attempt + 1)); });
      if (await this.attachmentExists(listName, itemId, fileName)) {
        return true;
      }
    }

    return this.attachmentExistsFromCollection(listName, itemId, fileName);
  }

  private async deleteMarkedAttachments(listName: string, itemId: number, schema: FormSchema): Promise<void> {
    var fields = getAllFields(schema);
    for (var i = 0; i < fields.length; i += 1) {
      var field = fields[i];
      if (field.type !== 'attachment') {
        continue;
      }

      var marked = this.getAttachmentsMarkedForDelete(field.id).slice(0);
      if (!marked.length || !this.isAttachmentDeleteAllowed(field)) {
        continue;
      }

      for (var j = 0; j < marked.length; j += 1) {
        var attachment = marked[j];

        var existsBeforeDelete = await this.attachmentExistsFromCollection(listName, itemId, attachment.fileName);
        if (!existsBeforeDelete) {
          this.logDiagnostic('Attachment marked for delete is already missing. Skipping delete call. file=' + attachment.fileName + ', itemId=' + String(itemId));
          continue;
        }

        var endpoint = this.getWebUrl()
          + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items(" + itemId + ")/AttachmentFiles/getByFileName('"
          + encodeURIComponent(attachment.fileName).replace(/'/g, '%27')
          + "')";

        var response = await this.postRawWithAcceptFallback(endpoint, '', {
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE',
        });

        if (!response.ok) {
          var existsAfterDeleteError = await this.attachmentExistsFromCollection(listName, itemId, attachment.fileName);
          if (!existsAfterDeleteError) {
            this.logDiagnostic('Delete attachment returned non-OK (' + String(response.status) + ') but file no longer exists. Treating as success. file=' + attachment.fileName + ', itemId=' + String(itemId));
            continue;
          }

          await this.logSubmitFailure(response, 'Delete attachment', listName, { fieldId: field.id, fileName: attachment.fileName, itemId: itemId });
          throw new Error(strings.FormSubmitFailedDefault);
        }
      }

      var remainingAttachments = this._existingAttachments[field.id] || [];
      this._existingAttachments[field.id] = remainingAttachments.filter(function(item) {
        for (var index = 0; index < marked.length; index += 1) {
          if (marked[index].fileName === item.fileName) {
            return false;
          }
        }
        return true;
      });
      this._attachmentsMarkedForDelete[field.id] = [];
    }
  }

  private async uploadAttachments(listName: string, itemId: number, schema: FormSchema): Promise<void> {
    var fields = getAllFields(schema);
    for (var i = 0; i < fields.length; i += 1) {
      var field = fields[i];
      if (field.type !== 'attachment') {
        continue;
      }

      var files = this.getPendingAttachments(field.id);
      if (!files.length) {
        continue;
      }

      for (var j = 0; j < files.length; j += 1) {
        var file = files[j];
        var endpoint = this.getWebUrl()
          + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items(" + itemId + ")/AttachmentFiles/add(FileName='"
          + encodeURIComponent(file.name).replace(/'/g, '%27')
          + "')";

        var response = await this.postRawWithAcceptFallback(endpoint, file, {
          'Content-Type': 'application/octet-stream',
        });

        if (!response.ok) {
          var filePersisted = await this.waitForAttachmentVisibility(listName, itemId, file.name);
          if (filePersisted) {
            this.logDiagnostic('Attachment upload returned non-OK (' + String(response.status) + ') but file exists after verification. Treating as success. file=' + file.name + ', itemId=' + String(itemId));
          } else {
            await this.logSubmitFailure(response, 'Upload attachment', listName, { fieldId: field.id, fileName: file.name, itemId: itemId });
            throw new Error(this.buildAttachmentFailureMessage(file.name, response.statusText || strings.FormSubmitFailedDefault));
          }
        }

        var existing = this._existingAttachments[field.id] || [];
        var hasExisting = false;
        for (var existingIndex = 0; existingIndex < existing.length; existingIndex += 1) {
          if (existing[existingIndex].fileName === file.name) {
            hasExisting = true;
            break;
          }
        }
        if (!hasExisting) {
          existing.push({ fileName: file.name });
          this._existingAttachments[field.id] = existing;
        }
      }

      this.clearAttachmentInput(field.id);
    }
  }

  private getCreatedItemId(responseData: any): number {
    if (responseData && typeof responseData.Id === 'number') {
      return responseData.Id;
    }
    if (responseData && responseData.d && typeof responseData.d.Id === 'number') {
      return responseData.d.Id;
    }
    return 0;
  }

  private async loadFieldDescriptions(schema: FormSchema): Promise<void> {
    if (!schema || !schema.listName) {
      this._listFieldNameLookup = {};
      this._listFieldInternalNameLookup = {};
      this._listFieldTypeLookup = {};
      return;
    }

    try {
      console.log('[ReportFormsWebPart] loadFieldDescriptions START - List: ' + schema.listName);
      
      // Get all fields from the SharePoint list
      var response = await this.getWithAcceptFallback(
        this.props.context.pageContext.web.absoluteUrl + '/_api/web/lists/getByTitle(\'' + escapeODataText(schema.listName) + '\')/fields?$filter=Hidden eq false'
      );

      if (!response.ok) {
        console.warn('[ReportFormsWebPart] loadFieldDescriptions: Failed to fetch fields, status ' + response.status);
        return;
      }

      var data = await response.json();
      var fields = data && data.value ? data.value : [];
      if ((!fields || !Array.isArray(fields) || fields.length === 0) && data && data.d && data.d.results) {
        fields = data.d.results;
      }

      if (!fields || !Array.isArray(fields)) {
        console.warn('[ReportFormsWebPart] loadFieldDescriptions: Invalid response format');
        return;
      }

      // Build a map of field name -> metadata
      var fieldMetaMap: { [key: string]: { description?: string; richText?: boolean } } = {};
      var fieldNameLookup: { [lowerInternalName: string]: boolean } = {};
      var fieldInternalNameLookup: { [lowerInternalName: string]: string } = {};
      var fieldTypeLookup: { [lowerInternalName: string]: string } = {};
      fields.forEach((field: any) => {
        if (field.InternalName) {
          var lowerInternalName = String(field.InternalName).toLowerCase();
          fieldNameLookup[lowerInternalName] = true;
          fieldInternalNameLookup[lowerInternalName] = String(field.InternalName);
          fieldTypeLookup[lowerInternalName] = String(field.TypeAsString || '').toLowerCase();
          fieldMetaMap[field.InternalName] = {
            description: field.Description || undefined,
            richText: field.TypeAsString === 'Note' && field.RichText === true,
          };
          if (field.Description) {
            console.log('[ReportFormsWebPart] Found description for "' + field.InternalName + '": "' + field.Description + '"');
          }
        }
      });

      this._listFieldNameLookup = fieldNameLookup;
  this._listFieldInternalNameLookup = fieldInternalNameLookup;
  this._listFieldTypeLookup = fieldTypeLookup;

      // Update schema fields with descriptions and note-field rich text metadata.
      schema.steps.forEach((step) => {
        if (step.fields) {
          step.fields.forEach((field) => {
            var fieldMeta = fieldMetaMap[field.fieldName];
            if (!fieldMeta) {
              return;
            }

            if (fieldMeta.description && !field.description) {
              field.description = fieldMeta.description;
              console.log('[ReportFormsWebPart] Populated description for field "' + field.fieldName + '": "' + field.description + '"');
            }

            if (field.type === 'multiline' && fieldMeta.richText === true) {
              field.config = field.config || {};
              if (field.config.isRichText !== true) {
                field.config.isRichText = true;
                console.log('[ReportFormsWebPart] Marked field "' + field.fieldName + '" as rich text note field');
              }
            }
          });
        }
      });

      console.log('[ReportFormsWebPart] loadFieldDescriptions SUCCESS');
    } catch (error) {
      this._listFieldNameLookup = {};
      this._listFieldInternalNameLookup = {};
      this._listFieldTypeLookup = {};
      console.error('[ReportFormsWebPart] loadFieldDescriptions ERROR: ', error);
    }
  }

  private mapItemToValues(schema: FormSchema, item: any, currentValues: IValueMap): IValueMap {
    var values: IValueMap = {};
    var currentKey: string;
    for (currentKey in currentValues) {
      if (Object.prototype.hasOwnProperty.call(currentValues, currentKey)) {
        values[currentKey] = currentValues[currentKey];
      }
    }

    var fields = getAllFields(schema);
    for (var i = 0; i < fields.length; i += 1) {
      var field = fields[i];
      if (field.type === 'attachment') {
        var rawAttachments = item.AttachmentFiles;
        var attachmentItems = rawAttachments && rawAttachments.results ? rawAttachments.results : rawAttachments;
        if (attachmentItems && attachmentItems.length) {
          this._existingAttachments[field.id] = attachmentItems.map(function(attachment: any) {
            return {
              fileName: attachment.FileName || '',
              serverRelativeUrl: attachment.ServerRelativeUrl || undefined,
            } as IAttachmentItem;
          });
        } else {
          this._existingAttachments[field.id] = [];
        }
        this._attachmentsMarkedForDelete[field.id] = [];
        values[field.id] = [];
        continue;
      }

      var rawValue = item[field.fieldName];
      var rawIdValue = item[field.fieldName + 'Id'];
      if ((rawValue === undefined || rawValue === null) && rawIdValue !== undefined && rawIdValue !== null) {
        rawValue = rawIdValue;
      }
      if (rawValue === undefined || rawValue === null) {
        // Explicitly reset to a blank value for this field's type instead of skipping, so a blank
        // field on the newly loaded record does not keep showing the previous record's value.
        values[field.id] = getBlankValueForField(field);
        if (field.type === 'url') {
          this._urlFieldDescriptions[field.id] = '';
        }
        continue;
      }

      if (field.type === 'multiselect') {
        if (rawValue.results && rawValue.results.length) {
          values[field.id] = rawValue.results;
        } else if (rawValue.length && typeof rawValue !== 'string') {
          values[field.id] = rawValue;
        } else {
          values[field.id] = [];
        }
      } else if (field.type === 'lookup' && field.config && field.config.allowMultiple === true) {
        if (rawValue.results && rawValue.results.length) {
          values[field.id] = rawValue.results.map(function(result: any) {
            return String(result && result.Id !== undefined ? result.Id : result);
          });
        } else if (rawValue.length && typeof rawValue !== 'string') {
          values[field.id] = rawValue.map(function(result: any) {
            return String(result && result.Id !== undefined ? result.Id : result);
          });
        } else {
          values[field.id] = [];
        }
      } else if (field.type === 'person' && field.config && field.config.allowMultiple === true) {
        if (rawValue.results && rawValue.results.length) {
          values[field.id] = rawValue.results.map(function(result: any) {
            return String(result && result.Id !== undefined ? result.Id : result);
          });
        } else if (rawValue.length && typeof rawValue !== 'string') {
          values[field.id] = rawValue.map(function(result: any) {
            return String(result && result.Id !== undefined ? result.Id : result);
          });
        } else {
          values[field.id] = [];
        }
      } else if (field.type === 'boolean') {
        values[field.id] = rawValue === true;
      } else if (field.type === 'lookup' && typeof rawValue === 'object' && rawValue !== null && rawValue.Id) {
        values[field.id] = String(rawValue.Id);
      } else if (field.type === 'lookup') {
        values[field.id] = '';
      } else if (field.type === 'person' && typeof rawValue === 'object' && rawValue !== null && rawValue.Id) {
        values[field.id] = String(rawValue.Id);
      } else if (field.type === 'person' && (typeof rawValue === 'number' || (typeof rawValue === 'string' && rawValue !== ''))) {
        values[field.id] = String(rawValue);
      } else if (field.type === 'person') {
        values[field.id] = '';
      } else if (field.type === 'url' && typeof rawValue === 'object' && rawValue !== null && rawValue.Url) {
        values[field.id] = rawValue.Url;
        this._urlFieldDescriptions[field.id] = rawValue.Description || '';
      } else if (field.type === 'url') {
        values[field.id] = '';
        this._urlFieldDescriptions[field.id] = '';
      } else if (field.type === 'datetime') {
        var loadedTimeZone = field.config && field.config.timeZone;
        var loadedFormat = field.config && field.config.displayFormat;
        values[field.id] = loadedFormat === 'dateOnly'
          ? normalizeDateOnlyValue(rawValue, loadedTimeZone)
          : loadedFormat === 'timeOnly'
            ? normalizeTimeOnlyValue(rawValue, loadedTimeZone)
            : normalizeDateTimeLocalValue(rawValue, loadedTimeZone);
        console.log('[ReportFormsWebPart] Loaded datetime field "' + field.fieldName + '" (format=' + (loadedFormat || 'dateTime') + ', zone=' + (loadedTimeZone || 'UTC') + '): ' + values[field.id]);
      } else {
        values[field.id] = rawValue;
      }
    }

    return values;
  }

  private setFieldValue(fieldId: string, value: FieldValue): void {
    var values: IValueMap = {};
    var key: string;
    for (key in this.state.values) {
      if (Object.prototype.hasOwnProperty.call(this.state.values, key)) {
        values[key] = this.state.values[key];
      }
    }
    values[fieldId] = value;

    var fieldErrors: IFieldErrorMap = {};
    for (key in this.state.fieldErrors) {
      if (Object.prototype.hasOwnProperty.call(this.state.fieldErrors, key) && key !== fieldId) {
        fieldErrors[key] = this.state.fieldErrors[key];
      }
    }

    this.setState({ values: values, fieldErrors: fieldErrors });
  }

  private resolveFieldIdByReference(schema: FormSchema, fieldRef: string): string {
    var normalized = String(fieldRef || '').trim().toLowerCase();
    if (!normalized) {
      return '';
    }

    var fields = getAllFields(schema);
    for (var i = 0; i < fields.length; i += 1) {
      var f = fields[i];
      if ((f.id && f.id.toLowerCase() === normalized) || (f.fieldName && f.fieldName.toLowerCase() === normalized) || (f.label && f.label.toLowerCase() === normalized)) {
        return f.id;
      }
    }

    return '';
  }

  private validateAdvancedRules(schema: FormSchema, errors: IFieldErrorMap): string[] {
    var formErrors: string[] = [];
    var config = schema.advancedValidation;
    if (!config || config.enabled !== true || !config.rules || config.rules.length === 0) {
      return formErrors;
    }

    var fields = getAllFields(schema);
    for (var i = 0; i < config.rules.length; i += 1) {
      var rule = config.rules[i] as AdvancedValidationRule;
      if (!rule || !rule.expression || !rule.message) {
        continue;
      }

      try {
        var valid = !!evaluateAdvancedExpression(rule.expression, fields, this.state.values);
        if (valid) {
          continue;
        }

        var targetFieldId = rule.targetField ? this.resolveFieldIdByReference(schema, rule.targetField) : '';
        if (targetFieldId) {
          if (!errors[targetFieldId]) {
            errors[targetFieldId] = rule.message;
          }
        } else {
          formErrors.push(rule.message);
        }
      } catch (error) {
        formErrors.push('Validation rule failed: ' + (error && error.message ? error.message : 'Unknown error.'));
      }
    }

    return formErrors;
  }

  private validateForm(schema: FormSchema): { fieldErrors: IFieldErrorMap; formErrors: string[] } {
    var fields = getAllFields(schema);
    var errors: IFieldErrorMap = {};
    for (var i = 0; i < fields.length; i += 1) {
      var field = fields[i];
      if (field.required !== true) {
        continue;
      }
      var value = this.state.values[field.id];
      var isEmpty = value === null || value === undefined || value === '' || (Array.isArray(value) && value.length === 0);
      if (field.type === 'attachment') {
        var existingAttachments = this._existingAttachments[field.id] || [];
        var markedForDelete = this.getAttachmentsMarkedForDelete(field.id);
        var pendingAttachments = Array.isArray(value) ? value : [];
        isEmpty = (existingAttachments.length - markedForDelete.length) <= 0 && pendingAttachments.length === 0;
      }
      if (isEmpty) {
        errors[field.id] = field.requiredMessage || (field.label + ' is required');
      }
    }

    var formErrors = this.validateAdvancedRules(schema, errors);
    return { fieldErrors: errors, formErrors: formErrors };
  }

  private buildSubmitPayload(schema: FormSchema): any {
    var payload: any = {};
    var fields = getAllFields(schema);
    var hasListFieldLookup = Object.keys(this._listFieldNameLookup).length > 0;
    for (var i = 0; i < fields.length; i += 1) {
      var field = fields[i];
      var value = this.state.values[field.id];

      // Display-only and unsupported field types are not persisted to SharePoint items.
      if (
        field.type === 'richtext'
        || field.type === 'customimage'
        || field.type === 'divider'
        || field.type === 'newline'
        || field.type === 'attachment'
        || field.type === 'taxonomy'
        || field.type === 'image'
      ) {
        continue;
      }

      // Submit only controls directly bound to real SharePoint columns.
      if (!field.fieldName) {
        continue;
      }
      var normalizedFieldName = String(field.fieldName).trim().toLowerCase();
      if (!normalizedFieldName) {
        continue;
      }
      if (hasListFieldLookup && !this._listFieldNameLookup[normalizedFieldName]) {
        continue;
      }

      if (value === undefined) {
        continue;
      }

      if (field.type === 'lookup') {
        if (field.config && field.config.allowMultiple === true) {
          var lookupValues = (Array.isArray(value) ? value : []) as any[];
          payload[field.fieldName + 'Id'] = {
            __metadata: { type: 'Collection(Edm.Int32)' },
            results: lookupValues
              .map(function(item) { return parseInt(String(item), 10); })
              .filter(function(item) { return !isNaN(item); }),
          };
        } else {
          payload[field.fieldName + 'Id'] = value === '' || value === null ? null : parseInt(String(value), 10);
        }
      } else if (field.type === 'person') {
        if (field.config && field.config.allowMultiple === true) {
          var personValues = (Array.isArray(value) ? value : []) as any[];
          payload[field.fieldName + 'Id'] = {
            __metadata: { type: 'Collection(Edm.Int32)' },
            results: personValues
              .map(function(item) { return parseInt(String(item), 10); })
              .filter(function(item) { return !isNaN(item); }),
          };
        } else {
          payload[field.fieldName + 'Id'] = value === '' || value === null ? null : parseInt(String(value), 10);
        }
      } else if (field.type === 'multiselect') {
        payload[field.fieldName] = Array.isArray(value) ? value : [];
      } else if (field.type === 'boolean') {
        payload[field.fieldName] = value === true;
      } else if (field.type === 'number') {
        payload[field.fieldName] = value === '' ? null : Number(value);
      } else if (field.type === 'datetime') {
        payload[field.fieldName] = buildDateTimePayloadValue(value, field.config);
        console.log('[ReportFormsWebPart] Prepared datetime field "' + field.fieldName + '" (format=' + (field.config && field.config.displayFormat || 'dateTime') + ', zone=' + (field.config && field.config.timeZone || 'UTC') + ') payload value: ' + payload[field.fieldName]);
      } else if (field.type === 'url') {
        payload[field.fieldName] = value ? { Url: String(value), Description: this._urlFieldDescriptions[field.id] || '' } : null;
      } else {
        payload[field.fieldName] = value;
      }
    }

    var linkedTargetLower = String(this.props.linkedFieldTarget || '').trim().toLowerCase();
    var linkedTargetInternalName = this._listFieldInternalNameLookup[linkedTargetLower];
    var linkedTargetType = this._listFieldTypeLookup[linkedTargetLower];
    var linkedLookupId = getLookupIdValue(this.props.linkedFieldValue);
    if (this.props.useDynamicValueAsFilter && linkedTargetInternalName && linkedLookupId
      && (linkedTargetType === 'lookup' || linkedTargetType === 'lookupmulti')) {
      if (linkedTargetType === 'lookupmulti') {
        payload[linkedTargetInternalName + 'Id'] = {
          __metadata: { type: 'Collection(Edm.Int32)' },
          results: [parseInt(linkedLookupId, 10)]
        };
      } else {
        payload[linkedTargetInternalName + 'Id'] = parseInt(linkedLookupId, 10);
      }
      console.log('[ReportFormsWebPart] Applied dynamic parent lookup payload: ' + linkedTargetInternalName + 'Id=' + linkedLookupId);
    }
    return payload;
  }

  private async createItem(listName: string, payload: any): Promise<number> {
    console.log('[ReportFormsWebPart] createItem START - List: ' + listName + ', Payload: ', payload);
    var response = await this.postWithAcceptFallback(
      this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items",
      payload
    );
    if (!response.ok) {
      await this.logSubmitFailure(response, 'Create item', listName, payload);
      throw new Error(strings.FormSubmitFailedDefault);
    }
    var createdItemId = this.getCreatedItemId(await response.json());
    console.log('[ReportFormsWebPart] createItem SUCCESS - New item ID: ' + createdItemId);
    return createdItemId;
  }

  private async updateItem(listName: string, itemId: number, payload: any): Promise<void> {
    console.log('[ReportFormsWebPart] updateItem START - List: ' + listName + ', ItemId: ' + itemId + ', Payload: ', payload);
    var response = await this.postWithAcceptFallback(
      this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items(" + itemId + ')',
      payload,
      {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      }
    );
    if (!response.ok) {
      await this.logSubmitFailure(response, 'Update item', listName, payload);
      throw new Error(strings.FormSubmitFailedDefault);
    }
    console.log('[ReportFormsWebPart] updateItem SUCCESS - ItemId: ' + itemId);
  }

  private async handleSubmit(): Promise<void> {
    var schema = this.state.schema;
    if (!schema) {
      return;
    }
    var effectiveMode = this.getEffectiveMode();
    console.log('[ReportFormsWebPart] handleSubmit START - Mode: ' + effectiveMode + ', List: ' + (schema.listName || this.props.listName));

    var canSubmitByPermission = effectiveMode === 'edit' ? this.state.canEditRecords : this.state.canAddRecords;
    if (!canSubmitByPermission) {
      var deniedMessage = 'Permission gate blocked submit. Mode=' + effectiveMode + ', canAdd=' + this.state.canAddRecords + ', canEdit=' + this.state.canEditRecords;
      this.logDiagnostic(deniedMessage);
      this.setState({
        submitError: this.getPermissionDeniedMessage(schema),
        submitSuccess: null,
      });
      return;
    }

    var validationResult = this.validateForm(schema);
    if (Object.keys(validationResult.fieldErrors).length > 0 || validationResult.formErrors.length > 0) {
      var submitValidationMessage = strings.FormStepValidationError;
      if (validationResult.formErrors.length > 0) {
        submitValidationMessage += ' ' + validationResult.formErrors.join(' ');
      }
      this.setState({ fieldErrors: validationResult.fieldErrors, submitError: submitValidationMessage, submitSuccess: null });
      return;
    }

    this.setState({ isSubmitting: true, submitError: null, submitSuccess: null, fieldErrors: {} });

    try {
      var payload = this.buildSubmitPayload(schema);
      var listName = schema.listName || this.props.listName;
      var itemId = this.getResolvedItemId(schema);
      var wasCreate = !(effectiveMode === 'edit' && itemId > 0);
      if (effectiveMode === 'edit' && itemId > 0) {
        await this.updateItem(listName, itemId, payload);
      } else {
        itemId = await this.createItem(listName, payload);
      }

      if (itemId <= 0) {
        this.logDiagnostic('Item save completed but no item ID was available for follow-up attachment upload.');
      }

      if (itemId > 0) {
        await this.deleteMarkedAttachments(listName, itemId, schema);
        await this.uploadAttachments(listName, itemId, schema);
      }

      var nextValues: IValueMap = {};
      var key: string;
      for (key in this.state.values) {
        if (Object.prototype.hasOwnProperty.call(this.state.values, key)) {
          nextValues[key] = this.state.values[key];
        }
      }
      var attachmentFields = getAllFields(schema);
      for (var attachmentIndex = 0; attachmentIndex < attachmentFields.length; attachmentIndex += 1) {
        if (attachmentFields[attachmentIndex].type === 'attachment') {
          nextValues[attachmentFields[attachmentIndex].id] = [];
          this._attachmentsMarkedForDelete[attachmentFields[attachmentIndex].id] = [];
        }
      }

      this.notifyListControlRefresh(listName, wasCreate ? 'add' : 'edit');
      console.log('[ReportFormsWebPart] handleSubmit SUCCESS - ' + (wasCreate ? 'Created' : 'Updated') + ' item ID: ' + itemId);

      this.setState({
        values: nextValues,
        isSubmitting: false,
        isViewEditing: this.props.mode === 'view' ? false : this.state.isViewEditing,
        submitError: null,
        submitSuccess: this.props.onSubmitMessage || (effectiveMode === 'edit' ? strings.FormSubmitSuccessEdit : strings.FormSubmitSuccessNew),
      });

      var submitRedirectTarget = this.getPreferredRedirectUrl(this.props.submitRedirectUrl);
      if (submitRedirectTarget && typeof window !== 'undefined') {
        window.setTimeout(() => {
          if (!this.tryRedirect(submitRedirectTarget)) {
            this.setSafeState({
              submitError: strings.FormRedirectFailed
            });
          }
        }, 1000);
      }
    } catch (error) {
      console.error('[ReportFormsWebPart] handleSubmit ERROR: ', error);
      this.setState({
        isSubmitting: false,
        submitError: error && error.message ? error.message : strings.FormSubmitFailedDefault,
        submitSuccess: null,
      });
    }
  }

  private notifyListControlRefresh(listName: string, reason: string): void {
    if (typeof window === 'undefined' || !(window as any).dispatchEvent) {
      return;
    }

    try {
      var event: any;
      if (typeof CustomEvent === 'function') {
        event = new CustomEvent(LIST_CONTROL_REFRESH_EVENT, {
          detail: {
            listName: listName,
            reason: reason
          }
        });
      } else {
        event = document.createEvent('CustomEvent');
        event.initCustomEvent(LIST_CONTROL_REFRESH_EVENT, false, false, {
          listName: listName,
          reason: reason
        });
      }

      window.dispatchEvent(event);
    } catch (eventError) {
      console.warn('[ReportFormsWebPart] notifyListControlRefresh: Failed to dispatch refresh event for list "' + listName + '": ', eventError);
    }
  }

  private handleCancel(): void {
    console.log('[ReportFormsWebPart] handleCancel invoked. Mode: ' + this.props.mode + ', isViewEditing: ' + this.state.isViewEditing);
    var cancelRedirectTarget = this.getPreferredRedirectUrl(this.props.cancelRedirectUrl);
    if (cancelRedirectTarget && typeof window !== 'undefined') {
      if (!this.tryRedirect(cancelRedirectTarget)) {
        this.setState({ submitError: strings.FormRedirectFailed, submitSuccess: null });
      }
      return;
    }

    if (this.props.mode === 'view' && this.state.isViewEditing) {
      this.setState({ isViewEditing: false, fieldErrors: {}, submitError: null, submitSuccess: null });
    }

    this.loadData();
  }

  private handleDesignerChange(schema: FormSchema): void {
    this.setState({ schema: schema });
  }

  private handleDesignerSave(): void {
    if (!this.state.schema) {
      return;
    }

    console.log('[ReportFormsWebPart] handleDesignerSave: Saving schema for list "' + (this.state.schema.listName || this.props.listName) + '"');
    this.props.onSaveSchema(this.state.schema);
    this.props.onToggleDesignerMode();
  }

  private renderUnsupportedField(field: FormField): JSX.Element {
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        {field.label + ': feature not available in SharePoint SE compatibility mode.'}
      </MessageBar>
    );
  }

  private formatReportDateTime(value: any, displayFormat?: string, customFormat?: string, customFormatCase?: string): string {
    var rawValue = value === undefined || value === null ? '' : String(value).trim();
    if (!rawValue) {
      return '-';
    }

    if (displayFormat === 'custom') {
      var customDateTimeValue = new Date(rawValue);
      if (isNaN(customDateTimeValue.getTime())) {
        return rawValue;
      }
      var hours = customDateTimeValue.getHours();
      var twelveHour = hours % 12 || 12;
      var year = String(customDateTimeValue.getFullYear());
      var tokenValues: { [token: string]: string } = {
        YYYY: year,
        YY: year.length > 2 ? year.substring(year.length - 2) : year,
        MMMM: customDateTimeValue.toLocaleString(undefined, { month: 'long' }),
        MMM: customDateTimeValue.toLocaleString(undefined, { month: 'short' }),
        MM: padTwoDigits(customDateTimeValue.getMonth() + 1),
        M: String(customDateTimeValue.getMonth() + 1),
        DD: padTwoDigits(customDateTimeValue.getDate()),
        D: String(customDateTimeValue.getDate()),
        dddd: customDateTimeValue.toLocaleString(undefined, { weekday: 'long' }),
        ddd: customDateTimeValue.toLocaleString(undefined, { weekday: 'short' }),
        HH: padTwoDigits(hours),
        H: String(hours),
        hh: padTwoDigits(twelveHour),
        h: String(twelveHour),
        mm: padTwoDigits(customDateTimeValue.getMinutes()),
        ss: padTwoDigits(customDateTimeValue.getSeconds()),
        tt: hours >= 12 ? 'PM' : 'AM'
      };
      return applyTextCase(
        applyDateFormatPattern(String(customFormat || '').trim() || 'MM/DD/YYYY HH:mm', tokenValues),
        String(customFormatCase || '')
      );
    }

    if (displayFormat === 'timeOnly') {
      var timeParts = rawValue.split(':');
      if (timeParts.length >= 2) {
        var timeValue = new Date(2000, 0, 1, parseInt(timeParts[0], 10), parseInt(timeParts[1], 10), timeParts.length > 2 ? parseInt(timeParts[2], 10) : 0);
        if (!isNaN(timeValue.getTime())) {
          return timeValue.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
        }
      }
      return rawValue;
    }

    if (displayFormat === 'dateOnly') {
      var dateParts = rawValue.substring(0, 10).split('-');
      if (dateParts.length === 3) {
        var dateValue = new Date(parseInt(dateParts[0], 10), parseInt(dateParts[1], 10) - 1, parseInt(dateParts[2], 10));
        if (!isNaN(dateValue.getTime())) {
          return dateValue.toLocaleDateString();
        }
      }
      return rawValue;
    }

    var dateTimeValue = new Date(rawValue);
    return isNaN(dateTimeValue.getTime()) ? rawValue : dateTimeValue.toLocaleString();
  }

  private formatReportSelection(value: any, options: ISelectOption[]): string {
    var selectedValues = Array.isArray(value) ? value : [value];
    var displayValues = selectedValues
      .filter(function(selectedValue) {
        return selectedValue !== undefined && selectedValue !== null && String(selectedValue) !== '';
      })
      .map(function(selectedValue) {
        var selectedKey = String(selectedValue);
        var selectedOption = options.find(function(option) { return String(option.key) === selectedKey; });
        return selectedOption ? selectedOption.text : selectedKey;
      });
    return displayValues.length > 0 ? displayValues.join(', ') : '-';
  }

  private formatAttachmentSize(length?: number): string {
    if (length === undefined || length === null || isNaN(length)) {
      return '-';
    }
    if (length < 1024) {
      return String(length) + ' B';
    }

    var units = ['KB', 'MB', 'GB', 'TB'];
    var size = length / 1024;
    var unitIndex = 0;
    while (size >= 1024 && unitIndex < units.length - 1) {
      size = size / 1024;
      unitIndex += 1;
    }
    return String(Math.round(size * 10) / 10) + ' ' + units[unitIndex];
  }

  private getAttachmentFileType(fileName: string): string {
    var extensionIndex = String(fileName || '').lastIndexOf('.');
    if (extensionIndex < 0 || extensionIndex === fileName.length - 1) {
      return 'File';
    }
    return fileName.substring(extensionIndex + 1).toUpperCase();
  }

  private renderField(field: FormField): JSX.Element | null {
    var value = this.state.values[field.id];
    var effectiveMode = this.getEffectiveMode();
    
    // Check if this field is the linked field target (should always be readonly)
    var linkedFieldTarget = (this.props.linkedFieldTarget || '').trim().toLowerCase();
    var isLinkedField = linkedFieldTarget && (
      (field.id && field.id.toLowerCase() === linkedFieldTarget) ||
      (field.fieldName && field.fieldName.toLowerCase() === linkedFieldTarget)
    );
    
    var disabled = effectiveMode === 'view' || field.readOnly === true || field.disabled === true || isLinkedField;
    var errorMessage = this.state.fieldErrors[field.id];
    var placeholder = field.config && field.config.placeholder ? field.config.placeholder : undefined;
    var description = this.props.showFieldDescription ? (field.description || (field.config && field.config.helpText) || '') : '';

    // Build label styles with field-specific font settings; always resolved (own default when unset)
    // so a field's label never inherits font styling meant for its surrounding container.
    var labelStyle: React.CSSProperties = {
      fontSize: field.labelFontSize || 14,
      fontFamily: field.labelFontFamily || DEFAULT_THEME_FONT_FAMILY,
      fontWeight: (field.labelFontWeight || 'normal') as any,
      color: field.labelColor || '#000000',
    };

    // Build field wrapper styles with background and border
    var fieldWrapperStyle: React.CSSProperties = { padding: '12px', marginBottom: '16px' };
    // Determine label position: field override or theme default
    var labelPosition = field.labelPosition || (this.state.schema && this.state.schema.theme && this.state.schema.theme.labelPosition) || 'top';
    var labelWrapperStyle: React.CSSProperties = labelPosition === 'left' 
      ? { display: 'flex', gap: '16px', alignItems: 'flex-start' }
      : { display: 'block' };
    var labelBlockStyle: React.CSSProperties = Object.assign(
      labelPosition === 'left'
        ? { flexShrink: 0, minWidth: '120px', paddingTop: '6px' }
        : { marginBottom: '10px' },
      labelStyle
    );
    if (field.fieldBackgroundColor) { fieldWrapperStyle.backgroundColor = field.fieldBackgroundColor; }
    if (field.fieldBorderColor || field.fieldBorderWidth) {
      fieldWrapperStyle.border = (field.fieldBorderWidth || 1) + 'px solid ' + (field.fieldBorderColor || '#cccccc');
    }
    if (field.fieldBorderStyle === 'rounded') {
      fieldWrapperStyle.borderRadius = String(field.fieldBorderRadius || 8) + 'px';
    }

    // Build input styles with field-specific font settings; always resolved (own default when unset)
    // so a field's input never inherits font styling meant for its surrounding container.
    var inputStyle: React.CSSProperties = {
      width: '100%', padding: '8px', boxSizing: 'border-box',
      fontSize: field.inputFontSize || 14,
      fontFamily: field.inputFontFamily || DEFAULT_THEME_FONT_FAMILY,
      fontWeight: (field.inputFontWeight || 'normal') as any,
      color: field.inputColor || '#000000',
    };
    var reportValueStyle: React.CSSProperties = Object.assign({}, inputStyle, {
      padding: '8px',
      border: '1px solid #d1d1d1',
      borderRadius: '2px',
      backgroundColor: '#f3f2f1',
      minHeight: '36px',
      lineHeight: '20px',
      whiteSpace: 'pre-wrap',
      wordBreak: 'break-word',
    });

    if (isLinkedField) {
      console.log('[ReportFormsWebPart] Rendering linked field as READONLY: ' + field.fieldName + ' (ID: ' + field.id + '), Value: ' + value + ', Mode: ' + effectiveMode);
    }

    // Comprehensive description diagnostics
    console.log('[ReportFormsWebPart] DESCRIPTION_DEBUG - Field: "' + field.fieldName + '", showFieldDescription toggle: ' + (this.props.showFieldDescription === true) + ', field.description: ' + (field.description ? '"' + field.description + '"' : 'NONE') + ', field.config.helpText: ' + ((field.config && field.config.helpText) ? '"' + field.config.helpText + '"' : 'NONE') + ', final description: ' + (description ? '"' + description + '"' : 'EMPTY'));

    if (this.props.showFieldDescription && (field.description || (field.config && field.config.helpText))) {
      console.log('[ReportFormsWebPart] Field "' + field.fieldName + '" has description: ' + description);
    }

    if (field.visible === false || field.type === 'newline') {
      return null;
    }

    switch (field.type) {
      case 'listcontrol':
        var listControlConfig = field.config || {};
        var filterSourceRef = String(listControlConfig.listControlFilterSourceField || '').toLowerCase();
        var filterSourceField: FormField | null = null;
        if (filterSourceRef && this.state.schema) {
          for (var sourceStepIndex = 0; sourceStepIndex < this.state.schema.steps.length && !filterSourceField; sourceStepIndex += 1) {
            var sourceStepFields = this.state.schema.steps[sourceStepIndex].fields;
            for (var sourceFieldIndex = 0; sourceFieldIndex < sourceStepFields.length; sourceFieldIndex += 1) {
              var sourceField = sourceStepFields[sourceFieldIndex];
              if (sourceField && (String(sourceField.id || '').toLowerCase() === filterSourceRef || String(sourceField.fieldName || '').toLowerCase() === filterSourceRef)) {
                filterSourceField = sourceField;
                break;
              }
            }
          }
        }
        var runtimeFilterJson = '';
        var runtimeFilterValue: any = '';
        if (filterSourceField) {
          runtimeFilterValue = getListControlFilterValue(filterSourceField, this.state.values[filterSourceField.id]);
          if (!hasUsableValue(runtimeFilterValue)) {
            runtimeFilterValue = getListControlFilterValue(filterSourceField, getItemFieldValue(this.state.loadedItem, filterSourceField.fieldName));
          }
        } else if (filterSourceRef) {
          var loadedItem = this.state.loadedItem || {};
          var loadedItemKey = listControlConfig.listControlFilterSourceField || '';
          var rawRuntimeValue = getItemFieldValue(loadedItem, loadedItemKey);
          if ((rawRuntimeValue === undefined || rawRuntimeValue === null) && filterSourceRef === 'id') {
            rawRuntimeValue = loadedItem.ID !== undefined ? loadedItem.ID : (loadedItem.Id !== undefined ? loadedItem.Id : this.state.resolvedItemId);
          }
          runtimeFilterValue = getListControlFilterValue(null, rawRuntimeValue);
        }
        var linkedFilterTarget = String(this.props.linkedFieldTarget || '').trim().toLowerCase();
        if (!hasUsableValue(runtimeFilterValue) && this.props.useDynamicValueAsFilter && filterSourceRef === linkedFilterTarget) {
          var runtimeQueryKey = String(this.props.itemIdQueryParam || 'itemid').trim() || 'itemid';
          var runtimeQueryValue = this.getUrlQueryValue(runtimeQueryKey);
          var linkedRuntimeValue = hasTextValue(runtimeQueryValue) ? runtimeQueryValue : this.props.linkedFieldValue;
          runtimeFilterValue = getListControlFilterValue(filterSourceField, linkedRuntimeValue);
        }
        if (filterSourceRef && listControlConfig.listControlFilterTargetField) {
          runtimeFilterJson = JSON.stringify([{
            field: listControlConfig.listControlFilterTargetField,
            operator: listControlConfig.listControlFilterOperator || 'eq',
            logical: 'and',
            valueType: 'static',
            value: runtimeFilterValue
          }]);
        }
        return (
          <div style={fieldWrapperStyle}>
            <div style={labelWrapperStyle}>
              <div style={labelBlockStyle}>{field.label}{this.renderDescriptionIcon(description)}</div>
              <ListControlHost
                fieldId={field.id}
                sourceId={listControlConfig.listControlSourceId || ''}
                sourceName={listControlConfig.listControlSourceName || field.label}
                sourceListName={listControlConfig.listControlSourceListName || ''}
                runtimeFilterJson={runtimeFilterJson}
                isPageEditMode={this.props.isPageEditMode}
              />
            </div>
          </div>
        );
      case 'customimage':
        var customImageConfig = field.config || {};
        var customImageUrl = String(customImageConfig.imageUrl || '').trim();
        if (!customImageUrl) {
          return null;
        }
        var customImageWrapperStyle: React.CSSProperties = {
          marginBottom: '16px',
          textAlign: (customImageConfig.imageAlignment || 'left') as any
        };
        var customImageStyle: React.CSSProperties = {
          display: 'inline-block',
          maxWidth: '100%',
          width: customImageConfig.imageWidth ? String(customImageConfig.imageWidth) + 'px' : 'auto',
          height: customImageConfig.imageHeight ? String(customImageConfig.imageHeight) + 'px' : 'auto',
          objectFit: (customImageConfig.imageFit || 'contain') as any
        };
        return (
          <div style={customImageWrapperStyle}>
            <img src={customImageUrl} alt={String(customImageConfig.imageAltText || field.label || '')} style={customImageStyle} />
          </div>
        );
      case 'divider':
        var dividerConfig = field.config || {};
        var dividerSpacing = dividerConfig.dividerSpacing === undefined ? 16 : Math.max(0, dividerConfig.dividerSpacing);
        var dividerBorder = String(Math.max(1, dividerConfig.dividerThickness || 1)) + 'px ' + String(dividerConfig.dividerStyle || 'solid') + ' ' + String(dividerConfig.dividerColor || '#c8c6c4');
        if (dividerConfig.dividerOrientation === 'vertical') {
          return <div style={{ display: 'flex', justifyContent: 'center' }}><div role="separator" aria-orientation="vertical" style={{ borderLeft: dividerBorder, height: String(Math.max(1, dividerConfig.dividerLength || 80)) + 'px', margin: '0 ' + String(dividerSpacing) + 'px' }} /></div>;
        }
        return <hr style={{ border: 0, borderTop: dividerBorder, margin: String(dividerSpacing) + 'px 0' }} />;
      case 'richtext':
        // Custom richtext fields render as content divs, not form inputs
        // Always visible, never editable, for informational purposes
        var richTextStyle: React.CSSProperties = {
          padding: '12px', marginBottom: '16px', wordBreak: 'break-word', lineHeight: '1.6',
          fontSize: field.inputFontSize || 14,
          fontFamily: field.inputFontFamily || DEFAULT_THEME_FONT_FAMILY,
          fontWeight: (field.inputFontWeight || 'normal') as any,
          color: field.inputColor || '#000000',
        };
        if (field.fieldBackgroundColor) { richTextStyle.backgroundColor = field.fieldBackgroundColor; }
        if (field.fieldBorderColor || field.fieldBorderWidth) {
          richTextStyle.border = (field.fieldBorderWidth || 1) + 'px solid ' + (field.fieldBorderColor || '#cccccc');
        }
        if (field.fieldBorderStyle === 'rounded') {
          richTextStyle.borderRadius = String(field.fieldBorderRadius || 8) + 'px';
        }
        var richTextContent = decodeHtmlEntities(String(field.defaultValue || ''));
        return <div style={richTextStyle} dangerouslySetInnerHTML={{ __html: richTextContent } as any} />;
      case 'text':
      case 'multiline':
        var isMultilineRichText = isRichTextMultilineField(field);
        var richTextValueRaw = value === undefined || value === null ? '' : String(value);
        var richTextValue = decodeHtmlEntities(richTextValueRaw);
        var plainTextInputStyle = Object.assign({}, inputStyle, {
          direction: 'ltr' as any,
          textAlign: 'left' as any,
        });
        return (
          <div style={fieldWrapperStyle}>
            <label style={labelWrapperStyle}>
              <div style={labelBlockStyle}>{field.label}{this.renderDescriptionIcon(description)}</div>
            {field.type === 'text' ? (
              <input
                dir="ltr"
                type="text"
                disabled={disabled}
                required={field.required === true}
                value={value === undefined || value === null ? '' : String(value)}
                onChange={(ev) => this.setFieldValue(field.id, ev.currentTarget.value)}
                placeholder={placeholder}
                style={plainTextInputStyle}
              />
            ) : (
              isMultilineRichText ? (
                disabled ? (
                  <div
                    dir="ltr"
                    style={Object.assign({}, inputStyle, {
                      minHeight: '120px',
                      padding: '12px',
                      border: '1px solid #d1d1d1',
                      backgroundColor: '#f3f2f1',
                      lineHeight: '1.6',
                      overflowY: 'auto',
                      direction: 'ltr',
                      unicodeBidi: 'normal',
                      textAlign: 'left',
                    })}
                    dangerouslySetInnerHTML={{ __html: richTextValue } as any}
                  />
                ) : (
                  <div dir="ltr" style={{ direction: 'ltr', textAlign: 'left' }}>
                    <RichTextEditor
                      value={richTextValueRaw}
                      onChange={(html) => this.setFieldValue(field.id, html)}
                      placeholder={placeholder}
                    />
                  </div>
                )
              ) : (
                <textarea
                  dir="ltr"
                  disabled={disabled}
                  required={field.required === true}
                  value={value === undefined || value === null ? '' : String(value)}
                  onChange={(ev) => this.setFieldValue(field.id, ev.currentTarget.value)}
                  placeholder={placeholder}
                  rows={4}
                  style={plainTextInputStyle}
                />
              )
            )}
            </label>
            {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
          </div>
        );
      case 'number':
        return (
          <div style={fieldWrapperStyle}>
            <label style={labelWrapperStyle}>
              <div style={labelBlockStyle}>{field.label}{this.renderDescriptionIcon(description)}</div>
              <div style={reportValueStyle}>{value === undefined || value === null || value === '' ? '-' : String(value)}</div>
            </label>
            {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
          </div>
        );
      case 'boolean':
        var booleanText = field.config && field.config.booleanText ? field.config.booleanText : '';
        var booleanLayoutStyle: React.CSSProperties = { display: 'flex', gap: '10px' };
        if (labelPosition === 'bottom') {
          booleanLayoutStyle.flexDirection = 'column-reverse';
        } else if (labelPosition === 'left') {
          booleanLayoutStyle.flexDirection = 'row';
          booleanLayoutStyle.alignItems = 'center';
        } else if (labelPosition === 'right') {
          booleanLayoutStyle.flexDirection = 'row-reverse';
          booleanLayoutStyle.justifyContent = 'flex-end';
          booleanLayoutStyle.alignItems = 'center';
        } else {
          booleanLayoutStyle.flexDirection = 'column';
        }
        var booleanFieldLabelStyle: React.CSSProperties = Object.assign({}, labelStyle, {
          flexShrink: 0,
          marginBottom: 0,
        });
        var booleanControlStyle: React.CSSProperties = {
          display: 'inline-flex',
          alignItems: 'center',
          gap: '6px',
          width: 'fit-content',
        };
        var booleanTextStyle: React.CSSProperties = {
          fontSize: field.inputFontSize || 14,
          fontFamily: field.inputFontFamily || DEFAULT_THEME_FONT_FAMILY,
          fontWeight: (field.inputFontWeight || 'normal') as any,
          color: field.inputColor || '#000000',
        };
        return (
          <div style={fieldWrapperStyle}>
            <div style={booleanLayoutStyle}>
              <div style={booleanFieldLabelStyle}>{field.label}{this.renderDescriptionIcon(description)}</div>
              <label style={booleanControlStyle}>
              <input
                type="checkbox"
                checked={value === true}
                disabled={disabled}
                aria-label={field.label}
                onChange={(ev) => this.setFieldValue(field.id, ev.currentTarget.checked)}
              />
                {booleanText && <span style={booleanTextStyle}>{booleanText}</span>}
              </label>
            </div>
            {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
          </div>
        );
      case 'dropdown':
        return (
          <div style={fieldWrapperStyle}>
            <div style={Object.assign({ marginBottom: '10px' }, labelStyle)}>{field.label}{this.renderDescriptionIcon(description)}</div>
            <div style={reportValueStyle}>{value === undefined || value === null || value === '' ? '-' : String(value)}</div>
            {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
          </div>
        );
      case 'lookup':
        var lookupItems = this.state.lookupOptions[field.id] || [];
        var lookupOptions: ISelectOption[] = lookupItems.map(function(item) {
          return { key: String(item.Id), text: item.Title || String(item.Id) };
        });
        if (effectiveMode === 'view') {
          return (
            <div style={fieldWrapperStyle}>
              <div style={Object.assign({ marginBottom: '10px' }, labelStyle)}>{field.label}{this.renderDescriptionIcon(description)}</div>
              <div style={reportValueStyle}>{this.formatReportSelection(value, lookupOptions)}</div>
              {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
            </div>
          );
        }
        var isMultiLookupField = !!(field.config && field.config.allowMultiple === true);
        if (isMultiLookupField) {
          var selectedLookupValues = Array.isArray(value) ? value as string[] : [];
          var lookupExclusiveChoiceValue = this.getExclusiveChoiceValue(field);
          var lookupExclusiveLocked = this.isExclusiveChoiceLocked(field, selectedLookupValues);
          return (
            <div style={fieldWrapperStyle}>
              <div style={Object.assign({ marginBottom: '10px' }, labelStyle)}>{field.label}{this.renderDescriptionIcon(description)}</div>
              {lookupOptions.map((option) => {
                var checked = selectedLookupValues.indexOf(option.key) >= 0;
                var locked = lookupExclusiveLocked && !checked;
                return (
                  <label key={option.key} style={{ display: 'block', marginLeft: '12px' }}>
                    <input
                      type="checkbox"
                      checked={checked}
                      disabled={disabled || locked}
                      onChange={(ev) => {
                        var nextValues = this.updateExclusiveMultiChoiceSelection(field, selectedLookupValues, option.key, ev.currentTarget.checked);
                        this.setFieldValue(field.id, nextValues);
                      }}
                    />{' '}
                    {option.text}
                  </label>
                );
              })}
              {lookupExclusiveChoiceValue && <div style={{ marginTop: '6px', fontSize: '12px', color: '#605e5c' }}>Exclusive value: {lookupExclusiveChoiceValue}</div>}
              {this.renderFieldHelpAndError(description, errorMessage)}
            </div>
          );
        }
        var lookupValue = value === undefined || value === null ? '' : String(value);
        if (lookupValue || lookupItems.length > 0) {
          var foundOption = lookupOptions.find(function(opt) { return opt.key === lookupValue; });
          var optionKeysList = lookupOptions.map(function(opt) { return '"' + opt.key + '"'; }).join(', ');
          console.log('[ReportFormsWebPart] ===== LOOKUP FIELD DIAGNOSTIC =====');
          console.log('[ReportFormsWebPart] Field: ' + field.label + ' (' + field.fieldName + ', ID: ' + field.id + ')');
          console.log('[ReportFormsWebPart] Current Value: "' + lookupValue + '" (type: ' + typeof lookupValue + ', length: ' + String(lookupValue).length + ')');
          console.log('[ReportFormsWebPart] Available Options (' + lookupOptions.length + '): [' + optionKeysList + ']');
          console.log('[ReportFormsWebPart] Selected Option: ' + (foundOption ? '"' + foundOption.text + '"' : 'NOT FOUND - VALUE MISMATCH!'));
          console.log('[ReportFormsWebPart] =====================================');
        }
        return (
          <div style={fieldWrapperStyle}>
            <div style={Object.assign({ marginBottom: '10px' }, labelStyle)}>{field.label}{this.renderDescriptionIcon(description)}</div>
            {this.getChoiceDisplayMode(field) === 'radio' ? (
              <div>
                {lookupOptions.map((option) => {
                  return (
                    <label key={option.key} style={{ display: 'block', marginLeft: '12px' }}>
                      <input
                        type="radio"
                        name={field.id}
                        checked={lookupValue === option.key}
                        disabled={disabled}
                        onChange={() => this.setFieldValue(field.id, option.key)}
                      />{' '}
                      {option.text}
                    </label>
                  );
                })}
              </div>
            ) : (
              <select
                disabled={disabled}
                value={lookupValue}
                onChange={(ev) => this.setFieldValue(field.id, ev.currentTarget.value)}
                style={inputStyle}
              >
                <option value="">{placeholder || strings.FieldPlaceholderSelect}</option>
                {lookupOptions.map(function(option) {
                  return <option key={option.key} value={option.key}>{option.text}</option>;
                })}
              </select>
            )}
              {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
          </div>
        );
      case 'multiselect':
        var selectedValues = Array.isArray(value) ? value as string[] : [];
        var choices = field.config && field.config.choices ? field.config.choices : [];
        var exclusiveChoiceValue = this.getExclusiveChoiceValue(field);
        var multiSelectExclusiveLocked = this.isExclusiveChoiceLocked(field, selectedValues);
        return (
          <div style={fieldWrapperStyle}>
            <div style={Object.assign({ marginBottom: '10px' }, labelStyle)}>{field.label}{this.renderDescriptionIcon(description)}</div>
            {choices.map((choice) => {
              var checked = selectedValues.indexOf(choice) >= 0;
              var disabledChoice = disabled || (multiSelectExclusiveLocked && choice !== exclusiveChoiceValue);
              return (
                <label key={choice} style={{ display: 'block', marginLeft: '12px' }}>
                  <input
                    type="checkbox"
                    checked={checked}
                    disabled={disabledChoice}
                    onChange={(ev) => {
                      var nextValues = this.updateExclusiveMultiChoiceSelection(field, selectedValues, choice, ev.currentTarget.checked);
                      this.setFieldValue(field.id, nextValues);
                    }}
                  />{' '}
                  {choice}
                </label>
              );
            })}
            {this.renderFieldHelpAndError(description, errorMessage)}
          </div>
        );
      case 'person':
        var personItems = this.state.lookupOptions[field.id] || [];
        var personOptions: ISelectOption[] = personItems.map(function(item) {
          var label = item.Title || item.Email || String(item.Id);
          return { key: String(item.Id), text: item.Email ? label + ' (' + item.Email + ')' : label };
        });
        if (effectiveMode === 'view') {
          return (
            <div style={fieldWrapperStyle}>
              <div style={Object.assign({ marginBottom: '10px' }, labelStyle)}>{field.label}{this.renderDescriptionIcon(description)}</div>
              <div style={reportValueStyle}>{this.formatReportSelection(value, personOptions)}</div>
              {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
            </div>
          );
        }
        var isMultiPersonField = !!(field.config && field.config.allowMultiple === true);
        if (isMultiPersonField) {
          var selectedPeople = Array.isArray(value) ? value as string[] : [];
          return (
            <div style={fieldWrapperStyle}>
              <div style={Object.assign({ marginBottom: '10px' }, labelStyle)}>{field.label}{this.renderDescriptionIcon(description)}</div>
              {personOptions.map((option) => {
                var checked = selectedPeople.indexOf(option.key) >= 0;
                return (
                  <label key={option.key} style={{ display: 'block', marginLeft: '12px' }}>
                    <input
                      type="checkbox"
                      checked={checked}
                      disabled={disabled}
                      onChange={(ev) => {
                        var nextValues = selectedPeople.slice(0);
                        var existingIndex = nextValues.indexOf(option.key);
                        if (ev.currentTarget.checked) {
                          if (existingIndex < 0) {
                            nextValues.push(option.key);
                          }
                        } else if (existingIndex >= 0) {
                          nextValues.splice(existingIndex, 1);
                        }
                        this.setFieldValue(field.id, nextValues);
                      }}
                    />{' '}
                    {option.text}
                  </label>
                );
              })}
              {this.renderFieldHelpAndError(description, errorMessage)}
            </div>
          );
        }
        var personValue = value === undefined || value === null ? '' : String(value);
        return (
          <div style={fieldWrapperStyle}>
            <label style={labelWrapperStyle}>
              <div style={labelBlockStyle}>{field.label}{this.renderDescriptionIcon(description)}</div>
              <select
                disabled={disabled}
                value={personValue}
                onChange={(ev) => this.setFieldValue(field.id, ev.currentTarget.value)}
                style={inputStyle}
              >
                <option value="">{placeholder || strings.FieldPlaceholderSelect}</option>
                {personOptions.map(function(option) {
                  return <option key={option.key} value={option.key}>{option.text}</option>;
                })}
              </select>
            </label>
            {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
          </div>
        );
      case 'attachment':
        var existingAttachments = this._existingAttachments[field.id] || [];
        var attachmentDeleteAllowed = this.isAttachmentDeleteAllowed(field);
        var pendingAttachments = this.getPendingAttachments(field.id);
        return (
          <div style={fieldWrapperStyle}>
            <div style={labelBlockStyle}>{field.label}{this.renderDescriptionIcon(description)}</div>
            {!disabled && <div style={{ marginBottom: '8px', fontSize: '12px', color: '#605e5c' }}>{strings.FieldAttachmentUploadOnSubmit}</div>}
            {!disabled && (
              <div style={{ marginBottom: '10px' }}>
                <input
                  ref={(input) => { this._attachmentInputRefs[field.id] = input; }}
                  type="file"
                  multiple={true}
                  onChange={(ev) => {
                    var files = ev.currentTarget.files;
                    var selectedFiles: File[] = [];
                    if (files) {
                      for (var fileIndex = 0; fileIndex < files.length; fileIndex += 1) {
                        selectedFiles.push(files[fileIndex]);
                      }
                    }
                    this.setFieldValue(field.id, selectedFiles as any);
                  }}
                />
                <div style={{ marginTop: '8px' }}>
                  <DefaultButton
                    text={strings.FieldAttachmentClear}
                    onClick={() => {
                      this.clearAttachmentInput(field.id);
                      this.setFieldValue(field.id, [] as any);
                    }}
                    disabled={pendingAttachments.length === 0}
                  />
                </div>
              </div>
            )}
            {existingAttachments.length > 0 && (
              <div style={{ marginBottom: '10px', overflowX: 'auto' }}>
                <table style={{ width: '100%', minWidth: '680px', borderCollapse: 'collapse', border: '1px solid #edebe9', textAlign: 'left' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#f3f2f1' }}>
                      <th scope="col" style={{ padding: '8px 10px', borderBottom: '1px solid #d2d0ce', fontWeight: 600 }}>File</th>
                      <th scope="col" style={{ padding: '8px 10px', borderBottom: '1px solid #d2d0ce', fontWeight: 600 }}>Type</th>
                      <th scope="col" style={{ padding: '8px 10px', borderBottom: '1px solid #d2d0ce', fontWeight: 600 }}>Size</th>
                      <th scope="col" style={{ padding: '8px 10px', borderBottom: '1px solid #d2d0ce', fontWeight: 600 }}>Modified</th>
                      <th scope="col" style={{ padding: '8px 10px', borderBottom: '1px solid #d2d0ce', fontWeight: 600 }}>Created</th>
                    </tr>
                  </thead>
                  <tbody>
                    {existingAttachments.map((attachment) => {
                      var markedForDelete = this.isAttachmentMarkedForDelete(field.id, attachment);
                      var attachmentTextStyle = markedForDelete ? { textDecoration: 'line-through', color: '#a80000' } : undefined;
                      return (
                        <tr key={attachment.fileName}>
                          <td style={{ padding: '9px 10px', borderBottom: '1px solid #edebe9', wordBreak: 'break-word' }}>
                            <span style={attachmentTextStyle}>
                              {attachment.serverRelativeUrl ? (
                                <a href={attachment.serverRelativeUrl} target="_blank" rel="noopener noreferrer">{attachment.fileName}</a>
                              ) : attachment.fileName}
                            </span>
                            {markedForDelete && <span style={{ marginLeft: '8px', fontSize: '12px', color: '#a80000' }}>{strings.FieldAttachmentWillDelete}</span>}
                            {!disabled && attachmentDeleteAllowed && (
                              <div style={{ marginTop: '6px' }}>
                                <DefaultButton
                                  text={markedForDelete ? strings.FieldAttachmentUndoDelete : strings.FieldAttachmentDelete}
                                  ariaLabel={strings.FieldAttachmentRemoveAria}
                                  onClick={() => this.toggleAttachmentMarkedForDelete(field.id, attachment)}
                                />
                              </div>
                            )}
                          </td>
                          <td style={{ padding: '9px 10px', borderBottom: '1px solid #edebe9' }}>{this.getAttachmentFileType(attachment.fileName)}</td>
                          <td style={{ padding: '9px 10px', borderBottom: '1px solid #edebe9', whiteSpace: 'nowrap' }}>{this.formatAttachmentSize(attachment.length)}</td>
                          <td style={{ padding: '9px 10px', borderBottom: '1px solid #edebe9', whiteSpace: 'nowrap' }}>{this.formatReportDateTime(attachment.timeLastModified)}</td>
                          <td style={{ padding: '9px 10px', borderBottom: '1px solid #edebe9', whiteSpace: 'nowrap' }}>{this.formatReportDateTime(attachment.timeCreated)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
            {pendingAttachments.length > 0 && (
              <div style={{ marginBottom: '10px' }}>
                {pendingAttachments.map(function(file) {
                  return <div key={file.name}>{file.name}</div>;
                })}
              </div>
            )}
            {existingAttachments.length === 0 && pendingAttachments.length === 0 && <div style={{ fontSize: '12px', color: '#605e5c' }}>{strings.FieldAttachmentEmpty}</div>}
            {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
          </div>
        );
      case 'datetime':
        return (
          <div style={fieldWrapperStyle}>
            <label style={labelWrapperStyle}>
              <div style={labelBlockStyle}>{field.label}{this.renderDescriptionIcon(description)}</div>
              <div style={reportValueStyle}>{this.formatReportDateTime(value, field.config && field.config.displayFormat, field.config && field.config.customDateFormat, field.config && field.config.customDateFormatCase)}</div>
            </label>
            {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
          </div>
        );
      case 'url':
        var urlValue = value === undefined || value === null ? '' : String(value).trim();
        var urlText = this._urlFieldDescriptions[field.id] || urlValue;
        return (
          <div style={fieldWrapperStyle}>
            <label style={labelWrapperStyle}>
              <div style={labelBlockStyle}>{field.label}{this.renderDescriptionIcon(description)}</div>
              <div style={reportValueStyle}>
                {urlValue ? <a href={urlValue} target="_blank" rel="noopener noreferrer">{urlText}</a> : '-'}
              </div>
            </label>
            {this.renderFieldHelpAndError(description, errorMessage, labelPosition)}
          </div>
        );
      default:
        return this.renderUnsupportedField(field);
    }
  }

  private renderPreviewField(field: FormField): JSX.Element {
    var gridStyle = field.startNewRow === true
      ? { gridColumn: '1 / -1' }
      : { gridColumn: 'span ' + String(Math.max(1, field.columnSpan || 1)) };

    return (
      <div key={field.id} style={gridStyle}>
        {this.renderField(field)}
      </div>
    );
  }

  private getContainerWidthStyle(): React.CSSProperties {
    var configuredWidth = Number(this.props.containerWidth || 0);
    var containerStyle: React.CSSProperties = {
      width: '100%',
      maxWidth: configuredWidth > 0 ? String(configuredWidth) + 'px' : 'none',
      marginLeft: 'auto',
      marginRight: 'auto',
      boxSizing: 'border-box',
    };
    if (this.props.isFullWidth) {
      containerStyle.paddingRight = '30px';
    }
    return containerStyle;
  }

  public render(): JSX.Element {
    if (this.props.isInDesignerMode) {
      return (
        <div className={styles.designerMode}>
          <div className={styles.designerToolbar}>
            <h2>{strings.DesignerTitle}</h2>
            <div className={styles.designerActions}>
              <DefaultButton onClick={this.props.onToggleDesignerMode}>{strings.DesignerBackButton}</DefaultButton>
              <PrimaryButton onClick={this.handleDesignerSave} disabled={!this.state.schema}>{strings.CommonSave}</PrimaryButton>
            </div>
          </div>

          {this.state.loading && <div className={styles.loading}>{strings.CommonLoading}</div>}
          {!this.state.loading && this.state.error && <MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar>}
          {!this.state.loading && !this.state.error && this.state.schema && (
            <FormDesigner
              schema={this.state.schema}
              context={this.props.context}
              listName={this.state.schema.listName || this.props.listName}
              onChange={this.handleDesignerChange}
            />
          )}
        </div>
      );
    }

    if (this.state.loading) {
      return (
        <div className={styles.container} style={this.getContainerWidthStyle()}>
          <div className={styles.loading}>{strings.CommonLoading}</div>
        </div>
      );
    }

    if (this.state.error) {
      return (
        <div className={styles.container} style={this.getContainerWidthStyle()}>
          <MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar>
        </div>
      );
    }

    if (!this.state.schema) {
      return (
        <div className={styles.container} style={this.getContainerWidthStyle()}>
          <MessageBar messageBarType={MessageBarType.warning}>{strings.RuntimeNoSchemaWarning}</MessageBar>
        </div>
      );
    }

    var schema = this.state.schema;
    var effectiveMode = this.getEffectiveMode();
    var canSubmit = effectiveMode !== 'view';
    var canSwitchToEdit = this.state.canEditRecords;
    var canSubmitByPermission = effectiveMode === 'edit' ? this.state.canEditRecords : this.state.canAddRecords;
    var permissionDeniedForCurrentMode = effectiveMode === 'new' ? !this.state.canAddRecords : !this.state.canEditRecords;
    var disableSubmitButton = this.state.isSubmitting || !canSubmitByPermission;
    var disableEditButton = this.state.isSubmitting || !canSwitchToEdit;
    var addSubmitLabel =
      (hasTextValue(this.props.addSubmitButtonLabel) ? this.props.addSubmitButtonLabel : undefined) ||
      (hasTextValue(this.props.submitButtonLabel) ? this.props.submitButtonLabel : undefined) ||
      strings.FormSubmitLabelAdd;
    var editSubmitLabel =
      (hasTextValue(this.props.editSubmitButtonLabel) ? this.props.editSubmitButtonLabel : undefined) ||
      (hasTextValue(this.props.submitButtonLabel) ? this.props.submitButtonLabel : undefined) ||
      strings.FormSubmitLabelEdit;
    var submitLabel = effectiveMode === 'edit' ? editSubmitLabel : addSubmitLabel;
    var cancelLabel = this.props.cancelButtonLabel || strings.CommonCancel;
    var hasDynamicItemBinding = !!String(this.props.dynamicItemReference || '').trim();
    var viewBackTarget = this.getPreferredRedirectUrl(this.props.cancelRedirectUrl);
    var showViewBackButton = this.props.mode === 'view' && this.state.resolvedItemId > 0 && !this.state.isViewEditing && !hasDynamicItemBinding && !!viewBackTarget;
    var buttonTextColor = this.props.buttonTextColor || '#000000';
    var buttonBackgroundColor = this.props.buttonBackgroundColor || '#f0f0f0';
    var buttonBorderColor = this.props.buttonBorderColor || buttonBackgroundColor;
    var buttonBorderWidth = toPositiveInt(this.props.buttonBorderWidth);
    if (buttonBorderWidth < 0) {
      buttonBorderWidth = 0;
    }
    var buttonFontFamily = this.props.buttonFontFamily || DEFAULT_THEME_FONT_FAMILY;
    var buttonFontSize = toPositiveInt(this.props.buttonFontSize) || 14;
    var buttonFontStyle = this.props.buttonFontStyle || 'normal';
    var buttonFontWeight = this.props.buttonFontBold ? 'bold' : 'normal';
    var buttonCornerStyle = this.props.buttonCornerStyle || 'square';
    var buttonCornerRadius = toPositiveInt(this.props.buttonCornerRadius) || 4;
    var buttonBorderRadius = buttonCornerStyle === 'rounded' ? (String(buttonCornerRadius) + 'px') : '0px';

    var buttonRootStyles: any = {
      backgroundColor: buttonBackgroundColor,
      borderColor: buttonBorderColor,
      borderWidth: String(buttonBorderWidth) + 'px',
      borderStyle: 'solid',
      borderRadius: buttonBorderRadius,
      color: buttonTextColor,
      fontFamily: buttonFontFamily,
      fontSize: String(buttonFontSize) + 'px',
      fontStyle: buttonFontStyle,
      fontWeight: buttonFontWeight,
    };

    var buttonLabelStyles = {
      color: buttonTextColor,
      fontFamily: buttonFontFamily,
      fontSize: String(buttonFontSize) + 'px',
      fontStyle: buttonFontStyle,
      fontWeight: buttonFontWeight,
    };

    var primaryButtonStyles = {
      root: buttonRootStyles,
      rootHovered: buttonRootStyles,
      rootPressed: buttonRootStyles,
      rootChecked: buttonRootStyles,
      rootFocused: buttonRootStyles,
      rootDisabled: {
        ...buttonRootStyles,
        opacity: 0.6,
      },
      label: buttonLabelStyles,
      labelHovered: buttonLabelStyles,
      labelPressed: buttonLabelStyles,
      icon: { color: buttonTextColor },
      iconHovered: { color: buttonTextColor },
      iconPressed: { color: buttonTextColor },
    };

    var defaultButtonStyles = {
      root: buttonRootStyles,
      rootHovered: buttonRootStyles,
      rootPressed: buttonRootStyles,
      rootChecked: buttonRootStyles,
      rootFocused: buttonRootStyles,
      rootDisabled: {
        ...buttonRootStyles,
        opacity: 0.6,
      },
      label: buttonLabelStyles,
      labelHovered: buttonLabelStyles,
      labelPressed: buttonLabelStyles,
      icon: { color: buttonTextColor },
      iconHovered: { color: buttonTextColor },
      iconPressed: { color: buttonTextColor },
    };

    // Build container styles from theme
    var containerStyle: React.CSSProperties = this.getContainerWidthStyle();
    var formTitleStyle: React.CSSProperties = { textAlign: schema.nameAlignment || 'left' };
    // Form font settings apply to the form heading only; always resolved (own default when unset)
    // so the title never inherits font styling meant for a different level.
    formTitleStyle.fontSize = (schema.theme && schema.theme.fontSize) || 14;
    formTitleStyle.fontFamily = (schema.theme && schema.theme.fontFamily) || DEFAULT_THEME_FONT_FAMILY;
    formTitleStyle.fontWeight = ((schema.theme && schema.theme.fontWeight) || 'normal') as any;
    formTitleStyle.color = (schema.theme && schema.theme.color) || '#000000';
    if (schema.theme) {
      // Background and border settings for form wrapper
      if (schema.theme.backgroundColor) { containerStyle.backgroundColor = schema.theme.backgroundColor; }
      if (schema.theme.borderColor || schema.theme.borderWidth) {
        containerStyle.border = (schema.theme.borderWidth || 1) + 'px solid ' + (schema.theme.borderColor || '#cccccc');
        containerStyle.padding = '16px';
      }
      if (schema.theme.borderStyle === 'rounded') {
        containerStyle.borderRadius = String(schema.theme.borderRadius || 8) + 'px';
      }
      // Add padding if background is set
      if (schema.theme.backgroundColor && !schema.theme.borderColor && !schema.theme.borderWidth) {
        containerStyle.padding = '16px';
      }
    }
    if (this.props.isFullWidth) {
      var themedContainerPadding = schema.theme && (schema.theme.backgroundColor || schema.theme.borderColor || schema.theme.borderWidth) ? 16 : 20;
      containerStyle.paddingRight = String(themedContainerPadding + 10) + 'px';
    }

    var dynamicDebugMessage = 'Dynamic binding status: itemId=' + String(this.props.dynamicItemId || 0)
      + ', resolvedItemId=' + String(this.state.resolvedItemId || 0)
      + ', mode=' + String(effectiveMode)
      + ', itemRef=' + String(this.props.dynamicItemReference || '(empty)')
      + ', modeRef=' + String(this.props.dynamicModeReference || '(empty)')
      + ', preferredSource=' + String(this.props.dynamicPreferredSourceInstanceId || '(empty)')
      + ', linkedFieldTarget=' + String(this.props.linkedFieldTarget || '(none)')
      + ', linkedFieldValue=' + (this.props.linkedFieldValue === undefined || this.props.linkedFieldValue === null ? '(empty)' : String(this.props.linkedFieldValue));

    this.logDynamicDiagnosticsToConsole(dynamicDebugMessage);

    return (
      <div className={styles.sharePointDynamicForm + ' ' + styles.container + (this.props.isDarkTheme ? ' ' + styles.dark : '')} style={containerStyle}>
        {this.props.isPageEditMode && (
          <div className={styles.editBar}>
            <DefaultButton onClick={this.props.onToggleDesignerMode}>{strings.DesignerButton}</DefaultButton>
          </div>
        )}

        {this.state.submitError && <MessageBar messageBarType={MessageBarType.error}>{this.state.submitError}</MessageBar>}
        {this.state.submitSuccess && <MessageBar messageBarType={MessageBarType.success}>{this.state.submitSuccess}</MessageBar>}
        {(this.state.permissionStatusMessage && permissionDeniedForCurrentMode) && (
          <MessageBar messageBarType={MessageBarType.warning}>{this.state.permissionStatusMessage}</MessageBar>
        )}

        {schema.showTitle !== false && (schema.name || schema.description || schema.logoUrl) && <div style={{ display: 'flex', alignItems: 'flex-start', gap: '16px', width: '100%' }}>
          <div style={{ flex: '1 1 auto', minWidth: 0 }}>
            {schema.name && <h1 style={formTitleStyle}>{schema.name}</h1>}
            {schema.description && <p style={{
              fontSize: (schema.theme && schema.theme.descriptionFontSize) || 14,
              fontFamily: (schema.theme && schema.theme.descriptionFontFamily) || DEFAULT_THEME_FONT_FAMILY,
              fontWeight: ((schema.theme && schema.theme.descriptionFontWeight) || 'normal') as any,
              color: (schema.theme && schema.theme.descriptionColor) || '#666666',
              textAlign: schema.descriptionAlignment || 'left'
            }}>{schema.description}</p>}
          </div>
          {schema.logoUrl && <img src={schema.logoUrl} alt={schema.logoAltText || ''} style={{ flex: '0 0 auto', maxWidth: '160px', maxHeight: '80px', objectFit: 'contain' }} />}
        </div>}

        {(() => {
          var formGridLayout = !!(schema.theme && schema.theme.layout === 'grid' && schema.theme.columns && schema.theme.columns > 1);
          var formColumns = formGridLayout ? (schema.theme.columns || 2) : 1;
          var stepsContent = schema.steps.map((step) => {
            if (step.visible === false) {
              return null;
            }

            var gridLayout = isGridLayout(schema, step);
          var stepColumns = getStepColumns(schema, step);
          
          // Build step container styles from step-level theme
          var stepContainerStyle: React.CSSProperties = gridLayout
            ? {
              display: 'grid',
              gridTemplateColumns: 'repeat(' + String(stepColumns) + ', minmax(0, 1fr))',
              gap: '16px',
              marginBottom: '24px',
            }
            : { marginBottom: '24px' };
          var stepTextStyle: React.CSSProperties = {};
          var stepDescriptionStyle: React.CSSProperties = {};

          // Font settings apply to the container's own title/description; always resolved (own
          // default when unset) so the container never inherits form-level font styling.
          stepTextStyle.fontSize = (step.theme && step.theme.fontSize) || 14;
          stepTextStyle.fontFamily = (step.theme && step.theme.fontFamily) || DEFAULT_THEME_FONT_FAMILY;
          stepTextStyle.fontWeight = ((step.theme && step.theme.fontWeight) || 'normal') as any;
          stepTextStyle.color = (step.theme && step.theme.color) || '#000000';
          stepDescriptionStyle.fontSize = (step.theme && step.theme.descriptionFontSize) || 14;
          stepDescriptionStyle.fontFamily = (step.theme && step.theme.descriptionFontFamily) || DEFAULT_THEME_FONT_FAMILY;
          stepDescriptionStyle.fontWeight = ((step.theme && step.theme.descriptionFontWeight) || 'normal') as any;
          stepDescriptionStyle.color = (step.theme && step.theme.descriptionColor) || '#666666';
          if (step.theme) {
            // Background and border settings for step container
            if (step.theme.backgroundColor) { stepContainerStyle.backgroundColor = step.theme.backgroundColor; }
            if (step.theme.borderColor || step.theme.borderWidth) {
              stepContainerStyle.border = (step.theme.borderWidth || 1) + 'px solid ' + (step.theme.borderColor || '#cccccc');
            }
            if (step.theme.borderStyle === 'rounded') {
              stepContainerStyle.borderRadius = String(step.theme.borderRadius || 8) + 'px';
            }
          }

          // Add padding if background or border is set
          if (step.theme && (step.theme.backgroundColor || step.theme.borderColor || step.theme.borderWidth)) {
            stepContainerStyle.padding = '16px';
          }

          return (
          <div key={step.id}>
            {step.showTitle !== false && step.title && <h2 style={Object.assign({}, stepTextStyle, { textAlign: step.titleAlignment || 'left' })}>{step.title}</h2>}
            {step.showTitle !== false && step.description && <p style={Object.assign({}, stepDescriptionStyle, { textAlign: step.descriptionAlignment || 'left' })}>{step.description}</p>}
            <div style={stepContainerStyle}>
            {step.fields.map((field) => {
              if (!field) {
                return null;
              }
              return (
                <div key={field.id} style={gridLayout ? (field.startNewRow === true ? { gridColumn: '1 / -1' } : { gridColumn: 'span ' + String(Math.max(1, field.columnSpan || 1)) }) : { marginBottom: '16px' }}>
                  {this.renderField(field)}
                </div>
              );
            })}
            </div>
          </div>
        );});

          if (formGridLayout) {
            return (
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(' + String(formColumns) + ', minmax(0, 1fr))', gap: '24px' }}>
                {stepsContent}
              </div>
            );
          }
          return <div>{stepsContent}</div>;
        })()}        {(this.props.mode === 'view' && this.state.resolvedItemId > 0 && !this.state.isViewEditing) && (
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '24px' }}>
            {showViewBackButton && (
              <DefaultButton
                styles={defaultButtonStyles as any}
                onClick={this.handleCancel}
                disabled={this.state.isSubmitting}
              >
                {strings.DesignerBackButton}
              </DefaultButton>
            )}
          </div>
        )}

        {canSubmit && (
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '24px' }}>
            {this.props.showCancelButton !== false && (
              <DefaultButton styles={defaultButtonStyles as any} onClick={this.handleCancel} disabled={this.state.isSubmitting}>{cancelLabel}</DefaultButton>
            )}
            <PrimaryButton styles={primaryButtonStyles as any} onClick={this.handleSubmit} disabled={disableSubmitButton}>
              {this.state.isSubmitting ? strings.CommonSubmitting : submitLabel}
            </PrimaryButton>
          </div>
        )}
      </div>
    );
  }
}
