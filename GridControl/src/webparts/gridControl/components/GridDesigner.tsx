import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { evaluateGridValidationExpression, IGridAdvancedValidationRule, IGridValidationField } from './GridValidation';
import './GridDesigner.css';

export interface IGridDesignerProps {
  context: any;
  listName: string;
  viewId: string;
  schemaJson: string;
  onSave: (schemaJson: string) => void;
  onCancel: () => void;
}

export interface ISharePointField {
  internalName: string;
  title: string;
  description: string;
  type: string;
  required: boolean;
  readOnly: boolean;
  choices: string[];
  maxLength?: number;
  displayFormat?: string;
}

export interface IGridDesignerField {
  id: string;
  fieldName: string;
  sharePointType?: string;
  sharePointRequired?: boolean;
  sharePointDescription?: string;
  type: string;
  label: string;
  description?: string;
  visible?: boolean;
  required?: boolean;
  requiredMessage?: string;
  readOnly?: boolean;
  disabled?: boolean;
  defaultValue?: any;
  config?: any;
  validation?: any[];
  gridWidth?: string;
}

export interface IGridGroupingConfig {
  field1: string;
  field2: string;
  collapsedByDefault: boolean;
  showCount: boolean;
}

export interface IGridDesignerState {
  fields: IGridDesignerField[];
  sharePointFields: ISharePointField[];
  selectedFieldId: string;
  loading: boolean;
  error: string;
  grouping: IGridGroupingConfig;
  viewLoading: boolean;
  viewLoadMessage: string;
  advancedValidationEnabled: boolean;
  advancedValidationRules: IGridAdvancedValidationRule[];
  validationExpression: string;
  validationMessage: string;
  validationTargetField: string;
  selectedValidationRuleIndex: string;
  validationDesignerMessage: string;
}

var SYSTEM_FIELDS: { [name: string]: boolean } = {
  ID: true,
  GUID: true,
  ContentType: true,
  AppAuthor: true,
  AppEditor: true,
  Edit: true,
  ItemChildCount: true,
  FolderChildCount: true,
  ComplianceAssetId: true
};

function createId(): string {
  return 'gridfield_' + new Date().getTime().toString(36) + '_' + Math.floor(Math.random() * 100000).toString(36);
}

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

function trimGuidBraces(value: string): string {
  return String(value || '').replace(/^[{]/, '').replace(/[}]$/, '');
}

function buildViewIdCandidates(selectedViewId: string): string[] {
  var normalized = trimGuidBraces(selectedViewId);
  var candidates = [String(selectedViewId || ''), normalized, '{' + normalized + '}'];
  var unique: string[] = [];
  for (var i = 0; i < candidates.length; i += 1) {
    var value = String(candidates[i] || '');
    if (value && unique.indexOf(value) < 0) {
      unique.push(value);
    }
  }
  return unique;
}

function parseGroupByFieldNames(viewQuery: string): { fieldNames: string[]; collapsed: boolean } {
  var queryText = String(viewQuery || '').trim();
  if (!queryText) { return { fieldNames: [], collapsed: false }; }
  try {
    var xmlDocument = new DOMParser().parseFromString(
      /^<Query(?:\s|>)/i.test(queryText) ? queryText : '<Query>' + queryText + '</Query>',
      'text/xml'
    );
    if (xmlDocument.getElementsByTagName('parsererror').length > 0) { return { fieldNames: [], collapsed: false }; }
    var groupByElements = xmlDocument.getElementsByTagName('GroupBy');
    if (groupByElements.length === 0) { return { fieldNames: [], collapsed: false }; }
    var groupByElement = groupByElements[0];
    var fieldRefs = groupByElement.getElementsByTagName('FieldRef');
    var fieldNames: string[] = [];
    for (var i = 0; i < fieldRefs.length && fieldNames.length < 2; i += 1) {
      var name = fieldRefs[i].getAttribute('Name');
      if (name) { fieldNames.push(name); }
    }
    var collapsed = String(groupByElement.getAttribute('Collapse') || '').toUpperCase() === 'TRUE';
    return { fieldNames: fieldNames, collapsed: collapsed };
  } catch (_error) {
    return { fieldNames: [], collapsed: false };
  }
}

function copyFields(fields: IGridDesignerField[]): IGridDesignerField[] {
  return JSON.parse(JSON.stringify(fields || []));
}

function toArray(value: any): any[] {
  if (Array.isArray(value)) {
    return value;
  }
  if (value && Array.isArray(value.results)) {
    return value.results;
  }
  return [];
}

function mapFieldType(type: string): string {
  switch (type) {
    case 'Note': return 'multiline';
    case 'Number':
    case 'Currency':
    case 'Integer': return 'number';
    case 'DateTime': return 'datetime';
    case 'Choice': return 'dropdown';
    case 'MultiChoice': return 'multiselect';
    case 'Lookup':
    case 'LookupMulti': return 'lookup';
    case 'User':
    case 'UserMulti': return 'person';
    case 'Boolean': return 'boolean';
    case 'URL':
    case 'Hyperlink': return 'url';
    case 'Image':
    case 'Thumbnail': return 'image';
    case 'TaxonomyFieldType':
    case 'TaxonomyFieldTypeMulti': return 'taxonomy';
    case 'Attachments': return 'attachment';
    default: return 'text';
  }
}

function getCompatibleControlTypes(sharePointType: string): string[] {
  switch (sharePointType) {
    case 'Text':
    case 'Note':
      return ['text', 'multiline', 'number'];
    case 'Number':
    case 'Currency':
    case 'Integer':
      return ['number'];
    case 'Choice': return ['dropdown'];
    case 'MultiChoice': return ['multiselect'];
    case 'DateTime': return ['datetime'];
    case 'Boolean': return ['boolean'];
    case 'URL':
    case 'Hyperlink':
      return ['url', 'text'];
    case 'Lookup':
    case 'LookupMulti':
      return ['lookup'];
    case 'User':
    case 'UserMulti':
      return ['person'];
    case 'TaxonomyFieldType':
    case 'TaxonomyFieldTypeMulti':
      return ['taxonomy'];
    case 'Image':
    case 'Thumbnail':
      return ['image'];
    case 'Attachments': return ['attachment'];
    default: return [mapFieldType(sharePointType)];
  }
}

function isCompatibleControlType(sharePointType: string, controlType: string): boolean {
  return getCompatibleControlTypes(sharePointType).indexOf(controlType) >= 0;
}

function getTypeLabel(type: string): string {
  var labels: { [type: string]: string } = {
    text: 'Text', multiline: 'Multiline text', number: 'Number', datetime: 'Date and time',
    dropdown: 'Choice', multiselect: 'Multiple choice', lookup: 'Lookup', person: 'Person',
    boolean: 'Yes/No', url: 'URL', image: 'Image', taxonomy: 'Managed metadata', attachment: 'Attachments'
  };
  return labels[type] || type;
}

function parseSchemaFields(schemaJson: string): IGridDesignerField[] {
  if (!String(schemaJson || '').trim()) {
    return [];
  }
  try {
    var schema = JSON.parse(schemaJson);
    var steps = schema && Array.isArray(schema.steps) ? schema.steps : [];
    var fields: IGridDesignerField[] = [];
    for (var stepIndex = 0; stepIndex < steps.length; stepIndex += 1) {
      var stepFields = steps[stepIndex] && Array.isArray(steps[stepIndex].fields) ? steps[stepIndex].fields : [];
      for (var fieldIndex = 0; fieldIndex < stepFields.length; fieldIndex += 1) {
        if (stepFields[fieldIndex] && stepFields[fieldIndex].fieldName) {
          var parsedField = stepFields[fieldIndex] as IGridDesignerField;
          if (parsedField.disabled === true) {
            parsedField.readOnly = true;
            delete parsedField.disabled;
          }
          fields.push(parsedField);
        }
      }
    }
    return fields;
  } catch (_error) {
    return [];
  }
}

function parseGrouping(schemaJson: string): IGridGroupingConfig {
  var defaultGrouping: IGridGroupingConfig = { field1: '', field2: '', collapsedByDefault: false, showCount: true };
  if (!String(schemaJson || '').trim()) { return defaultGrouping; }
  try {
    var schema = JSON.parse(schemaJson);
    var grouping = schema && schema.grouping;
    if (!grouping) { return defaultGrouping; }
    return {
      field1: String(grouping.field1 || ''),
      field2: String(grouping.field2 || ''),
      collapsedByDefault: grouping.collapsedByDefault === true,
      showCount: grouping.showCount !== false
    };
  } catch (_error) {
    return defaultGrouping;
  }
}

function parseAdvancedValidation(schemaJson: string): { enabled: boolean; rules: IGridAdvancedValidationRule[] } {
  if (!String(schemaJson || '').trim()) { return { enabled: false, rules: [] }; }
  try {
    var schema = JSON.parse(schemaJson);
    var config = schema && schema.advancedValidation;
    return {
      enabled: !!(config && config.enabled === true),
      rules: config && Array.isArray(config.rules) ? config.rules : []
    };
  } catch (_error) {
    return { enabled: false, rules: [] };
  }
}

export class GridDesigner extends React.Component<IGridDesignerProps, IGridDesignerState> {
  public constructor(props: IGridDesignerProps) {
    super(props);
    var initialFields = parseSchemaFields(props.schemaJson);
    var initialAdvancedValidation = parseAdvancedValidation(props.schemaJson);
    this.state = {
      fields: initialFields,
      sharePointFields: [],
      selectedFieldId: initialFields.length > 0 ? initialFields[0].id : '',
      loading: false,
      error: '',
      grouping: parseGrouping(props.schemaJson),
      viewLoading: false,
      viewLoadMessage: '',
      advancedValidationEnabled: initialAdvancedValidation.enabled,
      advancedValidationRules: initialAdvancedValidation.rules,
      validationExpression: '',
      validationMessage: '',
      validationTargetField: '',
      selectedValidationRuleIndex: '',
      validationDesignerMessage: ''
    };
  }

  public componentDidMount(): void {
    this.loadFields();
  }

  private getWebUrl(): string {
    return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
  }

  private async getFieldsResponse(url: string): Promise<any> {
    var acceptHeaders = [
      '',
      'application/json;odata=verbose',
      'application/json;odata=minimalmetadata',
      'application/json;odata=nometadata'
    ];
    var response: any;
    for (var i = 0; i < acceptHeaders.length; i += 1) {
      response = await this.props.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        acceptHeaders[i] ? { headers: { Accept: acceptHeaders[i] } } : undefined
      );
      if (response.ok) {
        return response;
      }
    }
    return response;
  }

  private async loadFields(): Promise<void> {
    if (!this.props.listName) {
      this.setState({ error: 'Select a SharePoint list before opening Grid Designer.' });
      return;
    }
    this.setState({ loading: true, error: '' });
    try {
      var fieldsEndpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=";
      var url = fieldsEndpoint
        + 'InternalName,Title,Description,TypeAsString,Required,ReadOnlyField,Hidden,FromBaseType,Choices,MaxLength,DisplayFormat';
      var response = await this.getFieldsResponse(url);
      if (!response.ok) {
        url = fieldsEndpoint + 'InternalName,Title,Description,TypeAsString,Required,ReadOnlyField,Hidden,Choices,DisplayFormat';
        response = await this.getFieldsResponse(url);
      }
      if (!response.ok) {
        throw new Error('SharePoint fields could not be loaded. HTTP ' + String(response.status) + ' ' + String(response.statusText || ''));
      }
      var data = await response.json();
      var sourceFields = toArray(data && data.value).length > 0 ? toArray(data.value) : toArray(data && data.d && data.d.results);
      var fields = sourceFields.filter(function(field: any) {
        var isAttachment = field.InternalName === 'Attachments';
        var isCommonSystemField = field.InternalName === 'Author' || field.InternalName === 'Editor'
          || field.InternalName === 'Created' || field.InternalName === 'Modified';
        return !field.Hidden && !SYSTEM_FIELDS[field.InternalName]
          && (!field.FromBaseType || field.InternalName === 'Title' || isAttachment || isCommonSystemField);
      }).map(function(field: any) {
        return {
          internalName: String(field.InternalName || ''),
          title: String(field.Title || field.InternalName || ''),
          description: String(field.Description || ''),
          type: String(field.TypeAsString || 'Text'),
          required: field.Required === true,
          readOnly: field.ReadOnlyField === true,
          choices: toArray(field.Choices).map(function(choice: any) { return String(choice); }),
          maxLength: field.MaxLength,
          displayFormat: String(field.DisplayFormat) === '0' ? 'dateOnly' : 'dateTime'
        } as ISharePointField;
      });
      var configuredFields = copyFields(this.state.fields);
      for (var configuredIndex = 0; configuredIndex < configuredFields.length; configuredIndex += 1) {
        var configuredField = configuredFields[configuredIndex];
        for (var sourceIndex = 0; sourceIndex < fields.length; sourceIndex += 1) {
          if (fields[sourceIndex].internalName.toLowerCase() === String(configuredField.fieldName || '').toLowerCase()) {
            configuredField.sharePointType = fields[sourceIndex].type;
            configuredField.sharePointRequired = fields[sourceIndex].required;
            configuredField.sharePointDescription = fields[sourceIndex].description;
            if (fields[sourceIndex].required) {
              configuredField.required = true;
            }
            if (!isCompatibleControlType(fields[sourceIndex].type, configuredField.type)) {
              configuredField.type = mapFieldType(fields[sourceIndex].type);
            }
            break;
          }
        }
      }
      this.setState({ sharePointFields: fields, fields: configuredFields, loading: false });
    } catch (error) {
      this.setState({ loading: false, error: error && error.message ? error.message : 'SharePoint fields could not be loaded.' });
    }
  }

  private async loadFromView(): Promise<void> {
    if (!this.props.listName || !this.props.viewId) {
      this.setState({ viewLoadMessage: 'Select a SharePoint view in the web part properties first.' });
      return;
    }
    if (this.state.fields.length > 0 && typeof window !== 'undefined'
      && !window.confirm('Replace the current grid columns and grouping with the selected view\'s fields and grouping?')) {
      return;
    }

    this.setState({ viewLoading: true, viewLoadMessage: '' });
    try {
      var webUrl = this.getWebUrl();
      var listPath = "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')";
      var viewIds = buildViewIdCandidates(this.props.viewId);
      var viewQuery = '';
      var viewFieldNames: string[] = [];

      for (var i = 0; i < viewIds.length && !viewQuery; i += 1) {
        var encoded = encodeURIComponent(viewIds[i]);
        var normalized = encodeURIComponent(trimGuidBraces(viewIds[i]));
        var queryUrls = [
          webUrl + listPath + "/views/getById('" + encoded + "')?$select=ViewQuery",
          webUrl + listPath + "/views(guid'" + normalized + "')?$select=ViewQuery"
        ];
        for (var q = 0; q < queryUrls.length && !viewQuery; q += 1) {
          var queryResponse = await this.getFieldsResponse(queryUrls[q]);
          if (!queryResponse.ok) { continue; }
          var queryData = await queryResponse.json();
          var queryContainer = queryData && queryData.d ? queryData.d : queryData;
          if (queryContainer && queryContainer.ViewQuery !== undefined && queryContainer.ViewQuery !== null) {
            viewQuery = String(queryContainer.ViewQuery);
          }
        }

        var fieldsUrls = [
          webUrl + listPath + "/views/getById('" + encoded + "')/ViewFields",
          webUrl + listPath + "/views(guid'" + normalized + "')/ViewFields"
        ];
        for (var f = 0; f < fieldsUrls.length && viewFieldNames.length === 0; f += 1) {
          var fieldsResponse = await this.getFieldsResponse(fieldsUrls[f]);
          if (!fieldsResponse.ok) { continue; }
          var fieldsData = await fieldsResponse.json();
          var names = toArray(fieldsData.value);
          if (names.length === 0) { names = toArray(fieldsData.Items); }
          if (names.length === 0) { names = toArray(fieldsData && fieldsData.d && fieldsData.d.Items); }
          if (names.length > 0) {
            viewFieldNames = names.map(function(name: any) { return String(name || ''); }).filter(function(name: string) { return !!name; });
          }
        }
      }

      var groupByResult = parseGroupByFieldNames(viewQuery);

      // Only replace columns when the view actually returned an ordered field list; grouping is always applied.
      if (viewFieldNames.length > 0) {
        var nextFields: IGridDesignerField[] = [];
        var addedFieldNames: { [name: string]: boolean } = {};
        for (var v = 0; v < viewFieldNames.length; v += 1) {
          var viewFieldName = viewFieldNames[v];
          if (SYSTEM_FIELDS[viewFieldName]) { continue; }
          var matchedSharePointField: ISharePointField | undefined;
          for (var s = 0; s < this.state.sharePointFields.length; s += 1) {
            if (this.state.sharePointFields[s].internalName.toLowerCase() === viewFieldName.toLowerCase()) {
              matchedSharePointField = this.state.sharePointFields[s];
              break;
            }
          }
          if (!matchedSharePointField) { continue; }
          // A SharePoint view can list the same field more than once; only add it once.
          if (addedFieldNames[matchedSharePointField.internalName.toLowerCase()]) { continue; }
          addedFieldNames[matchedSharePointField.internalName.toLowerCase()] = true;
          var config: any = {};
          if (matchedSharePointField.choices.length > 0) { config.choices = matchedSharePointField.choices.slice(); }
          if (matchedSharePointField.maxLength) { config.maxLength = matchedSharePointField.maxLength; }
          if (matchedSharePointField.type === 'DateTime') { config.displayFormat = matchedSharePointField.displayFormat || 'dateTime'; }
          nextFields.push({
            id: createId(),
            fieldName: matchedSharePointField.internalName,
            sharePointType: matchedSharePointField.type,
            sharePointRequired: matchedSharePointField.required,
            sharePointDescription: matchedSharePointField.description,
            type: mapFieldType(matchedSharePointField.type),
            label: matchedSharePointField.title,
            description: '',
            visible: true,
            required: matchedSharePointField.required,
            readOnly: matchedSharePointField.readOnly,
            config: config
          });
        }
        if (nextFields.length > 0) {
          this.setState({ fields: nextFields, selectedFieldId: nextFields[0].id });
        }
      }

      var groupField1 = groupByResult.fieldNames.length > 0 ? groupByResult.fieldNames[0] : '';
      var groupField2 = groupByResult.fieldNames.length > 1 ? groupByResult.fieldNames[1] : '';
      this.setState({
        viewLoading: false,
        grouping: {
          field1: groupField1,
          field2: groupField2,
          collapsedByDefault: groupByResult.collapsed,
          showCount: this.state.grouping.showCount
        },
        viewLoadMessage: viewFieldNames.length > 0
          ? (groupField1 ? 'Loaded columns and grouping from the selected view.' : 'Loaded columns from the selected view. The view has no grouping configured.')
          : (groupField1 ? 'Loaded grouping from the selected view.' : 'The selected view has no fields or grouping to load.')
      });
    } catch (error) {
      this.setState({ viewLoading: false, viewLoadMessage: error && error.message ? error.message : 'Failed to load the selected view.' });
    }
  }

  private getAvailableFields(): ISharePointField[] {
    var used: { [name: string]: boolean } = {};
    this.state.fields.forEach(function(field) { used[String(field.fieldName || '').toLowerCase()] = true; });
    return this.state.sharePointFields.filter(function(field) { return !used[field.internalName.toLowerCase()]; });
  }

  private addField(source: ISharePointField): void {
    var config: any = {};
    if (source.choices.length > 0) { config.choices = source.choices.slice(); }
    if (source.maxLength) { config.maxLength = source.maxLength; }
    if (source.type === 'DateTime') { config.displayFormat = source.displayFormat || 'dateTime'; }
    var field: IGridDesignerField = {
      id: createId(),
      fieldName: source.internalName,
      sharePointType: source.type,
      sharePointRequired: source.required,
      sharePointDescription: source.description,
      type: mapFieldType(source.type),
      label: source.title,
      description: '',
      visible: true,
      required: source.required,
      readOnly: source.readOnly,
      config: config
    };
    var fields = copyFields(this.state.fields);
    fields.push(field);
    this.setState({ fields: fields, selectedFieldId: field.id });
  }

  private getSelectedField(): IGridDesignerField | undefined {
    for (var i = 0; i < this.state.fields.length; i += 1) {
      if (this.state.fields[i].id === this.state.selectedFieldId) {
        return this.state.fields[i];
      }
    }
    return undefined;
  }

  private updateSelectedField(mutator: (field: IGridDesignerField) => void): void {
    var fields = copyFields(this.state.fields);
    for (var i = 0; i < fields.length; i += 1) {
      if (fields[i].id === this.state.selectedFieldId) {
        mutator(fields[i]);
        break;
      }
    }
    this.setState({ fields: fields });
  }

  private moveField(fieldId: string, direction: number): void {
    var fields = copyFields(this.state.fields);
    var index = -1;
    for (var i = 0; i < fields.length; i += 1) {
      if (fields[i].id === fieldId) { index = i; break; }
    }
    var target = index + direction;
    if (index < 0 || target < 0 || target >= fields.length) { return; }
    var moved = fields[index];
    fields[index] = fields[target];
    fields[target] = moved;
    this.setState({ fields: fields });
  }

  private removeField(fieldId: string): void {
    var fields = this.state.fields.filter(function(field) { return field.id !== fieldId; });
    this.setState({
      fields: fields,
      selectedFieldId: this.state.selectedFieldId === fieldId ? (fields.length > 0 ? fields[0].id : '') : this.state.selectedFieldId
    });
  }

  private updateConfig(name: string, value: any): void {
    this.updateSelectedField(function(field) {
      field.config = field.config || {};
      field.config[name] = value;
    });
  }

  private updateGrouping(mutator: (grouping: IGridGroupingConfig) => void): void {
    var grouping: IGridGroupingConfig = {
      field1: this.state.grouping.field1,
      field2: this.state.grouping.field2,
      collapsedByDefault: this.state.grouping.collapsedByDefault,
      showCount: this.state.grouping.showCount
    };
    mutator(grouping);
    if (!grouping.field1) { grouping.field2 = ''; }
    this.setState({ grouping: grouping });
  }

  private changeSelectedControlType(controlType: string): void {
    this.updateSelectedField(function(field) {
      field.type = controlType;
      var allowedValidationTypes = controlType === 'number' ? ['min', 'max']
        : (controlType === 'text' || controlType === 'multiline' || controlType === 'url')
          ? ['minLength', 'maxLength', 'pattern'] : [];
      field.validation = (field.validation || []).filter(function(rule: any) {
        return allowedValidationTypes.indexOf(rule.type) >= 0;
      });
      field.config = field.config || {};
      if (controlType === 'number') {
        delete field.config.maxLength;
        delete field.config.choices;
        delete field.config.displayFormat;
      } else if (controlType === 'text' || controlType === 'multiline' || controlType === 'url') {
        delete field.config.min;
        delete field.config.max;
        delete field.config.decimals;
        delete field.config.choices;
        delete field.config.displayFormat;
      }
    });
  }

  private updateValidation(type: string, value: string, message: string): void {
    this.updateSelectedField(function(field) {
      field.validation = (field.validation || []).filter(function(rule: any) { return rule.type !== type; });
      if (String(value || '').trim()) {
        field.validation.push({ type: type, value: value, message: message || 'The value is invalid.' });
      }
    });
  }

  private updateValidationMessages(types: string[], message: string): void {
    this.updateSelectedField(function(field) {
      var rules = field.validation || [];
      for (var ruleIndex = 0; ruleIndex < rules.length; ruleIndex += 1) {
        if (types.indexOf(rules[ruleIndex].type) >= 0) {
          rules[ruleIndex].message = message || 'The value is invalid.';
        }
      }
      field.validation = rules;
    });
  }

  private getValidation(field: IGridDesignerField, type: string): any {
    var rules = field.validation || [];
    for (var i = 0; i < rules.length; i += 1) {
      if (rules[i].type === type) { return rules[i]; }
    }
    return {};
  }

  private getPatternError(pattern: string): string {
    if (!String(pattern || '').trim()) { return ''; }
    try { new RegExp(pattern); return ''; } catch (error) {
      return error && error.message ? error.message : 'The regular expression is invalid.';
    }
  }

  private getValidationFields(): IGridValidationField[] {
    return this.state.fields.map(function(field) {
      return { id: field.id, fieldName: field.fieldName, label: field.label };
    });
  }

  private validateAdvancedExpression(expression: string): string {
    try {
      evaluateGridValidationExpression(expression, this.getValidationFields(), {});
      return '';
    } catch (error) {
      return error && error.message ? error.message : 'The validation expression is invalid.';
    }
  }

  private appendValidationExpression(fragment: string): void {
    var current = String(this.state.validationExpression || '').trim();
    this.setState({ validationExpression: current ? current + ' ' + fragment : fragment, validationDesignerMessage: '' });
  }

  private loadValidationRule(indexValue: string): void {
    var index = parseInt(indexValue, 10);
    var rule = !isNaN(index) ? this.state.advancedValidationRules[index] : undefined;
    this.setState({
      selectedValidationRuleIndex: indexValue,
      validationExpression: rule ? rule.expression : '',
      validationMessage: rule ? rule.message : '',
      validationTargetField: rule && rule.targetField ? rule.targetField : '',
      validationDesignerMessage: ''
    });
  }

  private saveValidationRule(updateExisting: boolean): void {
    var expression = String(this.state.validationExpression || '').trim();
    var message = String(this.state.validationMessage || '').trim();
    if (!expression || !message) {
      this.setState({ validationDesignerMessage: 'Enter both an expression and validation message.' });
      return;
    }
    var expressionError = this.validateAdvancedExpression(expression);
    if (expressionError) {
      this.setState({ validationDesignerMessage: 'Expression error: ' + expressionError });
      return;
    }
    var rules = this.state.advancedValidationRules.slice(0);
    var selectedIndex = parseInt(this.state.selectedValidationRuleIndex, 10);
    var rule: IGridAdvancedValidationRule = {
      id: !isNaN(selectedIndex) && rules[selectedIndex] && rules[selectedIndex].id
        ? rules[selectedIndex].id : createId(),
      expression: expression,
      message: message,
      targetField: this.state.validationTargetField || undefined
    };
    if (updateExisting && !isNaN(selectedIndex) && rules[selectedIndex]) {
      rules[selectedIndex] = rule;
    } else {
      rules.push(rule);
      selectedIndex = rules.length - 1;
    }
    this.setState({
      advancedValidationRules: rules,
      advancedValidationEnabled: true,
      selectedValidationRuleIndex: String(selectedIndex),
      validationDesignerMessage: updateExisting ? 'Validation rule updated.' : 'Validation rule added.'
    });
  }

  private removeValidationRule(): void {
    var index = parseInt(this.state.selectedValidationRuleIndex, 10);
    if (isNaN(index) || !this.state.advancedValidationRules[index]) { return; }
    var rules = this.state.advancedValidationRules.slice(0);
    rules.splice(index, 1);
    this.setState({
      advancedValidationRules: rules,
      advancedValidationEnabled: rules.length > 0 && this.state.advancedValidationEnabled,
      selectedValidationRuleIndex: '',
      validationExpression: '',
      validationMessage: '',
      validationTargetField: '',
      validationDesignerMessage: 'Validation rule removed.'
    });
  }

  private save(): void {
    if (this.state.loading) {
      this.setState({ validationDesignerMessage: 'Wait for SharePoint field metadata to finish loading.' });
      return;
    }
    var fieldsToSave = copyFields(this.state.fields);
    for (var fieldIndex = 0; fieldIndex < fieldsToSave.length; fieldIndex += 1) {
      for (var sourceIndex = 0; sourceIndex < this.state.sharePointFields.length; sourceIndex += 1) {
        if (this.state.sharePointFields[sourceIndex].internalName.toLowerCase() === String(fieldsToSave[fieldIndex].fieldName || '').toLowerCase()
          && this.state.sharePointFields[sourceIndex].required) {
          fieldsToSave[fieldIndex].required = true;
          fieldsToSave[fieldIndex].sharePointRequired = true;
          break;
        }
      }
    }
    for (var fieldIndex = 0; fieldIndex < fieldsToSave.length; fieldIndex += 1) {
      var patternRule = this.getValidation(fieldsToSave[fieldIndex], 'pattern');
      var patternError = this.getPatternError(patternRule.value || '');
      if (patternError) {
        this.setState({ selectedFieldId: fieldsToSave[fieldIndex].id, validationDesignerMessage: 'Fix the invalid pattern before saving: ' + patternError });
        return;
      }
    }
    for (var ruleIndex = 0; ruleIndex < this.state.advancedValidationRules.length; ruleIndex += 1) {
      var expressionError = this.validateAdvancedExpression(this.state.advancedValidationRules[ruleIndex].expression);
      if (expressionError) {
        this.setState({ selectedValidationRuleIndex: String(ruleIndex), validationDesignerMessage: 'Fix validation rule ' + String(ruleIndex + 1) + ': ' + expressionError });
        return;
      }
    }
    var schema = {
      id: 'grid_' + String(this.props.listName || '').replace(/[^a-z0-9]/gi, '_').toLowerCase(),
      name: this.props.listName + ' Grid',
      mode: 'edit',
      listName: this.props.listName,
      grouping: {
        enabled: !!this.state.grouping.field1,
        field1: this.state.grouping.field1,
        field2: this.state.grouping.field1 ? this.state.grouping.field2 : '',
        collapsedByDefault: this.state.grouping.collapsedByDefault,
        showCount: this.state.grouping.showCount
      },
      advancedValidation: {
        enabled: this.state.advancedValidationEnabled && this.state.advancedValidationRules.length > 0,
        rules: this.state.advancedValidationRules
      },
      steps: [{ id: 'grid_columns', title: 'Grid columns', showTitle: false, visible: true, fields: fieldsToSave }]
    };
    this.props.onSave(JSON.stringify(schema));
  }

  private renderPalette(): JSX.Element {
    var available = this.getAvailableFields();
    return (
      <aside className="gd-panel gd-palette">
        <h3>SharePoint fields</h3>
        <p>Add fields as editable grid columns.</p>
        {this.state.loading && <div className="gd-empty">Loading fields...</div>}
        {this.state.error && <div className="gd-error">{this.state.error}</div>}
        {!this.state.loading && !this.state.error && available.length === 0 && <div className="gd-empty">All available fields are in the grid.</div>}
        {available.map((field) => (
          <button key={field.internalName} type="button" className="gd-palette-item" onClick={() => this.addField(field)}>
            <span>{field.title}</span>
            <small>{getTypeLabel(mapFieldType(field.type))}</small>
          </button>
        ))}
      </aside>
    );
  }

  private getFieldLabel(fieldName: string): string {
    if (!fieldName) { return ''; }
    for (var i = 0; i < this.state.fields.length; i += 1) {
      if (this.state.fields[i].fieldName.toLowerCase() === fieldName.toLowerCase()) {
        return this.state.fields[i].label || this.state.fields[i].fieldName;
      }
    }
    for (var j = 0; j < this.state.sharePointFields.length; j += 1) {
      if (this.state.sharePointFields[j].internalName.toLowerCase() === fieldName.toLowerCase()) {
        return this.state.sharePointFields[j].title || fieldName;
      }
    }
    return fieldName;
  }

  private renderGroupingSettings(): JSX.Element {
    var grouping = this.state.grouping;
    return (
      <div className="gd-group-settings">
        <div className="gd-canvas-header">
          <div><h3>Grouping</h3><p>Group grid rows by up to two columns, similar to a SharePoint view's Group By.</p></div>
        </div>
        {this.props.viewId && (
          <div className="gd-group-view-load">
            <button type="button" disabled={this.state.viewLoading} onClick={() => this.loadFromView()}>
              {this.state.viewLoading ? 'Loading from view…' : 'Load columns & grouping from selected view'}
            </button>
            {this.state.viewLoadMessage && <div className="gd-hint">{this.state.viewLoadMessage}</div>}
          </div>
        )}
        <div className="gd-inline-fields gd-inline-fields-two">
          <label>Group by<select value={grouping.field1} onChange={(ev) => { var value = ev.currentTarget.value; this.updateGrouping(function(next) { next.field1 = value; }); }}>
            <option value="">None</option>
            {this.state.fields.map(function(field) { return <option key={field.id} value={field.fieldName}>{field.label || field.fieldName}</option>; })}
          </select></label>
          <label>Then by<select value={grouping.field2} disabled={!grouping.field1} onChange={(ev) => { var value = ev.currentTarget.value; this.updateGrouping(function(next) { next.field2 = value; }); }}>
            <option value="">None</option>
            {this.state.fields.filter(function(field) { return field.fieldName !== grouping.field1; }).map(function(field) { return <option key={field.id} value={field.fieldName}>{field.label || field.fieldName}</option>; })}
          </select></label>
        </div>
        <label className="gd-check"><input type="checkbox" checked={grouping.collapsedByDefault} disabled={!grouping.field1} onChange={(ev) => { var checked = ev.currentTarget.checked; this.updateGrouping(function(next) { next.collapsedByDefault = checked; }); }} /> Collapse groups by default</label>
        <label className="gd-check"><input type="checkbox" checked={grouping.showCount} disabled={!grouping.field1} onChange={(ev) => { var checked = ev.currentTarget.checked; this.updateGrouping(function(next) { next.showCount = checked; }); }} /> Show item count per group</label>
      </div>
    );
  }

  private renderColumnsList(): JSX.Element {
    return (
      <div>
        {this.state.fields.length === 0 && <div className="gd-empty gd-empty-canvas">Add SharePoint fields from the left panel.</div>}
        {this.state.fields.map((field, index) => (
          <div key={field.id} className={'gd-column ' + (field.id === this.state.selectedFieldId ? 'gd-column-selected' : '')} onClick={() => this.setState({ selectedFieldId: field.id })}>
            <div className="gd-column-order">{index + 1}</div>
            <div className="gd-column-main">
              <strong>{field.label || field.fieldName}</strong>
              <span>{field.fieldName} · {getTypeLabel(field.type)}</span>
              <div className="gd-chips">
                {field.required === true && <em>Required</em>}
                {field.readOnly === true && <em>Read only</em>}
                {field.visible === false && <em>Hidden</em>}
              </div>
            </div>
            <div className="gd-column-actions">
              <button type="button" title="Move up" disabled={index === 0} onClick={(ev) => { ev.stopPropagation(); this.moveField(field.id, -1); }}>↑</button>
              <button type="button" title="Move down" disabled={index === this.state.fields.length - 1} onClick={(ev) => { ev.stopPropagation(); this.moveField(field.id, 1); }}>↓</button>
              <button type="button" title="Remove column" onClick={(ev) => { ev.stopPropagation(); this.removeField(field.id); }}>×</button>
            </div>
          </div>
        ))}
      </div>
    );
  }

  private renderCanvas(): JSX.Element {
    var grouping = this.state.grouping;
    var columnsPreview = this.renderColumnsList();
    if (grouping.field1) {
      var innerPreview = grouping.field2 ? (
        <div className="gd-group-container gd-group-container-nested">
          <div className="gd-group-container-header">
            <span className="gd-group-container-icon">▾</span>
            <span>Then by: {this.getFieldLabel(grouping.field2)} = <em>(example value)</em></span>
            {grouping.showCount && <span className="gd-group-container-count">3 items</span>}
          </div>
          <div className="gd-group-container-body">{columnsPreview}</div>
        </div>
      ) : columnsPreview;
      columnsPreview = (
        <div className="gd-group-container">
          <div className="gd-group-container-header">
            <span className="gd-group-container-icon">{grouping.collapsedByDefault ? '▸' : '▾'}</span>
            <span>Grouped by: {this.getFieldLabel(grouping.field1)} = <em>(example value)</em></span>
            {grouping.showCount && <span className="gd-group-container-count">{grouping.field2 ? '' : '5 items'}</span>}
          </div>
          <div className="gd-group-container-body">{innerPreview}</div>
        </div>
      );
    }
    return (
      <main className="gd-canvas">
        {this.renderGroupingSettings()}
        <div className="gd-canvas-header">
          <div><h3>Grid columns</h3><p>Column order follows this list from top to bottom.</p></div>
          <span>{this.state.fields.length} columns</span>
        </div>
        {columnsPreview}
      </main>
    );
  }

  private renderFieldEditor(): JSX.Element {
    var field = this.getSelectedField();
    if (!field) {
      return <aside className="gd-panel gd-properties"><div className="gd-empty">Select a grid column to configure it.</div></aside>;
    }
    var config = field.config || {};
    var compatibleControlTypes = getCompatibleControlTypes(field.sharePointType || 'Text');
    var minLengthRule = this.getValidation(field, 'minLength');
    var maxLengthRule = this.getValidation(field, 'maxLength');
    var minRule = this.getValidation(field, 'min');
    var maxRule = this.getValidation(field, 'max');
    var patternRule = this.getValidation(field, 'pattern');
    var patternError = this.getPatternError(patternRule.value || '');
    return (
      <aside className="gd-panel gd-properties">
        <h3>Column settings</h3>
        <label>Column label<input type="text" value={field.label || ''} onChange={(ev) => this.updateSelectedField(function(next) { next.label = ev.currentTarget.value; })} /></label>
        {field.sharePointDescription && <label>SharePoint description<div className="gd-readonly-value">{field.sharePointDescription}</div></label>}
        <label>{field.sharePointDescription ? 'Designer description (SharePoint description takes precedence)' : 'Description'}<textarea value={field.description || ''} onChange={(ev) => this.updateSelectedField(function(next) { next.description = ev.currentTarget.value; })} /></label>
        <label>SharePoint field type<div className="gd-readonly-value">{field.sharePointType || 'Unknown'}</div></label>
        <label>Cell control<select value={field.type} onChange={(ev) => this.changeSelectedControlType(ev.currentTarget.value)}>
          {compatibleControlTypes.map(function(controlType) {
            return <option key={controlType} value={controlType}>{getTypeLabel(controlType)}</option>;
          })}
        </select></label>
        <div className="gd-hint">Only controls that can be safely saved to this SharePoint field are available.</div>
        <label>Column width<input type="text" value={field.gridWidth || ''} placeholder="180px or 20%" onChange={(ev) => this.updateSelectedField(function(next) { next.gridWidth = ev.currentTarget.value; })} /></label>
        <label className="gd-check"><input type="checkbox" checked={field.visible !== false} onChange={(ev) => this.updateSelectedField(function(next) { next.visible = ev.currentTarget.checked; })} /> Visible</label>
        <label className="gd-check"><input type="checkbox" checked={field.required === true || field.sharePointRequired === true} disabled={field.sharePointRequired === true} onChange={(ev) => this.updateSelectedField(function(next) { next.required = ev.currentTarget.checked; })} /> {field.sharePointRequired === true ? 'Required by SharePoint' : 'Required'}</label>
        <label className="gd-check"><input type="checkbox" checked={field.readOnly === true} onChange={(ev) => this.updateSelectedField(function(next) { next.readOnly = ev.currentTarget.checked; })} /> Read only</label>
        {field.required === true && <label>Required message<input type="text" value={field.requiredMessage || ''} onChange={(ev) => this.updateSelectedField(function(next) { next.requiredMessage = ev.currentTarget.value; })} /></label>}
        <label>Placeholder<input type="text" value={config.placeholder || ''} onChange={(ev) => this.updateConfig('placeholder', ev.currentTarget.value)} /></label>
        <label>Default value<input type="text" value={field.defaultValue === undefined ? '' : String(field.defaultValue)} onChange={(ev) => this.updateSelectedField(function(next) { next.defaultValue = ev.currentTarget.value; })} /></label>
        {(field.type === 'text' || field.type === 'multiline') && <label>Maximum length<input type="number" min="0" value={config.maxLength || ''} onChange={(ev) => this.updateConfig('maxLength', parseInt(ev.currentTarget.value, 10) || undefined)} /></label>}
        {field.type === 'number' && <div className="gd-inline-fields"><label>Minimum<input type="number" value={config.min === undefined ? '' : config.min} onChange={(ev) => this.updateConfig('min', ev.currentTarget.value === '' ? undefined : Number(ev.currentTarget.value))} /></label><label>Maximum<input type="number" value={config.max === undefined ? '' : config.max} onChange={(ev) => this.updateConfig('max', ev.currentTarget.value === '' ? undefined : Number(ev.currentTarget.value))} /></label><label>Decimals<input type="number" min="0" max="10" value={config.decimals === undefined ? '' : config.decimals} onChange={(ev) => this.updateConfig('decimals', ev.currentTarget.value === '' ? undefined : Number(ev.currentTarget.value))} /></label></div>}
        {(field.type === 'dropdown' || field.type === 'multiselect') && <label>Choices, one per line<textarea value={(config.choices || []).join('\n')} onChange={(ev) => this.updateConfig('choices', ev.currentTarget.value.split(/\r?\n/).filter(function(choice) { return !!choice.trim(); }))} /></label>}
        {field.type === 'datetime' && <label>Date/time format<select value={config.displayFormat || 'dateTime'} onChange={(ev) => this.updateConfig('displayFormat', ev.currentTarget.value)}><option value="dateOnly">Date only</option><option value="dateTime">Date and time</option><option value="timeOnly">Time only</option></select></label>}
        <div className="gd-section-title">Validation</div>
        {(field.type === 'text' || field.type === 'multiline' || field.type === 'url') && (
          <div>
            <div className="gd-inline-fields gd-inline-fields-two">
              <label>Minimum length<input type="number" min="0" value={minLengthRule.value || ''} onChange={(ev) => this.updateValidation('minLength', ev.currentTarget.value, minLengthRule.message || 'The value is too short.')} /></label>
              <label>Maximum length<input type="number" min="0" value={maxLengthRule.value || ''} onChange={(ev) => this.updateValidation('maxLength', ev.currentTarget.value, maxLengthRule.message || 'The value is too long.')} /></label>
            </div>
            <label>Length validation message<input type="text" value={minLengthRule.message || maxLengthRule.message || ''} onChange={(ev) => this.updateValidationMessages(['minLength', 'maxLength'], ev.currentTarget.value)} /></label>
            <label>Pattern (optional)<input type="text" value={patternRule.value || ''} placeholder="Regular expression" onChange={(ev) => this.updateValidation('pattern', ev.currentTarget.value, patternRule.message || 'The value format is invalid.')} /></label>
            {patternError && <div className="gd-error gd-validation-error">{patternError}</div>}
            <label>Pattern message<input type="text" value={patternRule.message || ''} onChange={(ev) => this.updateValidation('pattern', patternRule.value || '', ev.currentTarget.value)} /></label>
          </div>
        )}
        {field.type === 'number' && (
          <div>
            <div className="gd-inline-fields gd-inline-fields-two">
              <label>Minimum value<input type="number" value={minRule.value || ''} onChange={(ev) => this.updateValidation('min', ev.currentTarget.value, minRule.message || 'The value is below the minimum.')} /></label>
              <label>Maximum value<input type="number" value={maxRule.value || ''} onChange={(ev) => this.updateValidation('max', ev.currentTarget.value, maxRule.message || 'The value exceeds the maximum.')} /></label>
            </div>
            <label>Range validation message<input type="text" value={minRule.message || maxRule.message || ''} onChange={(ev) => this.updateValidationMessages(['min', 'max'], ev.currentTarget.value)} /></label>
          </div>
        )}
        <div className="gd-section-title">Advanced validation</div>
        <label className="gd-check"><input type="checkbox" checked={this.state.advancedValidationEnabled} onChange={(ev) => this.setState({ advancedValidationEnabled: ev.currentTarget.checked })} /> Enable advanced validation rules</label>
        <div className="gd-hint">Uses the same expression grammar as Dynamic Form: ==, !=, &gt;, &gt;=, &lt;, &lt;=, &amp;&amp;, ||, !, and parentheses.</div>
        <div className="gd-validation-actions">
          <button type="button" onClick={() => this.appendValidationExpression('field("' + field.fieldName.replace(/"/g, '\\"') + '")')}>Insert selected field</button>
          <button type="button" onClick={() => this.appendValidationExpression('hasValue(field("' + field.fieldName.replace(/"/g, '\\"') + '"))')}>Has value</button>
          <button type="button" onClick={() => this.appendValidationExpression('isEmpty(field("' + field.fieldName.replace(/"/g, '\\"') + '"))')}>Is empty</button>
          <button type="button" onClick={() => this.appendValidationExpression('between(field("' + field.fieldName.replace(/"/g, '\\"') + '"), 1, 10)')}>Between</button>
          <button type="button" onClick={() => this.appendValidationExpression('today()')}>Today</button>
          <button type="button" onClick={() => this.appendValidationExpression('daysFromToday(7)')}>Days from today</button>
        </div>
        <label>Expression<textarea value={this.state.validationExpression} placeholder={'field("' + field.fieldName + '") == "value"'} onChange={(ev) => this.setState({ validationExpression: ev.currentTarget.value, validationDesignerMessage: '' })} /></label>
        <label>Message target<select value={this.state.validationTargetField} onChange={(ev) => this.setState({ validationTargetField: ev.currentTarget.value })}>
          <option value="">Grid-level message</option>
          {this.state.fields.map(function(targetField) { return <option key={targetField.id} value={targetField.fieldName}>{targetField.label || targetField.fieldName}</option>; })}
        </select></label>
        <label>Validation message<input type="text" value={this.state.validationMessage} onChange={(ev) => this.setState({ validationMessage: ev.currentTarget.value, validationDesignerMessage: '' })} /></label>
        <label>Existing rules<select value={this.state.selectedValidationRuleIndex} onChange={(ev) => this.loadValidationRule(ev.currentTarget.value)}>
          <option value="">Select a rule</option>
          {this.state.advancedValidationRules.map(function(rule, ruleIndex) { return <option key={String(ruleIndex)} value={String(ruleIndex)}>{String(ruleIndex + 1) + '. ' + rule.message}</option>; })}
        </select></label>
        <div className="gd-validation-actions">
          <button type="button" onClick={() => this.saveValidationRule(false)}>Add rule</button>
          <button type="button" disabled={this.state.selectedValidationRuleIndex === ''} onClick={() => this.saveValidationRule(true)}>Update selected</button>
          <button type="button" disabled={this.state.selectedValidationRuleIndex === ''} onClick={() => this.removeValidationRule()}>Remove selected</button>
          <button type="button" onClick={() => this.setState({ validationExpression: '', validationMessage: '', validationTargetField: '', selectedValidationRuleIndex: '', validationDesignerMessage: '' })}>Clear editor</button>
        </div>
        {this.state.validationDesignerMessage && <div className="gd-validation-message">{this.state.validationDesignerMessage}</div>}
      </aside>
    );
  }

  public render(): JSX.Element {
    return (
      <div className="gd-overlay">
        <header className="gd-toolbar">
          <div><h2>Grid Designer</h2><span>{this.props.listName || 'No list selected'}</span></div>
          <div className="gd-toolbar-actions"><button type="button" onClick={this.props.onCancel}>Cancel</button><button type="button" className="gd-primary" onClick={() => this.save()} disabled={!this.props.listName || this.state.loading}>Save grid</button></div>
        </header>
        <div className="gd-workspace">{this.renderPalette()}{this.renderCanvas()}{this.renderFieldEditor()}</div>
      </div>
    );
  }
}