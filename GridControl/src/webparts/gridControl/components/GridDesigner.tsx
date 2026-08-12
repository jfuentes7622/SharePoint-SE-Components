import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { evaluateGridValidationExpression, IGridAdvancedValidationRule, IGridValidationField } from './GridValidation';
import './GridDesigner.css';

export interface IGridDesignerProps {
  context: any;
  listName: string;
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

export interface IGridDesignerState {
  fields: IGridDesignerField[];
  sharePointFields: ISharePointField[];
  selectedFieldId: string;
  loading: boolean;
  error: string;
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
  Created: true,
  Modified: true,
  Author: true,
  Editor: true,
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

  private async loadFields(): Promise<void> {
    if (!this.props.listName) {
      this.setState({ error: 'Select a SharePoint list before opening Grid Designer.' });
      return;
    }
    this.setState({ loading: true, error: '' });
    try {
      var url = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName)
        + "')/fields?$select=InternalName,Title,Description,TypeAsString,Required,ReadOnlyField,Hidden,FromBaseType,Choices,MaxLength,DisplayFormat";
      var response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      });
      if (!response.ok) {
        throw new Error('SharePoint fields could not be loaded.');
      }
      var data = await response.json();
      var sourceFields = toArray(data && data.value).length > 0 ? toArray(data.value) : toArray(data && data.d && data.d.results);
      var fields = sourceFields.filter(function(field: any) {
        var isAttachment = field.InternalName === 'Attachments';
        return !field.Hidden && !SYSTEM_FIELDS[field.InternalName]
          && (!field.FromBaseType || field.InternalName === 'Title' || isAttachment);
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

  private renderCanvas(): JSX.Element {
    return (
      <main className="gd-canvas">
        <div className="gd-canvas-header">
          <div><h3>Grid columns</h3><p>Column order follows this list from top to bottom.</p></div>
          <span>{this.state.fields.length} columns</span>
        </div>
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