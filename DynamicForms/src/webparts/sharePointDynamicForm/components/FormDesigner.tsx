import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './SharePointDynamicForm.module.scss';
import { FormField, FormSchema, FormStep, SPFieldInfo, SPFieldType, FieldConfig, FieldType } from '../../../formEngine/core/types';
import * as strings from 'SharePointDynamicFormWebPartStrings';
import { RichTextEditor } from './RichTextEditor';

type SPFxContext = any;

export interface FormDesignerProps {
  schema: FormSchema;
  context: SPFxContext;
  listName: string;
  onChange: (schema: FormSchema) => void;
}

interface FormDesignerState {
  spFields: SPFieldInfo[];
  loadingFields: boolean;
  fieldsError: string | null;
  selectedStepIndex: number;
  selectedFieldId: string | null;
  activeDesignerTab: 'form' | 'container';
}

interface ISPFieldResponse {
  Id: string;
  InternalName: string;
  Title: string;
  Description?: string;
  TypeAsString: string;
  RichText?: boolean;
  Required: boolean;
  ReadOnlyField: boolean;
  Hidden: boolean;
  FromBaseType?: boolean;
  Choices?: string[] | { results: string[] } | string;
  LookupList?: string;
  LookupField?: string;
  AllowMultipleValues?: boolean;
  MaxLength?: number;
  TextField?: string;
  TermSetId?: string;
  DisplayFormat?: number | string;
}

var SYSTEM_FIELDS: { [key: string]: boolean } = {
  ID: true,
  Created: true,
  Modified: true,
  Author: true,
  Editor: true,
  OData__UIVersionString: true,
  GUID: true,
  ContentType: true,
  AppAuthor: true,
  AppEditor: true,
  Edit: true,
  ItemChildCount: true,
  FolderChildCount: true,
  ComplianceAssetId: true,
};

function createId(prefix: string): string {
  return prefix + '_' + new Date().getTime().toString(36) + '_' + Math.floor(Math.random() * 100000).toString(36);
}

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

function copySchema(schema: FormSchema): FormSchema {
  return JSON.parse(JSON.stringify(schema));
}

function getStepTheme(schema: FormSchema, step: FormStep): { layout: 'stack' | 'grid'; columns: number } {
  var theme = step.theme || schema.theme || {};
  return {
    layout: theme.layout === 'grid' ? 'grid' : 'stack',
    columns: theme.columns && theme.columns > 1 ? theme.columns : 1,
  };
}

function parseChoices(choices: string[] | { results: string[] } | string | undefined): string[] | undefined {
  if (!choices) {
    return undefined;
  }
  if (Array.isArray(choices)) {
    return choices.slice(0);
  }
  if (typeof choices === 'object' && Array.isArray(choices.results)) {
    return choices.results.slice(0);
  }
  if (typeof choices === 'string') {
    return choices.split(';').filter(function(choice) { return !!choice; });
  }
  return undefined;
}
function decodeHtmlEntities(value: string): string {
  if (!value || value.indexOf('&') < 0 || typeof document === 'undefined') {
    return value;
  }
  var decoder = document.createElement('textarea');
  decoder.innerHTML = value;
  return decoder.value;
}

function mapSPFieldType(type: string): SPFieldType {
  switch (type) {
    case 'Text':
      return SPFieldType.Text;
    case 'Note':
      return SPFieldType.Note;
    case 'Number':
    case 'Currency':
      return SPFieldType.Number;
    case 'Integer':
      return SPFieldType.Integer;
    case 'DateTime':
      return SPFieldType.DateTime;
    case 'Choice':
      return SPFieldType.Choice;
    case 'MultiChoice':
      return SPFieldType.MultiChoice;
    case 'Lookup':
      return SPFieldType.Lookup;
    case 'User':
      return SPFieldType.User;
    case 'UserMulti':
      return SPFieldType.UserMulti;
    case 'Boolean':
      return SPFieldType.Boolean;
    case 'URL':
    case 'Hyperlink':
      return SPFieldType.URL;
    case 'Image':
    case 'Thumbnail':
      return SPFieldType.Image;
    case 'TaxonomyFieldType':
      return SPFieldType.Taxonomy;
    case 'TaxonomyFieldTypeMulti':
      return SPFieldType.TaxonomyMulti;
    case 'Attachments':
      return SPFieldType.Attachments;
    default:
      return SPFieldType.Text;
  }
}

function mapFormFieldType(type: SPFieldType): FieldType {
  switch (type) {
    case SPFieldType.Note:
      return 'multiline';
    case SPFieldType.Number:
    case SPFieldType.Integer:
      return 'number';
    case SPFieldType.DateTime:
      return 'datetime';
    case SPFieldType.Choice:
      return 'dropdown';
    case SPFieldType.MultiChoice:
      return 'multiselect';
    case SPFieldType.Lookup:
      return 'lookup';
    case SPFieldType.User:
    case SPFieldType.UserMulti:
      return 'person';
    case SPFieldType.Boolean:
      return 'boolean';
    case SPFieldType.URL:
    case SPFieldType.Hyperlink:
      return 'url';
    case SPFieldType.Image:
      return 'image';
    case SPFieldType.Taxonomy:
    case SPFieldType.TaxonomyMulti:
      return 'taxonomy';
    case SPFieldType.Attachments:
      return 'attachment';
    default:
      return 'text';
  }
}

function getFieldTypeLabel(type: FieldType): string {
  switch (type) {
    case 'text':
      return strings.PropertyFieldTypeText;
    case 'multiline':
      return strings.PropertyFieldTypeMultiline;
    case 'number':
      return strings.PropertyFieldTypeNumber;
    case 'datetime':
      return strings.PropertyFieldTypeDatetime;
    case 'dropdown':
      return strings.PropertyFieldTypeDropdown;
    case 'multiselect':
      return strings.PropertyFieldTypeMultiselect;
    case 'lookup':
      return strings.PropertyFieldTypeLookup;
    case 'person':
      return strings.PropertyFieldTypePerson;
    case 'boolean':
      return strings.PropertyFieldTypeBoolean;
    case 'image':
      return strings.PropertyFieldTypeImage;
    case 'url':
      return strings.PropertyFieldTypeUrl;
    case 'taxonomy':
      return strings.PropertyFieldTypeTaxonomy;
    case 'attachment':
      return strings.PropertyFieldTypeAttachment;
    case 'richtext':
      return strings.PropertyFieldTypeRichtext;
    default:
      return type;
  }
}

export class FormDesigner extends React.Component<FormDesignerProps, FormDesignerState> {
  public constructor(props: FormDesignerProps) {
    super(props);
    this.state = {
      spFields: [],
      loadingFields: false,
      fieldsError: null,
      selectedStepIndex: 0,
      selectedFieldId: null,
      activeDesignerTab: 'form',
    };

    this.handleAddRichText = this.handleAddRichText.bind(this);
  }

  public componentDidMount(): void {
    this.loadFields();
  }

  public componentDidUpdate(prevProps: FormDesignerProps): void {
    if (prevProps.listName !== this.props.listName) {
      this.loadFields();
    }

    if (this.state.selectedStepIndex >= this.props.schema.steps.length) {
      this.setState({ selectedStepIndex: Math.max(0, this.props.schema.steps.length - 1), selectedFieldId: null });
    }
  }

  private getWebUrl(): string {
    return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
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

  private async loadFields(): Promise<void> {
    if (!this.props.listName) {
      this.setState({ spFields: [], loadingFields: false, fieldsError: strings.DesignerNoListSelected });
      return;
    }

    this.setState({ loadingFields: true, fieldsError: null });
    try {
      var response = await this.getWithAcceptFallback(
        this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=Id,InternalName,Title,Description,TypeAsString,RichText,Required,ReadOnlyField,Hidden,FromBaseType,Choices,LookupList,LookupField,AllowMultipleValues,MaxLength,TextField,TermSetId,DisplayFormat"
      );
      if (!response.ok) {
        throw new Error(strings.DesignerLoadFieldsFailed);
      }

      var data = await response.json();
      var fields = data && data.value ? data.value as ISPFieldResponse[] : [];
      if (!fields || fields.length === 0) {
        fields = data && data.d && data.d.results ? data.d.results as ISPFieldResponse[] : [];
      }
      var mapped = fields
        .filter(function(field) {
          var isAttachmentField = field.InternalName === 'Attachments';
          if (field.Hidden) {
            return false;
          }
          if (field.FromBaseType && field.InternalName !== 'Title' && !isAttachmentField) {
            return false;
          }
          return !SYSTEM_FIELDS[field.InternalName];
        })
        .map(function(field) {
          return {
            id: field.Id,
            internalName: field.InternalName,
            title: field.Title || field.InternalName,
            description: field.Description || undefined,
            type: mapSPFieldType(field.TypeAsString),
            richText: field.RichText === true,
            required: field.Required || false,
            readOnly: field.ReadOnlyField || false,
            choices: parseChoices(field.Choices),
            lookupList: field.LookupList || undefined,
            lookupField: field.LookupField || undefined,
            allowMultipleValues: field.AllowMultipleValues || false,
            maxLength: field.MaxLength || undefined,
            textField: field.TextField || undefined,
            termSetId: field.TermSetId || undefined,
            displayFormat: String(field.DisplayFormat) === '0' ? 'dateOnly' : 'dateTime',
          } as SPFieldInfo;
        });

      this.setState({ spFields: mapped, loadingFields: false, fieldsError: null });
    } catch (error) {
      this.setState({
        spFields: [],
        loadingFields: false,
        fieldsError: error && error.message ? error.message : strings.DesignerLoadFieldsFailed,
      });
    }
  }

  private getCurrentStep(): FormStep | null {
    if (!this.props.schema.steps || this.props.schema.steps.length === 0) {
      return null;
    }
    return this.props.schema.steps[this.state.selectedStepIndex] || null;
  }

  private getSelectedField(): FormField | null {
    var currentStep = this.getCurrentStep();
    var selectedFieldId = this.state.selectedFieldId;
    if (!currentStep || !selectedFieldId) {
      return null;
    }

    for (var i = 0; i < currentStep.fields.length; i += 1) {
      var field = currentStep.fields[i];
      if (field && field.id === selectedFieldId) {
        return field;
      }
    }

    return null;
  }

  private getAvailableFields(): SPFieldInfo[] {
    var used: { [key: string]: boolean } = {};
    for (var i = 0; i < this.props.schema.steps.length; i += 1) {
      var stepFields = this.props.schema.steps[i].fields;
      for (var j = 0; j < stepFields.length; j += 1) {
        var field = stepFields[j];
        if (field && field.type !== 'richtext') {
          used[field.fieldName] = true;
        }
      }
    }

    return this.state.spFields.filter(function(field) {
      return !used[field.internalName];
    });
  }

  private updateSchema(nextSchema: FormSchema, selectedFieldId?: string | null): void {
    this.props.onChange(nextSchema);
    if (selectedFieldId !== undefined) {
      this.setState({ selectedFieldId: selectedFieldId });
    }
  }

  private updateCurrentStep(mutator: (step: FormStep) => FormStep): void {
    var nextSchema = copySchema(this.props.schema);
    nextSchema.steps[this.state.selectedStepIndex] = mutator(nextSchema.steps[this.state.selectedStepIndex]);
    this.updateSchema(nextSchema);
  }

  private updateForm(mutator: (schema: FormSchema) => FormSchema): void {
    var nextSchema = copySchema(this.props.schema);
    this.updateSchema(mutator(nextSchema), this.state.selectedFieldId);
  }

  private updateSelectedField(mutator: (field: FormField) => FormField): void {
    var selectedField = this.getSelectedField();
    if (!selectedField) {
      return;
    }

    var nextSchema = copySchema(this.props.schema);
    var step = nextSchema.steps[this.state.selectedStepIndex];
    for (var i = 0; i < step.fields.length; i += 1) {
      var field = step.fields[i];
      if (field && field.id === selectedField.id) {
        step.fields[i] = mutator(field);
        break;
      }
    }
    this.updateSchema(nextSchema, selectedField.id);
  }

  private selectStep(index: number): void {
    this.setState({ selectedStepIndex: index, selectedFieldId: null });
  }

  private addStep(): void {
    var nextSchema = copySchema(this.props.schema);
    nextSchema.steps.push({
      id: createId('step'),
      title: strings.DesignerStepDefaultTitle.replace('{0}', String(nextSchema.steps.length + 1)),
      description: '',
      fields: [],
      visible: true,
    });
    this.props.onChange(nextSchema);
    this.setState({ selectedStepIndex: nextSchema.steps.length - 1, selectedFieldId: null });
  }

  private deleteCurrentStep(): void {
    if (this.props.schema.steps.length <= 1) {
      return;
    }

    var nextSchema = copySchema(this.props.schema);
    nextSchema.steps.splice(this.state.selectedStepIndex, 1);
    this.props.onChange(nextSchema);
    this.setState({
      selectedStepIndex: Math.max(0, this.state.selectedStepIndex - 1),
      selectedFieldId: null,
    });
  }

  private addFieldFromSPField(spField: SPFieldInfo): void {
    var currentStep = this.getCurrentStep();
    if (!currentStep) {
      return;
    }

    var config: FieldConfig = {};
    if (spField.choices && spField.choices.length) {
      config.choices = spField.choices.slice(0);
    }
    if (spField.type === SPFieldType.Choice) {
      config.choiceDisplay = 'dropdown';
    }
    if (spField.type === SPFieldType.MultiChoice) {
      config.choiceDisplay = 'checkboxes';
    }
    if (spField.lookupList) {
      config.lookupList = spField.lookupList;
    }
    if (spField.lookupField) {
      config.lookupField = spField.lookupField;
    }
    if (spField.allowMultipleValues) {
      config.allowMultiple = spField.allowMultipleValues;
    }
    if (spField.maxLength) {
      config.maxLength = spField.maxLength;
    }
    if (spField.type === SPFieldType.Note && spField.richText === true) {
      config.isRichText = true;
    }
    if (spField.type === SPFieldType.DateTime) {
      config.displayFormat = spField.displayFormat || 'dateTime';
    }
    if (spField.termSetId) {
      config.termSetId = spField.termSetId;
    }
    if (spField.internalName === 'Attachments' && this.props.listName) {
      config.listName = this.props.listName;
    }

    var nextField: FormField = {
      id: createId('field'),
      type: mapFormFieldType(spField.type),
      label: spField.title,
      fieldName: spField.internalName,
      description: spField.description,
      required: spField.required,
      readOnly: spField.readOnly,
      disabled: false,
      config: Object.keys(config).length > 0 ? config : undefined,
    };

    var nextSchema = copySchema(this.props.schema);
    nextSchema.steps[this.state.selectedStepIndex].fields.push(nextField);
    this.updateSchema(nextSchema, nextField.id);
  }

  private handleAddRichText(): void {
    var nextField: FormField = {
      id: createId('field'),
      type: 'richtext',
      label: strings.DesignerCustomRichTextTitle,
      fieldName: createId('richtext'),
      required: false,
      readOnly: false,
      disabled: false,
      defaultValue: '',
      config: {
        helpText: ''
      },
    };
    var nextSchema = copySchema(this.props.schema);
    nextSchema.steps[this.state.selectedStepIndex].fields.push(nextField);
    this.updateSchema(nextSchema, nextField.id);
  }

  private getChoiceDisplay(field: FormField): 'dropdown' | 'radio' | 'checkboxes' {
    var configured = String(field.config && field.config.choiceDisplay || '').trim().toLowerCase();
    if (field.type === 'multiselect') {
      return 'checkboxes';
    }
    if (configured === 'radio' || configured === 'checkboxes') {
      return configured as 'radio' | 'checkboxes';
    }
    return 'dropdown';
  }

  private deleteField(fieldId: string): void {
    var nextSchema = copySchema(this.props.schema);
    var fields = nextSchema.steps[this.state.selectedStepIndex].fields;
    nextSchema.steps[this.state.selectedStepIndex].fields = fields.filter(function(field) {
      return !(field && field.id === fieldId);
    });
    this.updateSchema(nextSchema, this.state.selectedFieldId === fieldId ? null : this.state.selectedFieldId);
  }

  private moveField(fieldId: string, direction: number): void {
    var nextSchema = copySchema(this.props.schema);
    var fields = nextSchema.steps[this.state.selectedStepIndex].fields;
    var index = -1;
    for (var i = 0; i < fields.length; i += 1) {
      if (fields[i] && fields[i]!.id === fieldId) {
        index = i;
        break;
      }
    }

    if (index < 0) {
      return;
    }

    var targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= fields.length) {
      return;
    }

    var temp = fields[index];
    fields[index] = fields[targetIndex];
    fields[targetIndex] = temp;
    this.updateSchema(nextSchema, fieldId);
  }

  private renderPreview(): JSX.Element {
    var schema = this.props.schema;

    // Form-level wrapper style
    var formWrapperStyle: React.CSSProperties = { padding: '12px' };
    if (schema.theme) {
      if (schema.theme.fontSize) { formWrapperStyle.fontSize = schema.theme.fontSize; }
      if (schema.theme.fontFamily) { formWrapperStyle.fontFamily = schema.theme.fontFamily; }
      if (schema.theme.fontWeight) { formWrapperStyle.fontWeight = schema.theme.fontWeight as any; }
      if (schema.theme.color) { formWrapperStyle.color = schema.theme.color; }
      if (schema.theme.backgroundColor) { formWrapperStyle.backgroundColor = schema.theme.backgroundColor; }
      if (schema.theme.borderColor || schema.theme.borderWidth) {
        formWrapperStyle.border = (schema.theme.borderWidth || 1) + 'px solid ' + (schema.theme.borderColor || '#cccccc');
        formWrapperStyle.padding = '16px';
      }
      if (schema.theme.borderStyle === 'rounded') {
        formWrapperStyle.borderRadius = String(schema.theme.borderRadius || 8) + 'px';
      }
    }

    // Form-level layout for containers
    var formGridLayout = !!(schema.theme && schema.theme.layout === 'grid' && schema.theme.columns && schema.theme.columns > 1);
    var formColumns = formGridLayout ? (schema.theme.columns || 2) : 1;
    var formLayoutStyle: React.CSSProperties = formGridLayout
      ? { display: 'grid', gridTemplateColumns: 'repeat(' + String(formColumns) + ', minmax(0, 1fr))', gap: '16px' }
      : {};

    var visibleSteps = schema.steps.filter(function(step) { return step.visible !== false; });

    // Build description style with separate font settings
    var descriptionStyle: React.CSSProperties = { marginBottom: '16px' };
    if (schema.theme) {
      if (schema.theme.descriptionFontSize) { descriptionStyle.fontSize = schema.theme.descriptionFontSize; }
      if (schema.theme.descriptionFontFamily) { descriptionStyle.fontFamily = schema.theme.descriptionFontFamily; }
      if (schema.theme.descriptionFontWeight) { descriptionStyle.fontWeight = schema.theme.descriptionFontWeight as any; }
      if (schema.theme.descriptionColor) { descriptionStyle.color = schema.theme.descriptionColor; }
    }

    return (
      <div className={styles.designerPreviewSection}>
        <div className={styles.designerPreviewTitle}>Preview</div>
        <div style={formWrapperStyle}>
          {schema.showTitle !== false && schema.name && <h2 className={styles.designerPreviewFormName}>{schema.name}</h2>}
          {schema.showTitle !== false && schema.description && <div style={descriptionStyle}>{schema.description}</div>}
          <div style={formLayoutStyle}>
            {visibleSteps.map((step) => {
              // Step container style
              var stepTheme = getStepTheme(schema, step);
              var previewGrid = stepTheme.layout === 'grid' && stepTheme.columns > 1;
              var stepContainerStyle: React.CSSProperties = previewGrid
                ? { display: 'grid', gridTemplateColumns: 'repeat(' + String(stepTheme.columns) + ', minmax(0, 1fr))', gap: '12px' }
                : {};

              if (step.theme) {
                if (step.theme.fontSize) { stepContainerStyle.fontSize = step.theme.fontSize; }
                if (step.theme.fontFamily) { stepContainerStyle.fontFamily = step.theme.fontFamily; }
                if (step.theme.fontWeight) { stepContainerStyle.fontWeight = step.theme.fontWeight as any; }
                if (step.theme.color) { stepContainerStyle.color = step.theme.color; }
                if (step.theme.backgroundColor) { stepContainerStyle.backgroundColor = step.theme.backgroundColor; }
                if (step.theme.borderColor || step.theme.borderWidth) {
                  stepContainerStyle.border = (step.theme.borderWidth || 1) + 'px solid ' + (step.theme.borderColor || '#cccccc');
                }
                if (step.theme.borderStyle === 'rounded') {
                  stepContainerStyle.borderRadius = String(step.theme.borderRadius || 8) + 'px';
                }
                if (step.theme.backgroundColor || step.theme.borderColor || step.theme.borderWidth) {
                  stepContainerStyle.padding = '12px';
                }
              }

              return (
                <div key={step.id} style={{ marginBottom: formGridLayout ? '0' : '20px' }}>
                  {step.showTitle !== false && step.title && <div className={styles.designerPreviewStepTitle}>{step.title}</div>}
                  {step.showTitle !== false && step.description && <div className={styles.designerPreviewStepDescription}>{step.description}</div>}
                  <div style={stepContainerStyle}>
                    {step.fields.map((field) => {
                      if (!field || field.visible === false) {
                        return null;
                      }

                      // Custom richtext fields render as content, not form inputs
                      if (field.type === 'richtext' && field.fieldName.startsWith('richtext')) {
                        var richTextCellStyle: React.CSSProperties = previewGrid
                          ? (field.startNewRow === true
                            ? { gridColumn: '1 / -1' }
                            : { gridColumn: 'span ' + String(Math.max(1, field.columnSpan || 1)) })
                          : { marginBottom: '8px' };
                        var richTextStyle: React.CSSProperties = {
                          border: '1px dashed #c8c6c4',
                          borderRadius: '0px',
                          padding: '12px',
                          background: '#faf9f8',
                          wordBreak: 'break-word',
                          lineHeight: '1.6',
                          fontSize: '12px',
                          color: '#605e5c'
                        };
                        if (field.inputFontSize) { richTextStyle.fontSize = field.inputFontSize; }
                        if (field.inputFontFamily) { richTextStyle.fontFamily = field.inputFontFamily; }
                        if (field.inputFontWeight) { richTextStyle.fontWeight = field.inputFontWeight as any; }
                        if (field.inputColor) { richTextStyle.color = field.inputColor; }
                        if (field.fieldBackgroundColor) { richTextStyle.backgroundColor = field.fieldBackgroundColor; }
                        if (field.fieldBorderColor) { richTextStyle.borderColor = field.fieldBorderColor; }
                        if (field.fieldBorderStyle === 'rounded') {
                          richTextStyle.borderRadius = String(field.fieldBorderRadius || 8) + 'px';
                        }
                        var richTextContentRaw = field.defaultValue || '(empty content)';
                        var richTextContent = decodeHtmlEntities(String(richTextContentRaw));
                        return (
                          <div key={field.id} style={richTextCellStyle}>
                            <div style={richTextStyle} dangerouslySetInnerHTML={{ __html: richTextContent } as any} />
                          </div>
                        );
                      }

                      var fieldCellStyle: React.CSSProperties = previewGrid
                        ? (field.startNewRow === true
                          ? { gridColumn: '1 / -1' }
                          : { gridColumn: 'span ' + String(Math.max(1, field.columnSpan || 1)) })
                        : { marginBottom: '8px' };

                      // Field wrapper style
                      var fieldWrapperStyle: React.CSSProperties = {
                        border: '1px dashed #c8c6c4',
                        borderRadius: '0px',
                        padding: '8px 10px',
                        background: '#faf9f8',
                      };
                      if (field.fieldBackgroundColor) { fieldWrapperStyle.backgroundColor = field.fieldBackgroundColor; }
                      if (field.fieldBorderColor || field.fieldBorderWidth) {
                        fieldWrapperStyle.border = (field.fieldBorderWidth || 1) + 'px solid ' + (field.fieldBorderColor || '#cccccc');
                      }
                      if (field.fieldBorderStyle === 'rounded') {
                        fieldWrapperStyle.borderRadius = String(field.fieldBorderRadius || 8) + 'px';
                      }

                      // Label style
                      var labelStyle: React.CSSProperties = { fontWeight: 600 };
                      if (field.labelFontSize) { labelStyle.fontSize = field.labelFontSize; }
                      if (field.labelFontFamily) { labelStyle.fontFamily = field.labelFontFamily; }
                      if (field.labelFontWeight) { labelStyle.fontWeight = field.labelFontWeight as any; }
                      if (field.labelColor) { labelStyle.color = field.labelColor; }

                      // Input/value style
                      var inputStyle: React.CSSProperties = { fontSize: '12px', color: '#605e5c', marginTop: '4px' };
                      if (field.inputFontSize) { inputStyle.fontSize = field.inputFontSize; }
                      if (field.inputFontFamily) { inputStyle.fontFamily = field.inputFontFamily; }
                      if (field.inputFontWeight) { inputStyle.fontWeight = field.inputFontWeight as any; }
                      if (field.inputColor) { inputStyle.color = field.inputColor; }

                      // Determine label position
                      var previewLabelPosition = field.labelPosition || (schema.theme && schema.theme.labelPosition) || 'top';
                      var fieldContentStyle: React.CSSProperties = { display: 'flex', gap: '10px' };
                      if (field.type === 'boolean' && previewLabelPosition === 'bottom') {
                        fieldContentStyle.flexDirection = 'column-reverse';
                      } else if (previewLabelPosition === 'left') {
                        fieldContentStyle.flexDirection = 'row';
                        fieldContentStyle.alignItems = 'center';
                      } else if (field.type === 'boolean' && previewLabelPosition === 'right') {
                        fieldContentStyle.flexDirection = 'row-reverse';
                        fieldContentStyle.justifyContent = 'flex-end';
                        fieldContentStyle.alignItems = 'center';
                      } else {
                        fieldContentStyle.flexDirection = 'column';
                      }
                      var previewLabelStyle: React.CSSProperties = Object.assign(
                        {},
                        labelStyle,
                        previewLabelPosition === 'left' || previewLabelPosition === 'right'
                          ? { flexShrink: 0, minWidth: '100px' }
                          : {}
                      );
                      var previewBooleanText = field.config && field.config.booleanText ? field.config.booleanText : '';

                      return (
                        <div key={field.id} style={fieldCellStyle}>
                          <div style={fieldWrapperStyle}>
                            <div style={fieldContentStyle}>
                              <div style={previewLabelStyle}>{field.label}</div>
                              {field.type === 'boolean' ? (
                                <label style={Object.assign({}, inputStyle, { display: 'inline-flex', alignItems: 'center', gap: '6px', width: 'fit-content', marginTop: 0 })}>
                                  <input type="checkbox" checked={field.defaultValue === true} readOnly={true} />
                                  {previewBooleanText && <span>{previewBooleanText}</span>}
                                </label>
                              ) : (
                                <div style={inputStyle}>{getFieldTypeLabel(field.type)}</div>
                              )}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    );
  }

  private renderPalette(): JSX.Element {
    var availableFields = this.getAvailableFields();

    return (
      <div className={styles.designerPanel}>
        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>{strings.DesignerCustomFields}</div>
          <div className={styles.designerPanelHint}>{strings.DesignerCustomFieldsDesc}</div>
          <button type="button" className={styles.designerPaletteButton} onClick={this.handleAddRichText}>
            <span>{strings.DesignerCustomRichTextTitle}</span>
            <span className={styles.designerPaletteMeta}>{strings.DesignerRichTextChip}</span>
          </button>
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>{strings.DesignerSPFields}</div>
          <div className={styles.designerPanelHint}>{strings.DesignerAddableFieldsCount.replace('{0}', String(availableFields.length))}</div>
          {this.state.loadingFields && <div className={styles.designerEmptyState}>{strings.CommonLoading}</div>}
          {!this.state.loadingFields && this.state.fieldsError && (
            <MessageBar messageBarType={MessageBarType.warning}>{this.state.fieldsError}</MessageBar>
          )}
          {!this.state.loadingFields && !this.state.fieldsError && availableFields.length === 0 && (
            <div className={styles.designerEmptyState}>{strings.DesignerNoAddableFields}</div>
          )}
          {!this.state.loadingFields && !this.state.fieldsError && availableFields.map((field) => (
            <button
              key={field.internalName}
              type="button"
              className={styles.designerPaletteButton}
              onClick={() => this.addFieldFromSPField(field)}
            >
              <span>{field.title}</span>
              <span className={styles.designerPaletteMeta}>{getFieldTypeLabel(mapFormFieldType(field.type))}</span>
            </button>
          ))}
        </div>
      </div>
    );
  }

  private renderCanvas(): JSX.Element {
    var currentStep = this.getCurrentStep();
    if (!currentStep) {
      return <div className={styles.designerEmptyState}>{strings.DesignerStepNotFound}</div>;
    }

    return (
      <div className={styles.designerCanvas}>
        <div className={styles.designerTabs}>
          {this.props.schema.steps.map((step, index) => (
            <button
              key={step.id}
              type="button"
              className={index === this.state.selectedStepIndex ? styles.designerTabActive : styles.designerTab}
              onClick={() => this.selectStep(index)}
            >
              {step.title}
            </button>
          ))}
          <button type="button" className={styles.designerActionButton} onClick={() => this.addStep()}>{strings.DesignerAddStep}</button>
          {this.props.schema.steps.length > 1 && (
            <button type="button" className={styles.designerDeleteButton} onClick={() => this.deleteCurrentStep()}>{strings.DesignerDeleteStep}</button>
          )}
        </div>

        <div className={styles.designerCanvasBody}>
          <div className={styles.designerStepHeader}>
            <div className={styles.designerStepTitle}>{currentStep.title || strings.DesignerStepNotFound}</div>
            {currentStep.description && <div className={styles.designerStepDescription}>{currentStep.description}</div>}
          </div>

          {currentStep.fields.length === 0 && (
            <div className={styles.designerEmptyState}>{strings.DesignerClickAddField}</div>
          )}

          {currentStep.fields.map((field) => {
            if (!field) {
              return null;
            }
            return (
              <div key={field.id} className={field.id === this.state.selectedFieldId ? styles.designerFieldCardActive : styles.designerFieldCard}>
                <div className={styles.designerFieldInfo}>
                  <div className={styles.designerFieldLabel}>{field.label}</div>
                  <div className={styles.designerFieldMeta}>{getFieldTypeLabel(field.type)}</div>
                </div>
                <div className={styles.designerFieldActions}>
                  <button type="button" className={styles.designerInlineButton} onClick={() => this.moveField(field.id, -1)} disabled={currentStep.fields[0] === field}>{strings.DesignerMoveUp}</button>
                  <button type="button" className={styles.designerInlineButton} onClick={() => this.moveField(field.id, 1)} disabled={currentStep.fields[currentStep.fields.length - 1] === field}>{strings.DesignerMoveDown}</button>
                  <button type="button" className={styles.designerInlineButton} onClick={() => this.setState({ selectedFieldId: field.id })}>{strings.DesignerEdit}</button>
                  <button type="button" className={styles.designerInlineButtonDanger} onClick={() => this.deleteField(field.id)}>{strings.DesignerDelete}</button>
                </div>
              </div>
            );
          })}

          {this.renderPreview()}
        </div>
      </div>
    );
  }

  private renderFieldEditor(field: FormField): JSX.Element {
    // Custom richtext fields have a different editor UI - no field properties
    if (field.type === 'richtext' && field.fieldName.startsWith('richtext')) {
      var contentValue = field.defaultValue || '';
      return (
        <div className={styles.designerPanel}>
          <div className={styles.designerPanelSection}>
            <div className={styles.designerPanelTitle}>Rich Text Content</div>
            <div className={styles.designerPanelHint}>Add instructional or informational text for the form. This content is not saved to SharePoint.</div>

            <label className={styles.designerFormLabel}>Content</label>
            <div dir="ltr" style={{ direction: 'ltr', textAlign: 'left', display: 'block' }}>
              <RichTextEditor
                value={String(contentValue)}
                onChange={(html) => this.updateSelectedField(function(nextField) {
                  nextField.defaultValue = html;
                  return nextField;
                })}
                placeholder="Enter your rich text content here..."
              />
            </div>
          </div>

          <div className={styles.designerPanelSection}>
            <div className={styles.designerPanelTitle}>Content Styling</div>

            <label className={styles.designerFormLabel}>Font Size (px)</label>
            <input
              className={styles.designerInput}
              type="number"
              min={8}
              max={72}
              value={String(field.inputFontSize || 14)}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                var parsed = parseInt(ev.currentTarget.value, 10);
                nextField.inputFontSize = isNaN(parsed) ? undefined : parsed;
                return nextField;
              })}
            />

            <label className={styles.designerFormLabel}>Font Family</label>
            <select
              className={styles.designerInput}
              value={field.inputFontFamily || 'inherit'}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.inputFontFamily = ev.currentTarget.value === 'inherit' ? undefined : ev.currentTarget.value;
                return nextField;
              })}
            >
              <option value="inherit">Inherit</option>
              <option value="Arial">Arial</option>
              <option value="'Segoe UI'">Segoe UI</option>
              <option value="Verdana">Verdana</option>
              <option value="Georgia">Georgia</option>
              <option value="Courier New">Courier New</option>
              <option value="Times New Roman">Times New Roman</option>
            </select>

            <label className={styles.designerFormLabel}>Font Weight</label>
            <select
              className={styles.designerInput}
              value={field.inputFontWeight || 'normal'}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.inputFontWeight = ev.currentTarget.value as 'normal' | 'bold' | '500' | '600' | '700' || undefined;
                return nextField;
              })}
            >
              <option value="normal">Normal</option>
              <option value="500">Medium (500)</option>
              <option value="600">Semi-Bold (600)</option>
              <option value="bold">Bold</option>
              <option value="700">Bold (700)</option>
            </select>

            <label className={styles.designerFormLabel}>Font Color</label>
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <input
                type="color"
                value={field.inputColor || '#000000'}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.inputColor = ev.currentTarget.value || undefined;
                  return nextField;
                })}
                style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
              />
              <input
                className={styles.designerInput}
                type="text"
                value={field.inputColor || ''}
                placeholder="#000000"
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.inputColor = ev.currentTarget.value || undefined;
                  return nextField;
                })}
                style={{ flex: 1 }}
              />
            </div>

            <label className={styles.designerFormLabel}>Background Color</label>
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <input
                type="color"
                value={field.fieldBackgroundColor || '#f5f5f5'}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.fieldBackgroundColor = ev.currentTarget.value || undefined;
                  return nextField;
                })}
                style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
              />
              <input
                className={styles.designerInput}
                type="text"
                value={field.fieldBackgroundColor || ''}
                placeholder="#f5f5f5"
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.fieldBackgroundColor = ev.currentTarget.value || undefined;
                  return nextField;
                })}
                style={{ flex: 1 }}
              />
            </div>
          </div>
        </div>
      );
    }

    var helpText = field.config && field.config.helpText ? field.config.helpText : '';
    var fieldChoices = field.config && field.config.choices ? field.config.choices : [];
    var choiceDisplay = this.getChoiceDisplay(field);
    var lookupDefault = field.defaultValue === undefined || field.defaultValue === null ? '' : String(field.defaultValue);
    var scalarDefault = field.defaultValue === undefined || field.defaultValue === null ? '' : String(field.defaultValue);
    var multiSelectDefault = '';
    if (Array.isArray(field.defaultValue)) {
      multiSelectDefault = (field.defaultValue as Array<string>).join(';');
    } else if (field.defaultValue !== undefined && field.defaultValue !== null) {
      multiSelectDefault = String(field.defaultValue);
    }
    var spField: SPFieldInfo | null = null;
    for (var i = 0; i < this.state.spFields.length; i += 1) {
      if (this.state.spFields[i].internalName === field.fieldName) {
        spField = this.state.spFields[i];
        break;
      }
    }

    return (
      <div className={styles.designerPanel}>
        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>{strings.PropertyPanelTitle}</div>

          <label className={styles.designerFormLabel}>{strings.PropertyPanelFieldName}</label>
          <input
            className={styles.designerInput}
            type="text"
            value={field.label}
            title={strings.PropertyPanelFieldName}
            placeholder={strings.PropertyPanelFieldName}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.label = ev.currentTarget.value;
              return nextField;
            })}
          />

          <label className={styles.designerFormLabel}>{strings.PropertyPanelFieldType}</label>
          <div className={styles.designerReadOnlyValue}>{getFieldTypeLabel(field.type)}</div>

          <label className={styles.designerFormLabel}>{strings.PropertyPanelFieldInternalName}</label>
          <div className={styles.designerReadOnlyValue}>{field.fieldName}</div>

          <label className={styles.designerFormLabel}>{strings.PropertyPanelColumnSpan}</label>
          <input
            className={styles.designerInput}
            type="number"
            min={1}
            max={12}
            value={String(Math.max(1, field.columnSpan || 1))}
            title={strings.PropertyPanelColumnSpan}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextField.columnSpan = isNaN(parsed) ? 1 : Math.max(1, parsed);
              return nextField;
            })}
          />

          <label className={styles.designerCheckboxRow}>
            <input
              type="checkbox"
              checked={field.startNewRow === true}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.startNewRow = ev.currentTarget.checked;
                return nextField;
              })}
            />
            <span>{strings.DesignerStartNewRow}</span>
          </label>

          <label className={styles.designerFormLabel}>Label Position</label>
          <select
            className={styles.designerInput}
            value={field.labelPosition || 'top'}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              var nextPosition = ev.currentTarget.value;
              nextField.labelPosition = (
                nextPosition === 'bottom' || nextPosition === 'left' || nextPosition === 'right'
                  ? nextPosition
                  : 'top'
              ) as 'top' | 'bottom' | 'left' | 'right';
              return nextField;
            })}
          >
            <option value="top">Above input (top)</option>
            {field.type === 'boolean' && <option value="bottom">Below input (bottom)</option>}
            <option value="left">Next to input (left)</option>
            {field.type === 'boolean' && <option value="right">After input (right)</option>}
          </select>

          {spField && spField.required ? (
            <div className={styles.designerPanelHint} style={{ color: '#a80000', fontWeight: 600, marginBottom: '8px' }}>
              ⚠ {strings.PropertyPanelFieldRequiredLocked}
            </div>
          ) : (
            <label className={styles.designerCheckboxRow}>
              <input
                type="checkbox"
                checked={field.required === true}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.required = ev.currentTarget.checked;
                  return nextField;
                })}
              />
              <span>{strings.PropertyPanelFieldRequired}</span>
            </label>
          )}

          {(field.required === true || (spField && spField.required)) && (
            <div>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelFieldRequiredMessage}</label>
              <input
                className={styles.designerInput}
                type="text"
                value={field.requiredMessage || ''}
                title={strings.PropertyPanelFieldRequiredMessage}
                placeholder={spField && spField.description ? spField.description : strings.PropertyPanelFieldRequiredMessagePlaceholder}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.requiredMessage = ev.currentTarget.value || undefined;
                  return nextField;
                })}
              />
            </div>
          )}

          <label className={styles.designerCheckboxRow}>
            <input
              type="checkbox"
              checked={field.visible !== false}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.visible = ev.currentTarget.checked;
                return nextField;
              })}
            />
            <span>{strings.DesignerShowStep}</span>
          </label>

          <label className={styles.designerCheckboxRow}>
            <input
              type="checkbox"
              checked={field.disabled !== true}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.disabled = !ev.currentTarget.checked;
                return nextField;
              })}
            />
            <span>{strings.PropertyPanelFieldEnabled}</span>
          </label>
          <div className={styles.designerPanelHint}>{strings.PropertyPanelFieldEnabledHint}</div>

          <label className={styles.designerFormLabel}>{strings.PropertyPanelHelpText}</label>
          <textarea
            className={styles.designerTextarea}
            value={helpText}
            title={strings.PropertyPanelHelpText}
            placeholder={strings.PropertyPanelHelpTextPlaceholder}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.config = nextField.config || {};
              nextField.config.helpText = ev.currentTarget.value;
              return nextField;
            })}
          />

          {field.type === 'attachment' && (
            <div>
              <label className={styles.designerCheckboxRow}>
                <input
                  type="checkbox"
                  checked={!field.config || field.config.allowAttachmentDelete !== false}
                  onChange={(ev) => this.updateSelectedField(function(nextField) {
                    nextField.config = nextField.config || {};
                    nextField.config.allowAttachmentDelete = ev.currentTarget.checked;
                    return nextField;
                  })}
                />
                <span>{strings.PropertyPanelAttachmentAllowDelete}</span>
              </label>
              <div className={styles.designerPanelHint}>{strings.PropertyPanelAttachmentAllowDeleteHint}</div>
            </div>
          )}

          {!spField && field.type !== 'richtext' && (
            <MessageBar messageBarType={MessageBarType.warning}>{strings.PropertyPanelMissingSPField}</MessageBar>
          )}
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>{strings.PropertyPanelDefaultsGroup}</div>
          <div className={styles.designerPanelHint}>{strings.PropertyPanelDefaultsGroupDesc}</div>

          {(field.type === 'text' || field.type === 'multiline' || field.type === 'url') && (
            <div>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelDefaultText}</label>
              <input
                className={styles.designerInput}
                type="text"
                value={scalarDefault}
                title={strings.PropertyPanelDefaultText}
                placeholder={strings.PropertyPanelDefaultTextPlaceholder}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.defaultValue = ev.currentTarget.value;
                  return nextField;
                })}
              />
            </div>
          )}

          {field.type === 'number' && (
            <div>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelDefaultNumber}</label>
              <input
                className={styles.designerInput}
                type="number"
                value={scalarDefault}
                title={strings.PropertyPanelDefaultNumber}
                placeholder={strings.PropertyPanelDefaultNumberPlaceholder}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.defaultValue = ev.currentTarget.value;
                  return nextField;
                })}
              />
            </div>
          )}

          {field.type === 'datetime' && (
            <div>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelDateFormat}</label>
              <select
                className={styles.designerInput}
                value={field.config && field.config.displayFormat || 'dateTime'}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  var nextFormat = ev.currentTarget.value === 'dateOnly' || ev.currentTarget.value === 'timeOnly' ? ev.currentTarget.value : 'dateTime';
                  nextField.config = nextField.config || {};
                  nextField.config.displayFormat = nextFormat as 'dateOnly' | 'dateTime' | 'timeOnly';
                  if (nextFormat === 'dateOnly' && nextField.defaultValue) {
                    nextField.defaultValue = String(nextField.defaultValue).substring(0, 10);
                  }
                  if (nextFormat === 'timeOnly' && nextField.defaultValue) {
                    var defaultText = String(nextField.defaultValue);
                    var timeIndex = defaultText.indexOf('T');
                    nextField.defaultValue = timeIndex >= 0 ? defaultText.substring(timeIndex + 1, timeIndex + 6) : defaultText;
                  }
                  console.log('[FormDesigner] datetime field "' + nextField.label + '" format changed to: ' + nextFormat);
                  return nextField;
                })}
              >
                <option value="dateOnly">{strings.PropertyPanelDateFormatDateOnly}</option>
                <option value="dateTime">{strings.PropertyPanelDateFormatDateTime}</option>
                <option value="timeOnly">{strings.PropertyPanelDateFormatTimeOnly}</option>
              </select>
              {(!field.config || field.config.displayFormat !== 'dateOnly') && (
                <div>
                  <label className={styles.designerFormLabel}>{strings.PropertyPanelTimeZone}</label>
                  <select
                    className={styles.designerInput}
                    value={(field.config && field.config.timeZone) || 'UTC'}
                    onChange={(ev) => this.updateSelectedField(function(nextField) {
                      var nextZone = (ev.currentTarget.value === 'local' ? 'local' : 'UTC') as 'UTC' | 'local';
                      nextField.config = nextField.config || {};
                      nextField.config.timeZone = nextZone;
                      console.log('[FormDesigner] datetime field "' + nextField.label + '" time zone changed to: ' + nextZone);
                      return nextField;
                    })}
                  >
                    <option value="UTC">{strings.PropertyPanelTimeZoneUtc}</option>
                    <option value="local">{strings.PropertyPanelTimeZoneLocal}</option>
                  </select>
                </div>
              )}
              <label className={styles.designerFormLabel}>{strings.PropertyPanelDefaultDateTime}</label>
              <input
                className={styles.designerInput}
                type={field.config && field.config.displayFormat === 'dateOnly' ? 'date' : field.config && field.config.displayFormat === 'timeOnly' ? 'time' : 'datetime-local'}
                value={scalarDefault}
                title={strings.PropertyPanelDefaultDateTime}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.defaultValue = ev.currentTarget.value;
                  return nextField;
                })}
              />
            </div>
          )}

          {field.type === 'boolean' && (
            <div>
              <label className={styles.designerCheckboxRow}>
                <input
                  type="checkbox"
                  checked={field.defaultValue === true}
                  onChange={(ev) => this.updateSelectedField(function(nextField) {
                    nextField.defaultValue = ev.currentTarget.checked;
                    return nextField;
                  })}
                />
                <span>{strings.PropertyPanelDefaultBoolean}</span>
              </label>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelBooleanText}</label>
              <input
                className={styles.designerInput}
                type="text"
                value={field.config && field.config.booleanText ? field.config.booleanText : ''}
                placeholder={strings.PropertyPanelBooleanTextPlaceholder}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.config = nextField.config || {};
                  nextField.config.booleanText = ev.currentTarget.value;
                  return nextField;
                })}
              />
              <div className={styles.designerPanelHint}>{strings.PropertyPanelBooleanTextHint}</div>
            </div>
          )}

          {field.type === 'dropdown' && (
            <div>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelChoiceDisplay}</label>
              <select
                className={styles.designerInput}
                value={choiceDisplay}
                title={strings.PropertyPanelChoiceDisplay}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.config = nextField.config || {};
                  nextField.config.choiceDisplay = ev.currentTarget.value as 'dropdown' | 'radio' | 'checkboxes';
                  return nextField;
                })}
              >
                <option value="dropdown">{strings.PropertyPanelChoiceDisplayDropdown}</option>
                <option value="radio">{strings.PropertyPanelChoiceDisplayRadio}</option>
              </select>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelDefaultChoice}</label>
              <select
                className={styles.designerInput}
                value={scalarDefault}
                title={strings.PropertyPanelDefaultChoice}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.defaultValue = ev.currentTarget.value;
                  return nextField;
                })}
              >
                <option value="">{strings.PropertyPanelDefaultNone}</option>
                {fieldChoices.map(function(choice) {
                  return <option key={choice} value={choice}>{choice}</option>;
                })}
              </select>
              {fieldChoices.length === 0 && <div className={styles.designerPanelHint}>{strings.PropertyPanelDefaultChoiceEmpty}</div>}
            </div>
          )}

          {field.type === 'multiselect' && (
            <div>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelExclusiveChoiceValue}</label>
              <input
                className={styles.designerInput}
                type="text"
                value={field.config && field.config.exclusiveChoiceValue ? field.config.exclusiveChoiceValue : ''}
                title={strings.PropertyPanelExclusiveChoiceValue}
                placeholder={strings.PropertyPanelExclusiveChoiceValuePlaceholder}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.config = nextField.config || {};
                  nextField.config.exclusiveChoiceValue = ev.currentTarget.value.trim();
                  return nextField;
                })}
              />
              <div className={styles.designerPanelHint}>{strings.PropertyPanelExclusiveChoiceValueHint}</div>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelDefaultMultiSelect}</label>
              <input
                className={styles.designerInput}
                type="text"
                value={multiSelectDefault}
                title={strings.PropertyPanelDefaultMultiSelect}
                placeholder={strings.PropertyPanelDefaultMultiSelectPlaceholder}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.defaultValue = ev.currentTarget.value
                    .split(';')
                    .map(function(choice) { return choice.trim(); })
                    .filter(function(choice) { return !!choice; });
                  return nextField;
                })}
              />
              <div className={styles.designerPanelHint}>{strings.PropertyPanelDefaultMultiSelectHint}</div>
            </div>
          )}

          {field.type === 'lookup' && (
            <div>
              {field.config && field.config.allowMultiple ? (
                <div className={styles.designerPanelHint}>{strings.PropertyPanelChoiceDisplayMultiHint}</div>
              ) : (
                <div>
                  <label className={styles.designerFormLabel}>{strings.PropertyPanelChoiceDisplay}</label>
                  <select
                    className={styles.designerInput}
                    value={choiceDisplay}
                    title={strings.PropertyPanelChoiceDisplay}
                    onChange={(ev) => this.updateSelectedField(function(nextField) {
                      nextField.config = nextField.config || {};
                      nextField.config.choiceDisplay = ev.currentTarget.value as 'dropdown' | 'radio' | 'checkboxes';
                      return nextField;
                    })}
                  >
                    <option value="dropdown">{strings.PropertyPanelChoiceDisplayDropdown}</option>
                    <option value="radio">{strings.PropertyPanelChoiceDisplayRadio}</option>
                  </select>
                </div>
              )}
              {field.config && field.config.allowMultiple && (
                <div>
                  <label className={styles.designerFormLabel}>{strings.PropertyPanelExclusiveChoiceValue}</label>
                  <input
                    className={styles.designerInput}
                    type="text"
                    value={field.config && field.config.exclusiveChoiceValue ? field.config.exclusiveChoiceValue : ''}
                    title={strings.PropertyPanelExclusiveChoiceValue}
                    placeholder={strings.PropertyPanelExclusiveChoiceValuePlaceholder}
                    onChange={(ev) => this.updateSelectedField(function(nextField) {
                      nextField.config = nextField.config || {};
                      nextField.config.exclusiveChoiceValue = ev.currentTarget.value.trim();
                      return nextField;
                    })}
                  />
                  <div className={styles.designerPanelHint}>{strings.PropertyPanelExclusiveChoiceValueHint}</div>
                </div>
              )}
              <label className={styles.designerFormLabel}>{strings.PropertyPanelDefaultLookupId}</label>
              <input
                className={styles.designerInput}
                type="number"
                min={1}
                step={1}
                value={lookupDefault}
                title={strings.PropertyPanelDefaultLookupId}
                placeholder={strings.PropertyPanelDefaultLookupIdPlaceholder}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.defaultValue = ev.currentTarget.value.trim();
                  return nextField;
                })}
              />
              <div className={styles.designerPanelHint}>{strings.PropertyPanelDefaultLookupIdHint}</div>
            </div>
          )}

          {field.type === 'richtext' && (
            <div>
              <label className={styles.designerFormLabel}>{strings.PropertyPanelRichTextContent}</label>
              <textarea
                className={styles.designerTextareaLarge}
                value={field.defaultValue ? String(field.defaultValue) : ''}
                title={strings.PropertyPanelRichTextContent}
                placeholder={strings.DesignerRichTextPlaceholder}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  nextField.defaultValue = ev.currentTarget.value;
                  return nextField;
                })}
              />
            </div>
          )}

          {(field.type === 'person' || field.type === 'taxonomy' || field.type === 'attachment' || field.type === 'image') && (
            <div className={styles.designerPanelHint}>{strings.PropertyPanelDefaultUnsupportedHint}</div>
          )}
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Label Font Settings</div>

          <label className={styles.designerFormLabel}>Font Size (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={8}
            max={72}
            value={String(field.labelFontSize || 14)}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextField.labelFontSize = isNaN(parsed) ? undefined : parsed;
              return nextField;
            })}
          />

          <label className={styles.designerFormLabel}>Font Family</label>
          <select
            className={styles.designerInput}
            value={field.labelFontFamily || 'inherit'}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.labelFontFamily = ev.currentTarget.value === 'inherit' ? undefined : ev.currentTarget.value;
              return nextField;
            })}
          >
            <option value="inherit">Inherit</option>
            <option value="Arial">Arial</option>
            <option value="'Segoe UI'">Segoe UI</option>
            <option value="Verdana">Verdana</option>
            <option value="Georgia">Georgia</option>
            <option value="Courier New">Courier New</option>
            <option value="Times New Roman">Times New Roman</option>
          </select>

          <label className={styles.designerFormLabel}>Font Weight</label>
          <select
            className={styles.designerInput}
            value={field.labelFontWeight || 'normal'}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.labelFontWeight = ev.currentTarget.value as 'normal' | 'bold' | '500' | '600' | '700' || undefined;
              return nextField;
            })}
          >
            <option value="normal">Normal</option>
            <option value="500">Medium (500)</option>
            <option value="600">Semi-Bold (600)</option>
            <option value="bold">Bold</option>
            <option value="700">Bold (700)</option>
          </select>

          <label className={styles.designerFormLabel}>Label Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={field.labelColor || '#000000'}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.labelColor = ev.currentTarget.value || undefined;
                return nextField;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={field.labelColor || ''}
              placeholder="#000000"
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.labelColor = ev.currentTarget.value || undefined;
                return nextField;
              })}
              style={{ flex: 1 }}
            />
          </div>
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Input Font Settings</div>

          <label className={styles.designerFormLabel}>Font Size (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={8}
            max={72}
            value={String(field.inputFontSize || 14)}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextField.inputFontSize = isNaN(parsed) ? undefined : parsed;
              return nextField;
            })}
          />

          <label className={styles.designerFormLabel}>Font Family</label>
          <select
            className={styles.designerInput}
            value={field.inputFontFamily || 'inherit'}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.inputFontFamily = ev.currentTarget.value === 'inherit' ? undefined : ev.currentTarget.value;
              return nextField;
            })}
          >
            <option value="inherit">Inherit</option>
            <option value="Arial">Arial</option>
            <option value="'Segoe UI'">Segoe UI</option>
            <option value="Verdana">Verdana</option>
            <option value="Georgia">Georgia</option>
            <option value="Courier New">Courier New</option>
            <option value="Times New Roman">Times New Roman</option>
          </select>

          <label className={styles.designerFormLabel}>Font Weight</label>
          <select
            className={styles.designerInput}
            value={field.inputFontWeight || 'normal'}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.inputFontWeight = ev.currentTarget.value as 'normal' | 'bold' | '500' | '600' | '700' || undefined;
              return nextField;
            })}
          >
            <option value="normal">Normal</option>
            <option value="500">Medium (500)</option>
            <option value="600">Semi-Bold (600)</option>
            <option value="bold">Bold</option>
            <option value="700">Bold (700)</option>
          </select>

          <label className={styles.designerFormLabel}>Input Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={field.inputColor || '#000000'}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.inputColor = ev.currentTarget.value || undefined;
                return nextField;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={field.inputColor || ''}
              placeholder="#000000"
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.inputColor = ev.currentTarget.value || undefined;
                return nextField;
              })}
              style={{ flex: 1 }}
            />
          </div>
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Field Background & Border</div>

          <label className={styles.designerFormLabel}>Background Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={field.fieldBackgroundColor || '#ffffff'}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.fieldBackgroundColor = ev.currentTarget.value || undefined;
                return nextField;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={field.fieldBackgroundColor || ''}
              placeholder="#ffffff"
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.fieldBackgroundColor = ev.currentTarget.value || undefined;
                return nextField;
              })}
              style={{ flex: 1 }}
            />
          </div>

          <label className={styles.designerFormLabel}>Border Style</label>
          <select
            className={styles.designerInput}
            value={field.fieldBorderStyle || 'square'}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.fieldBorderStyle = ev.currentTarget.value as 'square' | 'rounded' || undefined;
              return nextField;
            })}
          >
            <option value="square">Square Corners</option>
            <option value="rounded">Rounded Corners</option>
          </select>

          {(field.fieldBorderStyle || 'square') === 'rounded' && (
            <div>
              <label className={styles.designerFormLabel}>Corner Radius (px)</label>
              <input
                className={styles.designerInput}
                type="number"
                min={0}
                max={40}
                value={String(field.fieldBorderRadius || 8)}
                onChange={(ev) => this.updateSelectedField(function(nextField) {
                  var parsedRadius = parseInt(ev.currentTarget.value, 10);
                  nextField.fieldBorderRadius = isNaN(parsedRadius) ? 8 : parsedRadius;
                  return nextField;
                })}
              />
            </div>
          )}

          <label className={styles.designerFormLabel}>Border Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={field.fieldBorderColor || '#cccccc'}
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.fieldBorderColor = ev.currentTarget.value || undefined;
                return nextField;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={field.fieldBorderColor || ''}
              placeholder="#cccccc"
              onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.fieldBorderColor = ev.currentTarget.value || undefined;
                return nextField;
              })}
              style={{ flex: 1 }}
            />
          </div>

          <label className={styles.designerFormLabel}>Border Width (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={0}
            max={10}
            value={String(field.fieldBorderWidth || 1)}
            onChange={(ev) => this.updateSelectedField(function(nextField) {
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextField.fieldBorderWidth = isNaN(parsed) ? 1 : parsed;
              return nextField;
            })}
          />
        </div>
      </div>
    );
  }

  private renderFormSettings(): JSX.Element {
    var theme = (this.props.schema && this.props.schema.theme) || {};
    return (
      <div>
        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>{strings.PropertyPanelFormMessages}</div>
          <label className={styles.designerFormLabel}>{strings.PropertyPanelPermissionDeniedMessage}</label>
          <textarea
            className={styles.designerTextarea}
            value={this.props.schema.permissionDeniedMessage || strings.PermissionDeniedDefault}
            placeholder={strings.PermissionDeniedDefault}
            onChange={(ev) => this.updateForm(function(nextForm) {
              nextForm.permissionDeniedMessage = ev.currentTarget.value || undefined;
              return nextForm;
            })}
          />
          <div className={styles.designerPanelHint}>{strings.PropertyPanelPermissionDeniedMessageHint}</div>
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Form Font Settings</div>

          <label className={styles.designerFormLabel}>Font Size (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={8}
            max={72}
            value={String(theme.fontSize || 14)}
            onChange={(ev) => this.updateForm(function(nextForm) {
              if (!nextForm.theme) nextForm.theme = {};
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextForm.theme.fontSize = isNaN(parsed) ? undefined : parsed;
              return nextForm;
            })}
          />

          <label className={styles.designerFormLabel}>Font Family</label>
          <select
            className={styles.designerInput}
            value={theme.fontFamily || 'inherit'}
            onChange={(ev) => this.updateForm(function(nextForm) {
              if (!nextForm.theme) nextForm.theme = {};
              nextForm.theme.fontFamily = ev.currentTarget.value === 'inherit' ? undefined : ev.currentTarget.value;
              return nextForm;
            })}
          >
            <option value="inherit">Inherit</option>
            <option value="Arial">Arial</option>
            <option value="'Segoe UI'">Segoe UI</option>
            <option value="Verdana">Verdana</option>
            <option value="Georgia">Georgia</option>
            <option value="Courier New">Courier New</option>
            <option value="Times New Roman">Times New Roman</option>
          </select>

          <label className={styles.designerFormLabel}>Font Weight</label>
          <select
            className={styles.designerInput}
            value={theme.fontWeight || 'normal'}
            onChange={(ev) => this.updateForm(function(nextForm) {
              if (!nextForm.theme) nextForm.theme = {};
              nextForm.theme.fontWeight = ev.currentTarget.value as 'normal' | 'bold' | '500' | '600' | '700' || undefined;
              return nextForm;
            })}
          >
            <option value="normal">Normal</option>
            <option value="500">Medium (500)</option>
            <option value="600">Semi-Bold (600)</option>
            <option value="bold">Bold</option>
            <option value="700">Bold (700)</option>
          </select>

          <label className={styles.designerFormLabel}>Font Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={theme.color || '#000000'}
              onChange={(ev) => this.updateForm(function(nextForm) {
                if (!nextForm.theme) nextForm.theme = {};
                nextForm.theme.color = ev.currentTarget.value || undefined;
                return nextForm;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={theme.color || ''}
              placeholder="#000000"
              onChange={(ev) => this.updateForm(function(nextForm) {
                if (!nextForm.theme) nextForm.theme = {};
                nextForm.theme.color = ev.currentTarget.value || undefined;
                return nextForm;
              })}
              style={{ flex: 1 }}
            />
          </div>
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Description Font Settings</div>

          <label className={styles.designerFormLabel}>Font Size (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={8}
            max={72}
            value={String(theme.descriptionFontSize || 14)}
            onChange={(ev) => this.updateForm(function(nextForm) {
              if (!nextForm.theme) nextForm.theme = {};
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextForm.theme.descriptionFontSize = isNaN(parsed) ? undefined : parsed;
              return nextForm;
            })}
          />

          <label className={styles.designerFormLabel}>Font Family</label>
          <select
            className={styles.designerInput}
            value={theme.descriptionFontFamily || 'inherit'}
            onChange={(ev) => this.updateForm(function(nextForm) {
              if (!nextForm.theme) nextForm.theme = {};
              nextForm.theme.descriptionFontFamily = ev.currentTarget.value === 'inherit' ? undefined : ev.currentTarget.value;
              return nextForm;
            })}
          >
            <option value="inherit">Inherit</option>
            <option value="Arial">Arial</option>
            <option value="'Segoe UI'">Segoe UI</option>
            <option value="Verdana">Verdana</option>
            <option value="Georgia">Georgia</option>
            <option value="Courier New">Courier New</option>
            <option value="Times New Roman">Times New Roman</option>
          </select>

          <label className={styles.designerFormLabel}>Font Weight</label>
          <select
            className={styles.designerInput}
            value={theme.descriptionFontWeight || 'normal'}
            onChange={(ev) => this.updateForm(function(nextForm) {
              if (!nextForm.theme) nextForm.theme = {};
              nextForm.theme.descriptionFontWeight = ev.currentTarget.value as 'normal' | 'bold' | '500' | '600' | '700' || undefined;
              return nextForm;
            })}
          >
            <option value="normal">Normal</option>
            <option value="500">Medium (500)</option>
            <option value="600">Semi-Bold (600)</option>
            <option value="bold">Bold</option>
            <option value="700">Bold (700)</option>
          </select>

          <label className={styles.designerFormLabel}>Font Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={theme.descriptionColor || '#666666'}
              onChange={(ev) => this.updateForm(function(nextForm) {
                if (!nextForm.theme) nextForm.theme = {};
                nextForm.theme.descriptionColor = ev.currentTarget.value || undefined;
                return nextForm;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={theme.descriptionColor || ''}
              placeholder="#666666"
              onChange={(ev) => this.updateForm(function(nextForm) {
                if (!nextForm.theme) nextForm.theme = {};
                nextForm.theme.descriptionColor = ev.currentTarget.value || undefined;
                return nextForm;
              })}
              style={{ flex: 1 }}
            />
          </div>
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Form Background & Border</div>

          <label className={styles.designerFormLabel}>Background Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={theme.backgroundColor || '#ffffff'}
              onChange={(ev) => this.updateForm(function(nextForm) {
                if (!nextForm.theme) nextForm.theme = {};
                nextForm.theme.backgroundColor = ev.currentTarget.value || undefined;
                return nextForm;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={theme.backgroundColor || ''}
              placeholder="#ffffff"
              onChange={(ev) => this.updateForm(function(nextForm) {
                if (!nextForm.theme) nextForm.theme = {};
                nextForm.theme.backgroundColor = ev.currentTarget.value || undefined;
                return nextForm;
              })}
              style={{ flex: 1 }}
            />
          </div>

          <label className={styles.designerFormLabel}>Border Style</label>
          <select
            className={styles.designerInput}
            value={theme.borderStyle || 'square'}
            onChange={(ev) => this.updateForm(function(nextForm) {
              if (!nextForm.theme) nextForm.theme = {};
              nextForm.theme.borderStyle = ev.currentTarget.value as 'square' | 'rounded' || undefined;
              return nextForm;
            })}
          >
            <option value="square">Square Corners</option>
            <option value="rounded">Rounded Corners</option>
          </select>

          {(theme.borderStyle || 'square') === 'rounded' && (
            <div>
              <label className={styles.designerFormLabel}>Corner Radius (px)</label>
              <input
                className={styles.designerInput}
                type="number"
                min={0}
                max={40}
                value={String(theme.borderRadius || 8)}
                onChange={(ev) => this.updateForm(function(nextForm) {
                  if (!nextForm.theme) nextForm.theme = {};
                  var parsedRadius = parseInt(ev.currentTarget.value, 10);
                  nextForm.theme.borderRadius = isNaN(parsedRadius) ? 8 : parsedRadius;
                  return nextForm;
                })}
              />
            </div>
          )}

          <label className={styles.designerFormLabel}>Border Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={theme.borderColor || '#cccccc'}
              onChange={(ev) => this.updateForm(function(nextForm) {
                if (!nextForm.theme) nextForm.theme = {};
                nextForm.theme.borderColor = ev.currentTarget.value || undefined;
                return nextForm;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={theme.borderColor || ''}
              placeholder="#cccccc"
              onChange={(ev) => this.updateForm(function(nextForm) {
                if (!nextForm.theme) nextForm.theme = {};
                nextForm.theme.borderColor = ev.currentTarget.value || undefined;
                return nextForm;
              })}
              style={{ flex: 1 }}
            />
          </div>

          <label className={styles.designerFormLabel}>Border Width (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={0}
            max={10}
            value={String(theme.borderWidth || 1)}
            onChange={(ev) => this.updateForm(function(nextForm) {
              if (!nextForm.theme) nextForm.theme = {};
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextForm.theme.borderWidth = isNaN(parsed) ? 1 : parsed;
              return nextForm;
            })}
          />
        </div>
      </div>
    );
  }

  private renderStepContainerSettings(): JSX.Element {
    var currentStep = this.getCurrentStep();
    if (!currentStep) {
      return <div />;
    }
    var theme = (currentStep.theme) || {};
    return (
      <div>
        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Container Font Settings</div>

          <label className={styles.designerFormLabel}>Font Size (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={8}
            max={72}
            value={String(theme.fontSize || 14)}
            onChange={(ev) => this.updateCurrentStep(function(nextStep) {
              if (!nextStep.theme) nextStep.theme = {};
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextStep.theme.fontSize = isNaN(parsed) ? undefined : parsed;
              return nextStep;
            })}
          />

          <label className={styles.designerFormLabel}>Font Family</label>
          <select
            className={styles.designerInput}
            value={theme.fontFamily || 'inherit'}
            onChange={(ev) => this.updateCurrentStep(function(nextStep) {
              if (!nextStep.theme) nextStep.theme = {};
              nextStep.theme.fontFamily = ev.currentTarget.value === 'inherit' ? undefined : ev.currentTarget.value;
              return nextStep;
            })}
          >
            <option value="inherit">Inherit</option>
            <option value="Arial">Arial</option>
            <option value="'Segoe UI'">Segoe UI</option>
            <option value="Verdana">Verdana</option>
            <option value="Georgia">Georgia</option>
            <option value="Courier New">Courier New</option>
            <option value="Times New Roman">Times New Roman</option>
          </select>

          <label className={styles.designerFormLabel}>Font Weight</label>
          <select
            className={styles.designerInput}
            value={theme.fontWeight || 'normal'}
            onChange={(ev) => this.updateCurrentStep(function(nextStep) {
              if (!nextStep.theme) nextStep.theme = {};
              nextStep.theme.fontWeight = ev.currentTarget.value as 'normal' | 'bold' | '500' | '600' | '700' || undefined;
              return nextStep;
            })}
          >
            <option value="normal">Normal</option>
            <option value="500">Medium (500)</option>
            <option value="600">Semi-Bold (600)</option>
            <option value="bold">Bold</option>
            <option value="700">Bold (700)</option>
          </select>

          <label className={styles.designerFormLabel}>Font Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={theme.color || '#000000'}
              onChange={(ev) => this.updateCurrentStep(function(nextStep) {
                if (!nextStep.theme) nextStep.theme = {};
                nextStep.theme.color = ev.currentTarget.value || undefined;
                return nextStep;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={theme.color || ''}
              placeholder="#000000"
              onChange={(ev) => this.updateCurrentStep(function(nextStep) {
                if (!nextStep.theme) nextStep.theme = {};
                nextStep.theme.color = ev.currentTarget.value || undefined;
                return nextStep;
              })}
              style={{ flex: 1 }}
            />
          </div>
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Container Background & Border</div>

          <label className={styles.designerFormLabel}>Background Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={theme.backgroundColor || '#ffffff'}
              onChange={(ev) => this.updateCurrentStep(function(nextStep) {
                if (!nextStep.theme) nextStep.theme = {};
                nextStep.theme.backgroundColor = ev.currentTarget.value || undefined;
                return nextStep;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={theme.backgroundColor || ''}
              placeholder="#ffffff"
              onChange={(ev) => this.updateCurrentStep(function(nextStep) {
                if (!nextStep.theme) nextStep.theme = {};
                nextStep.theme.backgroundColor = ev.currentTarget.value || undefined;
                return nextStep;
              })}
              style={{ flex: 1 }}
            />
          </div>

          <label className={styles.designerFormLabel}>Border Style</label>
          <select
            className={styles.designerInput}
            value={theme.borderStyle || 'square'}
            onChange={(ev) => this.updateCurrentStep(function(nextStep) {
              if (!nextStep.theme) nextStep.theme = {};
              nextStep.theme.borderStyle = ev.currentTarget.value as 'square' | 'rounded' || undefined;
              return nextStep;
            })}
          >
            <option value="square">Square Corners</option>
            <option value="rounded">Rounded Corners</option>
          </select>

          {(theme.borderStyle || 'square') === 'rounded' && (
            <div>
              <label className={styles.designerFormLabel}>Corner Radius (px)</label>
              <input
                className={styles.designerInput}
                type="number"
                min={0}
                max={40}
                value={String(theme.borderRadius || 8)}
                onChange={(ev) => this.updateCurrentStep(function(nextStep) {
                  if (!nextStep.theme) nextStep.theme = {};
                  var parsedRadius = parseInt(ev.currentTarget.value, 10);
                  nextStep.theme.borderRadius = isNaN(parsedRadius) ? 8 : parsedRadius;
                  return nextStep;
                })}
              />
            </div>
          )}

          <label className={styles.designerFormLabel}>Border Color</label>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <input
              type="color"
              value={theme.borderColor || '#cccccc'}
              onChange={(ev) => this.updateCurrentStep(function(nextStep) {
                if (!nextStep.theme) nextStep.theme = {};
                nextStep.theme.borderColor = ev.currentTarget.value || undefined;
                return nextStep;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={theme.borderColor || ''}
              placeholder="#cccccc"
              onChange={(ev) => this.updateCurrentStep(function(nextStep) {
                if (!nextStep.theme) nextStep.theme = {};
                nextStep.theme.borderColor = ev.currentTarget.value || undefined;
                return nextStep;
              })}
              style={{ flex: 1 }}
            />
          </div>

          <label className={styles.designerFormLabel}>Border Width (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={0}
            max={10}
            value={String(theme.borderWidth || 1)}
            onChange={(ev) => this.updateCurrentStep(function(nextStep) {
              if (!nextStep.theme) nextStep.theme = {};
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextStep.theme.borderWidth = isNaN(parsed) ? 1 : parsed;
              return nextStep;
            })}
          />
        </div>
      </div>
    );
  }

  private renderStepEditor(): JSX.Element {
    var currentStep = this.getCurrentStep();
    if (!currentStep) {
      return <div className={styles.designerPanel}><div className={styles.designerEmptyState}>{strings.DesignerStepNotFound}</div></div>;
    }

    var activeTab = this.state.activeDesignerTab;

    return (
      <div className={styles.designerPanel}>
        <div style={{ display: 'flex', gap: '4px', paddingBottom: '12px', borderBottom: '1px solid #e1dfdd', marginBottom: '4px' }}>
          <button
            className={activeTab === 'form' ? styles.designerTabActive : styles.designerTab}
            onClick={() => this.setState({ activeDesignerTab: 'form' })}
          >
            Form
          </button>
          <button
            className={activeTab === 'container' ? styles.designerTabActive : styles.designerTab}
            onClick={() => this.setState({ activeDesignerTab: 'container' })}
          >
            Container
          </button>
        </div>

        {activeTab === 'form' && <div>
        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>{strings.DesignerTitle}</div>

          <label className={styles.designerFormLabel}>{strings.DesignerFormName}</label>
          <input
            className={styles.designerInput}
            type="text"
            value={this.props.schema.name || ''}
            title={strings.DesignerFormName}
            placeholder={strings.DesignerFormName}
            onChange={(ev) => this.updateForm(function(schema) {
              schema.name = ev.currentTarget.value;
              return schema;
            })}
          />

          <label className={styles.designerCheckboxRow}>
            <input
              type="checkbox"
              checked={this.props.schema.showTitle !== false}
              onChange={(ev) => this.updateForm(function(schema) {
                schema.showTitle = ev.currentTarget.checked;
                return schema;
              })}
            />
            <span>Show form name</span>
          </label>

          <label className={styles.designerFormLabel}>{strings.DescriptionFieldLabel}</label>
          <textarea
            className={styles.designerTextarea}
            value={this.props.schema.description || ''}
            title={strings.DescriptionFieldLabel}
            placeholder={strings.DescriptionFieldLabel}
            onChange={(ev) => this.updateForm(function(schema) {
              schema.description = ev.currentTarget.value;
              return schema;
            })}
          />
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Form Layout</div>
          <div className={styles.designerPanelHint}>How containers (steps) are arranged in the form</div>

          <label className={styles.designerFormLabel}>Layout Style</label>
          <select
            className={styles.designerInput}
            value={(this.props.schema.theme && this.props.schema.theme.layout) || 'stack'}
            onChange={(ev) => this.updateForm(function(schema) {
              if (!schema.theme) schema.theme = {};
              schema.theme.layout = ev.currentTarget.value === 'grid' ? 'grid' : 'stack';
              if (schema.theme.layout !== 'grid') {
                schema.theme.columns = 1;
              } else if (!schema.theme.columns || schema.theme.columns < 2) {
                schema.theme.columns = 2;
              }
              return schema;
            })}
          >
            <option value="stack">{strings.DesignerLayoutStack}</option>
            <option value="grid">{strings.DesignerLayoutGrid}</option>
          </select>

          <label className={styles.designerFormLabel}>Number of Columns</label>
          <input
            className={styles.designerInput}
            type="number"
            min={1}
            max={4}
            value={String((this.props.schema.theme && this.props.schema.theme.columns) || 1)}
            onChange={(ev) => this.updateForm(function(schema) {
              if (!schema.theme) schema.theme = {};
              var parsed = parseInt(ev.currentTarget.value, 10);
              schema.theme.columns = isNaN(parsed) ? 1 : Math.max(1, Math.min(4, parsed));
              if (schema.theme.columns <= 1) {
                schema.theme.layout = 'stack';
              } else if (!schema.theme.layout) {
                schema.theme.layout = 'grid';
              }
              return schema;
            })}
          />
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Form Styling</div>
          <div className={styles.designerPanelHint}>Configure styling for the entire form wrapper</div>
        </div>

        {this.renderFormSettings()}
        </div>}

        {activeTab === 'container' && <div>
        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Container Details</div>

          <label className={styles.designerFormLabel}>{strings.DesignerStepTitlePlaceholder}</label>
          <input
            className={styles.designerInput}
            type="text"
            value={currentStep.title}
            title={strings.DesignerStepTitlePlaceholder}
            placeholder={strings.DesignerStepTitlePlaceholder}
            onChange={(ev) => this.updateCurrentStep(function(step) {
              step.title = ev.currentTarget.value;
              return step;
            })}
          />

          <label className={styles.designerCheckboxRow}>
            <input
              type="checkbox"
              checked={currentStep.showTitle !== false}
              onChange={(ev) => this.updateCurrentStep(function(step) {
                step.showTitle = ev.currentTarget.checked;
                return step;
              })}
            />
            <span>Show container name</span>
          </label>

          <label className={styles.designerFormLabel}>{strings.DesignerStepDescriptionPlaceholder}</label>
          <textarea
            className={styles.designerTextarea}
            value={currentStep.description || ''}
            title={strings.DesignerStepDescriptionPlaceholder}
            placeholder={strings.DesignerStepDescriptionPlaceholder}
            onChange={(ev) => this.updateCurrentStep(function(step) {
              step.description = ev.currentTarget.value;
              return step;
            })}
          />

          <label className={styles.designerCheckboxRow}>
            <input
              type="checkbox"
              checked={currentStep.visible !== false}
              onChange={(ev) => this.updateCurrentStep(function(step) {
                step.visible = ev.currentTarget.checked;
                return step;
              })}
            />
            <span>{strings.DesignerShowStep}</span>
          </label>

          <div className={styles.designerPanelHint}>{strings.DesignerShowStepHelp}</div>
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Container Layout</div>
          <div className={styles.designerPanelHint}>How fields are arranged within this container</div>

          <label className={styles.designerFormLabel}>{strings.DesignerStepLayout}</label>
          <select
            className={styles.designerInput}
            value={getStepTheme(this.props.schema, currentStep).layout}
            title={strings.DesignerStepLayout}
            onChange={(ev) => this.updateCurrentStep(function(step) {
              step.theme = step.theme || {};
              step.theme.layout = ev.currentTarget.value === 'grid' ? 'grid' : 'stack';
              if (step.theme.layout !== 'grid') {
                step.theme.columns = 1;
              } else if (!step.theme.columns || step.theme.columns < 2) {
                step.theme.columns = 2;
              }
              return step;
            })}
          >
            <option value="stack">{strings.DesignerLayoutStack}</option>
            <option value="grid">{strings.DesignerLayoutGrid}</option>
          </select>

          <label className={styles.designerFormLabel}>{strings.DesignerColumns}</label>
          <input
            className={styles.designerInput}
            type="number"
            min={1}
            max={4}
            value={String(getStepTheme(this.props.schema, currentStep).columns)}
            title={strings.DesignerColumns}
            onChange={(ev) => this.updateCurrentStep(function(step) {
              step.theme = step.theme || {};
              var parsed = parseInt(ev.currentTarget.value, 10);
              step.theme.columns = isNaN(parsed) ? 1 : Math.max(1, Math.min(4, parsed));
              if (step.theme.columns <= 1) {
                step.theme.layout = 'stack';
              } else if (!step.theme.layout) {
                step.theme.layout = 'grid';
              }
              return step;
            })}
          />
        </div>

        <div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Container Styling</div>
          <div className={styles.designerPanelHint}>Font, background and border for this container</div>
        </div>

        {this.renderStepContainerSettings()}
        </div>}
      </div>
    );
  }

  public render(): JSX.Element {
    var selectedField = this.getSelectedField();

    return (
      <div className={styles.designerWorkspace}>
        {this.renderPalette()}
        {this.renderCanvas()}
        {selectedField ? this.renderFieldEditor(selectedField) : this.renderStepEditor()}
      </div>
    );
  }
}