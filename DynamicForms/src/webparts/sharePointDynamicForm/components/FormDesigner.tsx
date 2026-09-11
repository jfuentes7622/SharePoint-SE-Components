import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './SharePointDynamicForm.module.scss';
import { ConditionalFieldRule, FormField, FormSchema, FormStep, SPFieldInfo, SPFieldType, FieldConfig, FieldType } from '../../../formEngine/core/types';
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
  reportListFields: Array<{ key: string; text: string }>;
  gridControlSources: Array<{ key: string; text: string; listName?: string }>;
  gridControlFields: Array<{ key: string; text: string }>;
  gridControlFieldsListName: string;
  gridControlFieldsLoading: boolean;
  gridControlFieldsError: string | null;
  imageLibraries: Array<{ key: string; text: string }>;
  imageFiles: Array<{ key: string; text: string }>;
  selectedImageLibrary: string;
  imageFilesLoading: boolean;
  imageBrowserError: string | null;
  loadingFields: boolean;
  fieldsError: string | null;
  selectedStepIndex: number;
  selectedFieldId: string | null;
  selectedConditionalRuleId: string | null;
  conditionalLookupOptions: Array<{ key: string; text: string }>;
  conditionalLookupFieldId: string;
  conditionalLookupLoading: boolean;
  activeDesignerTab: 'form' | 'container' | 'conditional';
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

// Fallback used whenever a level (form/container/field) does not customize its own font, so the
// designer preview never silently cascades styling in from a different (unrelated) level's setting.
var DEFAULT_THEME_FONT_FAMILY = "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif";

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
    case 'customimage':
      return 'Image';
    case 'divider':
      return 'Divider';
    case 'gridcontrol':
      return 'Grid Control';
    default:
      return type;
  }
}

export class FormDesigner extends React.Component<FormDesignerProps, FormDesignerState> {
  public constructor(props: FormDesignerProps) {
    super(props);
    this.state = {
      spFields: [],
      reportListFields: [],
      gridControlSources: [],
      gridControlFields: [],
      gridControlFieldsListName: '',
      gridControlFieldsLoading: false,
      gridControlFieldsError: null,
      imageLibraries: [],
      imageFiles: [],
      selectedImageLibrary: '',
      imageFilesLoading: false,
      imageBrowserError: null,
      loadingFields: false,
      fieldsError: null,
      selectedStepIndex: 0,
      selectedFieldId: null,
      selectedConditionalRuleId: null,
      conditionalLookupOptions: [],
      conditionalLookupFieldId: '',
      conditionalLookupLoading: false,
      activeDesignerTab: 'form',
    };

    this.handleAddRichText = this.handleAddRichText.bind(this);
    this.handleAddCustomImage = this.handleAddCustomImage.bind(this);
    this.handleAddDivider = this.handleAddDivider.bind(this);
    this.handleAddGridControl = this.handleAddGridControl.bind(this);
  }

  public componentDidMount(): void {
    this.loadFields();
    this.loadImageLibraries();
    this.loadGridControlSources();
  }

  public componentDidUpdate(prevProps: FormDesignerProps, prevState: FormDesignerState): void {
    if (prevProps.listName !== this.props.listName) {
      this.loadFields();
    }

    if (this.state.selectedStepIndex >= this.props.schema.steps.length) {
      this.setState({ selectedStepIndex: Math.max(0, this.props.schema.steps.length - 1), selectedFieldId: null });
    }

    if (prevState.selectedFieldId !== this.state.selectedFieldId) {
      this.loadSelectedGridControlFields();
    }

    if (prevState.selectedConditionalRuleId !== this.state.selectedConditionalRuleId) {
      var selectedRule = this.getSelectedConditionalRule();
      this.loadConditionalLookupOptions(selectedRule ? selectedRule.sourceField : '');
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
      this.setState({ spFields: [], reportListFields: [], loadingFields: false, fieldsError: strings.DesignerNoListSelected });
      return;
    }

    this.setState({ loadingFields: true, fieldsError: null });
    try {
      var listResponse = await this.getWithAcceptFallback(
        this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')?$select=BaseType"
      );
      if (!listResponse.ok) {
        throw new Error(strings.DesignerLoadFieldsFailed);
      }
      var listData = await listResponse.json();
      var listInfo = listData && listData.d ? listData.d : listData;
      var isDocumentLibrary = Number(listInfo && listInfo.BaseType) === 1;
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
      var reportListFields = fields
        .filter(function(field) { return !!field.InternalName; })
        .map(function(field) {
          return { key: String(field.InternalName), text: String(field.Title || field.InternalName) + ' (' + String(field.InternalName) + ')' };
        });
      var mapped = fields
        .filter(function(field) {
          if (isDocumentLibrary && field.InternalName === 'FileLeafRef') {
            return false;
          }
          var isAttachmentField = field.InternalName === 'Attachments';
          var isCommonSystemField = field.InternalName === 'Author' || field.InternalName === 'Editor'
            || field.InternalName === 'Created' || field.InternalName === 'Modified';
          if (field.Hidden) {
            return false;
          }
          if (field.FromBaseType && field.InternalName !== 'Title' && !isAttachmentField && !isCommonSystemField) {
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

      if (isDocumentLibrary) {
        mapped.unshift({
          id: '__dynamicforms_document__',
          internalName: 'FileLeafRef',
          title: strings.PropertyFieldTypeDocument,
          description: strings.PropertyFieldTypeDocumentDescription,
          type: SPFieldType.Attachments,
          required: true,
          readOnly: false,
        } as SPFieldInfo);
      }

      this.setState({ spFields: mapped, reportListFields: reportListFields, loadingFields: false, fieldsError: null });
    } catch (error) {
      this.setState({
        spFields: [],
        reportListFields: [],
        loadingFields: false,
        fieldsError: error && error.message ? error.message : strings.DesignerLoadFieldsFailed,
      });
    }
  }

  private async loadImageLibraries(): Promise<void> {
    try {
      var response = await this.getWithAcceptFallback(this.getWebUrl() + '/_api/web/lists?$select=Title,BaseTemplate,Hidden&$filter=Hidden eq false');
      if (!response.ok) {
        throw new Error('Unable to load image libraries.');
      }
      var data = await response.json();
      var lists: any[] = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      var libraries = lists
        .filter(function(list) { return list.BaseTemplate === 101 || list.BaseTemplate === 109; })
        .map(function(list) { return { key: String(list.Title), text: String(list.Title) }; });
      this.setState({ imageLibraries: libraries, imageBrowserError: null });
    } catch (error) {
      this.setState({ imageLibraries: [], imageBrowserError: error && error.message ? error.message : 'Unable to load image libraries.' });
    }
  }

  private loadGridControlSources(): void {
    var provider = this.props.context.dynamicDataProvider || this.props.context._dynamicDataProvider;
    var options: Array<{ key: string; text: string; listName?: string }> = [];
    if (provider && provider.getAvailableSources) {
      var sources = provider.getAvailableSources() || [];
      for (var i = 0; i < sources.length; i += 1) {
        var source = sources[i];
        var metadata = source && source.metadata;
        var componentId = metadata && metadata.componentId ? String(metadata.componentId).toLowerCase() : '';
        var alias = metadata && metadata.alias ? String(metadata.alias).toLowerCase() : '';
        if (componentId !== 'de5f92e8-570c-4a7c-ba96-7ae89a098723' && alias !== 'gridcontrolwebpart') {
          continue;
        }
        var instanceName = '';
        var listName = '';
        if (source.getPropertyValue) {
          try { instanceName = String(source.getPropertyValue('instanceName') || '').trim(); } catch (_instanceNameError) { instanceName = ''; }
          try { listName = String(source.getPropertyValue('listName') || '').trim(); } catch (_listNameError) { listName = ''; }
        }
        options.push({
          key: String(source.id),
          text: String(instanceName || metadata.title || metadata.alias || source.id) + ' (' + String(source.id) + ')',
          listName: listName || undefined
        });
      }
    }
    this.setState({ gridControlSources: options }, () => this.loadSelectedGridControlFields());
  }

  private loadSelectedGridControlFields(): void {
    var selectedField = this.getSelectedField();
    var sourceId = selectedField && selectedField.type === 'gridcontrol' && selectedField.config
      ? String(selectedField.config.gridControlSourceId || '') : '';
    var listName = '';
    for (var i = 0; i < this.state.gridControlSources.length; i += 1) {
      if (this.state.gridControlSources[i].key === sourceId) {
        listName = String(this.state.gridControlSources[i].listName || '');
        break;
      }
    }
    if (!listName) {
      this.setState({ gridControlFields: [], gridControlFieldsListName: '', gridControlFieldsLoading: false, gridControlFieldsError: sourceId ? 'The selected Grid Control did not provide its configured list name.' : null });
      return;
    }
    if (this.state.gridControlFieldsListName === listName && (this.state.gridControlFieldsLoading || this.state.gridControlFields.length > 0)) {
      return;
    }
    this.loadGridControlFields(listName);
  }

  private async loadGridControlFields(listName: string): Promise<void> {
    this.setState({ gridControlFields: [], gridControlFieldsListName: listName, gridControlFieldsLoading: true, gridControlFieldsError: null });
    try {
      var response = await this.getWithAcceptFallback(
        this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/fields?$select=InternalName,Title,Hidden&$filter=Hidden eq false"
      );
      if (!response.ok) { throw new Error('Unable to load the selected Grid Control columns.'); }
      var data = await response.json();
      var fields: any[] = data && data.value ? data.value : [];
      if (!fields || fields.length === 0) { fields = data && data.d && data.d.results ? data.d.results : []; }
      var options = fields
        .filter(function(item) { return !!item.InternalName && item.InternalName !== 'Attachments'; })
        .map(function(item) { return { key: String(item.InternalName), text: String(item.Title || item.InternalName) + ' (' + String(item.InternalName) + ')' }; });
      this.setState({ gridControlFields: options, gridControlFieldsLoading: false, gridControlFieldsError: null });
    } catch (error) {
      this.setState({
        gridControlFields: [],
        gridControlFieldsLoading: false,
        gridControlFieldsError: error && error.message ? error.message : 'Unable to load the selected Grid Control columns.'
      });
    }
  }

  private async loadImageFiles(libraryName: string): Promise<void> {
    if (!libraryName) {
      this.setState({ selectedImageLibrary: '', imageFiles: [], imageFilesLoading: false, imageBrowserError: null });
      return;
    }
    this.setState({ selectedImageLibrary: libraryName, imageFiles: [], imageFilesLoading: true, imageBrowserError: null });
    try {
      var response = await this.getWithAcceptFallback(
        this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(libraryName) + "')/items?$select=FileLeafRef,FileRef,FSObjType&$filter=FSObjType eq 0&$top=5000"
      );
      if (!response.ok) {
        throw new Error('Unable to load images from the selected library.');
      }
      var data = await response.json();
      var items: any[] = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      var imagePattern = /\.(apng|avif|bmp|gif|jpe?g|png|svg|webp)$/i;
      var files = items
        .filter(function(item) { return imagePattern.test(String(item.FileLeafRef || '')) && !!item.FileRef; })
        .map(function(item) { return { key: String(item.FileRef), text: String(item.FileLeafRef) }; });
      this.setState({ imageFiles: files, imageFilesLoading: false, imageBrowserError: null });
    } catch (error) {
      this.setState({ imageFiles: [], imageFilesLoading: false, imageBrowserError: error && error.message ? error.message : 'Unable to load images from the selected library.' });
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

  private getConditionalRuleFields(): FormField[] {
    var fields: FormField[] = [];
    for (var stepIndex = 0; stepIndex < this.props.schema.steps.length; stepIndex += 1) {
      var stepFields = this.props.schema.steps[stepIndex].fields;
      for (var fieldIndex = 0; fieldIndex < stepFields.length; fieldIndex += 1) {
        if (stepFields[fieldIndex] && stepFields[fieldIndex]!.type !== 'newline') {
          fields.push(stepFields[fieldIndex]!);
        }
      }
    }
    return fields;
  }

  private getConditionalField(fieldReference: string): FormField | null {
    var normalized = String(fieldReference || '').toLowerCase();
    var fields = this.getConditionalRuleFields();
    for (var i = 0; i < fields.length; i += 1) {
      if (String(fields[i].id || '').toLowerCase() === normalized || String(fields[i].fieldName || '').toLowerCase() === normalized) {
        return fields[i];
      }
    }
    return null;
  }

  private getSelectedConditionalRule(): ConditionalFieldRule | null {
    var rules = this.props.schema.conditionalRules || [];
    for (var i = 0; i < rules.length; i += 1) {
      if (rules[i].id === this.state.selectedConditionalRuleId) { return rules[i]; }
    }
    return null;
  }

  private async loadConditionalLookupOptions(fieldReference: string): Promise<void> {
    var field = this.getConditionalField(fieldReference);
    if (!field || field.type !== 'lookup' || !field.config || !field.config.lookupList) {
      this.setState({ conditionalLookupOptions: [], conditionalLookupFieldId: '', conditionalLookupLoading: false });
      return;
    }
    this.setState({ conditionalLookupOptions: [], conditionalLookupFieldId: field.id, conditionalLookupLoading: true });
    try {
      var lookupList = String(field.config.lookupList);
      var lookupField = String(field.config.lookupField || 'Title');
      var cleanLookupList = lookupList.replace(/^\{|\}$/g, '');
      var listPath = /^\{?[0-9a-f]{8}-[0-9a-f-]{27}\}?$/i.test(lookupList)
        ? "/_api/web/lists(guid'" + cleanLookupList + "')"
        : "/_api/web/lists/getByTitle('" + escapeODataText(lookupList) + "')";
      var response = await this.getWithAcceptFallback(this.getWebUrl() + listPath + '/items?$select=' + encodeURIComponent('Id,' + lookupField) + '&$top=5000');
      if (!response.ok) { throw new Error('Unable to load lookup items.'); }
      var data = await response.json();
      var items: any[] = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      var options = items.map(function(item) {
        var itemId = item.Id !== undefined ? item.Id : item.ID;
        return { key: String(itemId), text: String(item[lookupField] || item.Title || itemId) };
      });
      if (this.state.conditionalLookupFieldId === field.id) {
        this.setState({ conditionalLookupOptions: options, conditionalLookupLoading: false });
      }
    } catch (_error) {
      if (this.state.conditionalLookupFieldId === field.id) {
        this.setState({ conditionalLookupOptions: [], conditionalLookupLoading: false });
      }
    }
  }

  private addConditionalRule(): void {
    var fields = this.getConditionalRuleFields();
    if (fields.length === 0) { return; }
    var ruleId = createId('condition');
    this.updateForm(function(schema) {
      schema.conditionalRules = schema.conditionalRules || [];
      schema.conditionalRules.push({
        id: ruleId,
        enabled: true,
        sourceField: fields[0].id,
        operator: 'eq',
        value: '',
        targetField: fields.length > 1 ? fields[1].id : fields[0].id,
        targetFields: [fields.length > 1 ? fields[1].id : fields[0].id],
        action: 'style',
        visible: true,
        style: { backgroundColor: '#fff1f0', color: '#000000', borderColor: '#d13438', borderWidth: 1 }
      });
      return schema;
    });
    this.setState({ selectedConditionalRuleId: ruleId });
  }

  private updateConditionalRule(ruleId: string, mutator: (rule: ConditionalFieldRule) => void): void {
    this.updateForm(function(schema) {
      var rules = schema.conditionalRules || [];
      for (var i = 0; i < rules.length; i += 1) {
        if (rules[i].id === ruleId) { mutator(rules[i]); break; }
      }
      return schema;
    });
  }

  private removeConditionalRule(ruleId: string): void {
    this.updateForm(function(schema) {
      schema.conditionalRules = (schema.conditionalRules || []).filter(function(rule) { return rule.id !== ruleId; });
      return schema;
    });
    if (this.state.selectedConditionalRuleId === ruleId) { this.setState({ selectedConditionalRuleId: null }); }
  }

  private moveConditionalRule(ruleId: string, direction: number): void {
    this.updateForm(function(schema) {
      var rules = schema.conditionalRules || [];
      var index = rules.map(function(rule) { return rule.id; }).indexOf(ruleId);
      var targetIndex = index + direction;
      if (index >= 0 && targetIndex >= 0 && targetIndex < rules.length) {
        var moved = rules.splice(index, 1)[0];
        rules.splice(targetIndex, 0, moved);
      }
      return schema;
    });
  }

  private getConditionalFieldLabel(fieldReference: string): string {
    var normalized = String(fieldReference || '').toLowerCase();
    var fields = this.getConditionalRuleFields();
    for (var i = 0; i < fields.length; i += 1) {
      if (String(fields[i].id || '').toLowerCase() === normalized || String(fields[i].fieldName || '').toLowerCase() === normalized) {
        return fields[i].label;
      }
    }
    return fieldReference || '(field)';
  }

  private getConditionalRuleTargets(rule: ConditionalFieldRule): string[] {
    var configuredTargets = rule.targetFields && rule.targetFields.length > 0
      ? rule.targetFields
      : [rule.targetField];
    var targets: string[] = [];
    var seen: { [target: string]: boolean } = {};
    for (var i = 0; i < configuredTargets.length; i += 1) {
      var target = String(configuredTargets[i] || '').trim();
      var targetKey = target.toLowerCase();
      if (target && !seen[targetKey]) {
        seen[targetKey] = true;
        targets.push(target);
      }
    }
    return targets;
  }

  private getConditionalRuleTargetLabels(rule: ConditionalFieldRule): string {
    return this.getConditionalRuleTargets(rule).map((target) => this.getConditionalFieldLabel(target)).join(', ');
  }

  private getConditionalRuleValueLabel(rule: ConditionalFieldRule): string {
    if (rule.operator === 'between' && rule.value && typeof rule.value === 'object') {
      return String(rule.value.from || '') + ' through ' + String(rule.value.to || '');
    }
    return String(rule.value === undefined || rule.value === null ? '' : rule.value);
  }

  private renderConditionalRulesCanvas(): JSX.Element {
    var rules = this.props.schema.conditionalRules || [];
    return (
      <div className={styles.designerCanvas}>
        <div className={styles.designerTabs}>
          <button type="button" className={styles.designerTab} onClick={() => this.setState({ activeDesignerTab: 'form', selectedConditionalRuleId: null })}>Fields</button>
          <button type="button" className={styles.designerTabActive}>Conditional styling</button>
          <button type="button" className={styles.designerActionButton} disabled={this.getConditionalRuleFields().length === 0} onClick={() => this.addConditionalRule()}>Add rule</button>
        </div>
        <div className={styles.designerCanvasBody}>
          <div className={styles.designerStepHeader}><div className={styles.designerStepTitle}>Conditional styling and visibility</div><div className={styles.designerStepDescription}>Rules are applied from top to bottom. When rules conflict, the last matching rule wins.</div></div>
          {rules.length === 0 && <div className={styles.designerEmptyState}>No conditional rules configured.</div>}
          {rules.map((rule, index) => (
            <div key={rule.id} className={rule.id === this.state.selectedConditionalRuleId ? styles.designerFieldCardActive : styles.designerFieldCard}>
              <div className={styles.designerFieldInfo}>
                <div className={styles.designerFieldLabel}>{'Rule ' + String(index + 1) + ': If ' + this.getConditionalFieldLabel(rule.sourceField) + ' ' + rule.operator + ' "' + this.getConditionalRuleValueLabel(rule) + '"'}</div>
                <div className={styles.designerFieldMeta}>{(rule.action === 'visibility'
                  ? 'Set ' + this.getConditionalRuleTargetLabels(rule) + ' ' + (rule.visible === false ? 'hidden' : 'visible')
                  : rule.action === 'disable'
                    ? 'Set ' + this.getConditionalRuleTargetLabels(rule) + ' ' + (rule.disabled === false ? 'enabled' : 'disabled')
                    : 'Apply style to ' + this.getConditionalRuleTargetLabels(rule)) + (rule.enabled === false ? ' (disabled)' : '')}</div>
              </div>
              <div className={styles.designerFieldActions}>
                <button type="button" className={styles.designerInlineButton} onClick={() => this.moveConditionalRule(rule.id, -1)} disabled={index === 0}>Move up</button>
                <button type="button" className={styles.designerInlineButton} onClick={() => this.moveConditionalRule(rule.id, 1)} disabled={index === rules.length - 1}>Move down</button>
                <button type="button" className={styles.designerInlineButton} onClick={() => this.setState({ selectedConditionalRuleId: rule.id })}>Edit</button>
                <button type="button" className={styles.designerInlineButtonDanger} onClick={() => this.removeConditionalRule(rule.id)}>Delete</button>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  private renderConditionalRules(): JSX.Element {
    var fields = this.getConditionalRuleFields();
    var rules = this.props.schema.conditionalRules || [];
    var selectedRule: ConditionalFieldRule | null = null;
    for (var ruleIndex = 0; ruleIndex < rules.length; ruleIndex += 1) {
      if (rules[ruleIndex].id === this.state.selectedConditionalRuleId) { selectedRule = rules[ruleIndex]; break; }
    }
    var fieldOptions = fields.map(function(field) {
      return <option key={field.id} value={field.id}>{field.label + ' (' + field.fieldName + ')'}</option>;
    });
    if (!selectedRule) {
      return <div className={styles.designerPanel}><div className={styles.designerPanelSection}><div className={styles.designerPanelTitle}>Rule editor</div><div className={styles.designerPanelHint}>Select a rule in the center panel to edit it.</div></div></div>;
    }
    var rule = selectedRule;
    var style = rule.style || {};
    var selectedTargetFields = this.getConditionalRuleTargets(rule);
    var sourceField = this.getConditionalField(rule.sourceField);
    var sourceChoices = sourceField && sourceField.config && sourceField.config.choices ? sourceField.config.choices : [];
    var isDateSource = !!sourceField && sourceField.type === 'datetime';
    var dateRangeValue: any = rule.value && typeof rule.value === 'object' ? rule.value : { from: '', to: '' };
    return (
      <div className={styles.designerPanel}><div className={styles.designerPanelSection}>
              <div className={styles.designerPanelTitle}>Edit conditional rule</div>
              <label className={styles.designerCheckboxRow}><input type="checkbox" checked={rule.enabled !== false} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.enabled = ev.currentTarget.checked; })} /><span>Enabled</span></label>
              <label className={styles.designerFormLabel}>If field</label>
              <select className={styles.designerInput} value={rule.sourceField} onChange={(ev) => {
                var sourceFieldId = ev.currentTarget.value;
                this.updateConditionalRule(rule.id, function(next) { next.sourceField = sourceFieldId; next.value = ''; if (next.operator === 'between') { next.operator = 'eq'; } });
                this.loadConditionalLookupOptions(sourceFieldId);
              }}>{fieldOptions}</select>
              <label className={styles.designerFormLabel}>Operator</label>
              <select className={styles.designerInput} value={rule.operator} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.operator = ev.currentTarget.value as any; })}>
                <option value="eq">Equals</option><option value="ne">Does not equal</option><option value="gt">Greater than</option><option value="ge">Greater than or equal</option><option value="lt">Less than</option><option value="le">Less than or equal</option><option value="contains">Contains</option><option value="notcontains">Does not contain</option><option value="startswith">Starts with</option><option value="endswith">Ends with</option>{isDateSource && <option value="between">Between (inclusive)</option>}
              </select>
              <label className={styles.designerFormLabel}>Value</label>
              {isDateSource && rule.operator === 'between' ? <div>
                <input className={styles.designerInput} type="date" value={String(dateRangeValue.from || '')} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.value = { from: ev.currentTarget.value, to: dateRangeValue.to || '' }; })} />
                <label className={styles.designerFormLabel}>Through</label>
                <input className={styles.designerInput} type="date" value={String(dateRangeValue.to || '')} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.value = { from: dateRangeValue.from || '', to: ev.currentTarget.value }; })} />
              </div> : sourceField && sourceField.type === 'lookup' ? <select className={styles.designerInput} value={String(rule.value || '')} disabled={this.state.conditionalLookupLoading} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.value = ev.currentTarget.value; })}><option value="">{this.state.conditionalLookupLoading ? 'Loading lookup items...' : 'Select an item'}</option>{this.state.conditionalLookupOptions.map(function(option) { return <option key={option.key} value={option.key}>{option.text}</option>; })}</select>
                : sourceField && (sourceField.type === 'dropdown' || sourceField.type === 'multiselect') ? <select className={styles.designerInput} value={String(rule.value || '')} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.value = ev.currentTarget.value; })}><option value="">Select a choice</option>{sourceChoices.map(function(choice) { return <option key={choice} value={choice}>{choice}</option>; })}</select>
                  : sourceField && sourceField.type === 'boolean' ? <select className={styles.designerInput} value={String(rule.value === true || rule.value === 'true' ? 'true' : rule.value === false || rule.value === 'false' ? 'false' : '')} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.value = ev.currentTarget.value === 'true'; })}><option value="">Select true or false</option><option value="true">True</option><option value="false">False</option></select>
                    : isDateSource ? <input className={styles.designerInput} type="date" value={String(rule.value || '')} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.value = ev.currentTarget.value; })} />
                      : <input className={styles.designerInput} type={sourceField && sourceField.type === 'number' ? 'number' : 'text'} value={String(rule.value === undefined || rule.value === null ? '' : rule.value)} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.value = ev.currentTarget.value; })} />}
              <label className={styles.designerFormLabel}>Target fields</label>
              <div className={styles.designerPanelHint}>Select one or more fields. Hold Ctrl while clicking to select multiple fields.</div>
              <select className={styles.designerInput} multiple={true} size={Math.min(8, Math.max(3, fields.length))} value={selectedTargetFields} onChange={(ev) => {
                var selectedTargets: string[] = [];
                for (var optionIndex = 0; optionIndex < ev.currentTarget.options.length; optionIndex += 1) {
                  var option = ev.currentTarget.options[optionIndex];
                  if (option.selected) { selectedTargets.push(option.value); }
                }
                this.updateConditionalRule(rule.id, function(next) {
                  next.targetFields = selectedTargets;
                  next.targetField = selectedTargets.length > 0 ? selectedTargets[0] : '';
                });
              }}>{fieldOptions}</select>
              <label className={styles.designerFormLabel}>Action</label>
              <select className={styles.designerInput} value={rule.action} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.action = ev.currentTarget.value as any; })}>
                <option value="style">Apply style</option><option value="visibility">Set visibility</option><option value="disable">Set disabled state</option>
              </select>
              {rule.action === 'visibility' ? <div>
                <label className={styles.designerFormLabel}>Visibility when condition is met</label>
                <select className={styles.designerInput} value={rule.visible === false ? 'hidden' : 'visible'} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.visible = ev.currentTarget.value !== 'hidden'; })}><option value="visible">Visible</option><option value="hidden">Hidden</option></select>
              </div> : rule.action === 'disable' ? <div>
                <label className={styles.designerFormLabel}>Field state when condition is met</label>
                <select className={styles.designerInput} value={rule.disabled === false ? 'enabled' : 'disabled'} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.disabled = ev.currentTarget.value !== 'enabled'; })}><option value="disabled">Disabled</option><option value="enabled">Enabled</option></select>
              </div> : <div>
                <label className={styles.designerFormLabel}>Background color</label><input className={styles.designerInput} type="color" value={style.backgroundColor || '#ffffff'} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.style = Object.assign({}, next.style || {}, { backgroundColor: ev.currentTarget.value }); })} />
                <label className={styles.designerFormLabel}>Text color</label><input className={styles.designerInput} type="color" value={style.color || '#000000'} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.style = Object.assign({}, next.style || {}, { color: ev.currentTarget.value }); })} />
                <label className={styles.designerFormLabel}>Border color</label><input className={styles.designerInput} type="color" value={style.borderColor || '#cccccc'} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.style = Object.assign({}, next.style || {}, { borderColor: ev.currentTarget.value }); })} />
                <label className={styles.designerFormLabel}>Border width</label><input className={styles.designerInput} type="number" min={0} max={10} value={String(style.borderWidth === undefined ? 1 : style.borderWidth)} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.style = Object.assign({}, next.style || {}, { borderWidth: parseInt(ev.currentTarget.value, 10) || 0 }); })} />
                <label className={styles.designerFormLabel}>Corner radius</label><input className={styles.designerInput} type="number" min={0} max={40} value={String(style.borderRadius || 0)} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.style = Object.assign({}, next.style || {}, { borderRadius: parseInt(ev.currentTarget.value, 10) || 0 }); })} />
                <label className={styles.designerFormLabel}>Font size</label><input className={styles.designerInput} type="number" min={8} max={48} value={String(style.fontSize || 14)} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.style = Object.assign({}, next.style || {}, { fontSize: parseInt(ev.currentTarget.value, 10) || 14 }); })} />
                <label className={styles.designerCheckboxRow}><input type="checkbox" checked={style.fontWeight === 'bold'} onChange={(ev) => this.updateConditionalRule(rule.id, function(next) { next.style = Object.assign({}, next.style || {}, { fontWeight: ev.currentTarget.checked ? 'bold' : 'normal' }); })} /><span>Bold</span></label>
              </div>}
      </div></div>
    );
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

  private moveStep(direction: number): void {
    var nextSchema = copySchema(this.props.schema);
    var steps = nextSchema.steps;
    var index = this.state.selectedStepIndex;
    var targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= steps.length) {
      return;
    }

    var temp = steps[index];
    steps[index] = steps[targetIndex];
    steps[targetIndex] = temp;
    this.props.onChange(nextSchema);
    this.setState({ selectedStepIndex: targetIndex });
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

  private handleAddCustomImage(): void {
    var nextField: FormField = {
      id: createId('field'), type: 'customimage', label: 'Image', fieldName: createId('customimage'),
      required: false, readOnly: true, disabled: false, columnSpan: 1,
      config: { imageUrl: '', imageAltText: '', imageFit: 'contain', imageAlignment: 'left' }
    };
    var nextSchema = copySchema(this.props.schema);
    nextSchema.steps[this.state.selectedStepIndex].fields.push(nextField);
    this.updateSchema(nextSchema, nextField.id);
  }

  private handleAddDivider(): void {
    var nextField: FormField = {
      id: createId('field'), type: 'divider', label: 'Divider', fieldName: createId('divider'),
      required: false, readOnly: true, disabled: false, columnSpan: 1,
      config: { dividerColor: '#c8c6c4', dividerThickness: 1, dividerStyle: 'solid', dividerSpacing: 16, dividerOrientation: 'horizontal', dividerLength: 80 }
    };
    var nextSchema = copySchema(this.props.schema);
    nextSchema.steps[this.state.selectedStepIndex].fields.push(nextField);
    this.updateSchema(nextSchema, nextField.id);
  }

  private handleAddGridControl(): void {
    var nextField: FormField = {
      id: createId('field'), type: 'gridcontrol', label: 'Grid Control', fieldName: createId('gridcontrol'),
      required: false, readOnly: false, disabled: false, columnSpan: 1,
      config: { gridControlFilterOperator: 'eq' }
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

    // Form-level wrapper style; font settings always resolve to their own default when unset so the
    // preview never shows a container/field silently inheriting the form's font styling.
    var formWrapperStyle: React.CSSProperties = { padding: '12px' };
    formWrapperStyle.fontSize = (schema.theme && schema.theme.fontSize) || 14;
    formWrapperStyle.fontFamily = (schema.theme && schema.theme.fontFamily) || DEFAULT_THEME_FONT_FAMILY;
    formWrapperStyle.fontWeight = ((schema.theme && schema.theme.fontWeight) || 'normal') as any;
    formWrapperStyle.color = (schema.theme && schema.theme.color) || '#000000';
    if (schema.theme) {
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
    var fieldVerticalSpacing = schema.theme && schema.theme.fieldVerticalSpacing !== undefined
      ? Math.max(0, Math.min(100, schema.theme.fieldVerticalSpacing))
      : 16;
    var formLayoutStyle: React.CSSProperties = formGridLayout
      ? { display: 'grid', gridTemplateColumns: 'repeat(' + String(formColumns) + ', minmax(0, 1fr))', gap: '16px' }
      : {};

    var visibleSteps = schema.steps.filter(function(step) { return step.visible !== false; });

    // Build description style with separate font settings; always resolved (own default when unset)
    var descriptionStyle: React.CSSProperties = { marginBottom: '16px' };
    descriptionStyle.textAlign = schema.descriptionAlignment || 'left';
    descriptionStyle.fontSize = (schema.theme && schema.theme.descriptionFontSize) || 14;
    descriptionStyle.fontFamily = (schema.theme && schema.theme.descriptionFontFamily) || DEFAULT_THEME_FONT_FAMILY;
    descriptionStyle.fontWeight = ((schema.theme && schema.theme.descriptionFontWeight) || 'normal') as any;
    descriptionStyle.color = (schema.theme && schema.theme.descriptionColor) || '#666666';

    return (
      <div className={styles.designerPreviewSection}>
        <div className={styles.designerPreviewTitle}>Preview</div>
        <div style={formWrapperStyle}>
          {schema.showTitle !== false && (schema.name || schema.description || schema.logoUrl) && <div style={{ display: 'flex', alignItems: 'flex-start', gap: '16px', width: '100%' }}>
            <div style={{ flex: '1 1 auto', minWidth: 0 }}>
              {schema.name && <h2 className={styles.designerPreviewFormName} style={{
                fontSize: formWrapperStyle.fontSize,
                fontFamily: formWrapperStyle.fontFamily,
                fontWeight: formWrapperStyle.fontWeight,
                color: formWrapperStyle.color,
                textAlign: schema.nameAlignment || 'left'
              }}>{schema.name}</h2>}
              {schema.description && <div style={descriptionStyle}>{schema.description}</div>}
            </div>
            {schema.logoUrl && <img src={schema.logoUrl} alt={schema.logoAltText || ''} style={{ flex: '0 0 auto', maxWidth: '160px', maxHeight: '80px', objectFit: 'contain' }} />}
          </div>}
          <div style={formLayoutStyle}>
            {visibleSteps.map((step) => {
              // Step container style
              var stepTheme = getStepTheme(schema, step);
              var previewGrid = stepTheme.layout === 'grid' && stepTheme.columns > 1;
              var stepContainerStyle: React.CSSProperties = previewGrid
                ? { display: 'grid', gridTemplateColumns: 'repeat(' + String(stepTheme.columns) + ', minmax(0, 1fr))', columnGap: '12px', rowGap: String(fieldVerticalSpacing) + 'px' }
                : {};

              // Font settings always resolve to their own default when unset so the preview never
              // shows a container silently inheriting the form's font styling.
              stepContainerStyle.fontSize = (step.theme && step.theme.fontSize) || 14;
              stepContainerStyle.fontFamily = (step.theme && step.theme.fontFamily) || DEFAULT_THEME_FONT_FAMILY;
              stepContainerStyle.fontWeight = ((step.theme && step.theme.fontWeight) || 'normal') as any;
              stepContainerStyle.color = (step.theme && step.theme.color) || '#000000';

              if (step.theme) {
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

              // Container description has its own independent font settings, separate from the
              // container's title/field styling above.
              var stepDescriptionStyle: React.CSSProperties = {
                fontSize: (step.theme && step.theme.descriptionFontSize) || 14,
                fontFamily: (step.theme && step.theme.descriptionFontFamily) || DEFAULT_THEME_FONT_FAMILY,
                fontWeight: ((step.theme && step.theme.descriptionFontWeight) || 'normal') as any,
                color: (step.theme && step.theme.descriptionColor) || '#666666',
              };

              return (
                <div key={step.id} style={{ marginBottom: formGridLayout ? '0' : '20px' }}>
                  {step.showTitle !== false && step.title && <div className={styles.designerPreviewStepTitle} style={{
                    fontSize: stepContainerStyle.fontSize,
                    fontFamily: stepContainerStyle.fontFamily,
                    fontWeight: stepContainerStyle.fontWeight,
                    color: stepContainerStyle.color,
                    textAlign: step.titleAlignment || 'left'
                  }}>{step.title}</div>}
                  {step.showTitle !== false && step.description && <div className={styles.designerPreviewStepDescription} style={{
                    fontSize: stepDescriptionStyle.fontSize,
                    fontFamily: stepDescriptionStyle.fontFamily,
                    fontWeight: stepDescriptionStyle.fontWeight,
                    color: stepDescriptionStyle.color,
                    textAlign: step.descriptionAlignment || 'left'
                  }}>{step.description}</div>}
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
                          : { marginBottom: String(fieldVerticalSpacing) + 'px' };
                        var richTextStyle: React.CSSProperties = {
                          border: '1px dashed #c8c6c4',
                          borderRadius: '0px',
                          padding: '12px',
                          background: '#faf9f8',
                          wordBreak: 'break-word',
                          lineHeight: '1.6',
                          fontSize: '12px',
                          fontFamily: DEFAULT_THEME_FONT_FAMILY,
                          fontWeight: 'normal',
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

                      if (field.type === 'gridcontrol') {
                        var gridControlCellStyle: React.CSSProperties = previewGrid
                          ? (field.startNewRow === true ? { gridColumn: '1 / -1' } : { gridColumn: 'span ' + String(Math.max(1, field.columnSpan || 1)) })
                          : { marginBottom: String(fieldVerticalSpacing) + 'px' };
                        var gridControlName = field.config && field.config.gridControlSourceName ? field.config.gridControlSourceName : 'No Grid Control selected';
                        return (
                          <div key={field.id} style={gridControlCellStyle}>
                            <div style={{ border: '1px dashed #00796b', padding: '12px', background: '#f2f8f7' }}>
                              <div style={{ fontWeight: 600 }}>{field.label}</div>
                              <div style={{ marginTop: '4px', fontSize: '12px', color: '#605e5c' }}>Embedded Grid Control: {gridControlName}</div>
                            </div>
                          </div>
                        );
                      }

                      if (field.type === 'customimage' || field.type === 'divider') {
                        var customCellStyle: React.CSSProperties = previewGrid
                          ? (field.startNewRow === true ? { gridColumn: '1 / -1' } : { gridColumn: 'span ' + String(Math.max(1, field.columnSpan || 1)) })
                          : { marginBottom: String(fieldVerticalSpacing) + 'px' };
                        var customConfig = field.config || {};
                        if (field.type === 'customimage') {
                          var previewImageUrl = String(customConfig.imageUrl || '').trim();
                          return (
                            <div key={field.id} style={Object.assign({}, customCellStyle, { textAlign: (customConfig.imageAlignment || 'left') as any })}>
                              {previewImageUrl
                                ? <img src={previewImageUrl} alt={String(customConfig.imageAltText || field.label || '')} style={{ maxWidth: '100%', width: customConfig.imageWidth ? String(customConfig.imageWidth) + 'px' : 'auto', height: customConfig.imageHeight ? String(customConfig.imageHeight) + 'px' : 'auto', objectFit: (customConfig.imageFit || 'contain') as any }} />
                                : <div style={{ border: '1px dashed #c8c6c4', padding: '16px', color: '#605e5c', textAlign: 'center' }}>Select an image from a SharePoint library</div>}
                            </div>
                          );
                        }
                        var previewDividerSpacing = customConfig.dividerSpacing === undefined ? 16 : Math.max(0, customConfig.dividerSpacing);
                        var previewDividerOrientation = customConfig.dividerOrientation || 'horizontal';
                        var previewDividerBorder = String(Math.max(1, customConfig.dividerThickness || 1)) + 'px ' + String(customConfig.dividerStyle || 'solid') + ' ' + String(customConfig.dividerColor || '#c8c6c4');
                        if (previewDividerOrientation === 'vertical') {
                          return <div key={field.id} style={Object.assign({}, customCellStyle, { display: 'flex', justifyContent: 'center' })}><div role="separator" aria-orientation="vertical" style={{ borderLeft: previewDividerBorder, height: String(Math.max(1, customConfig.dividerLength || 80)) + 'px', margin: '0 ' + String(previewDividerSpacing) + 'px' }} /></div>;
                        }
                        return <div key={field.id} style={customCellStyle}><hr style={{ border: 0, borderTop: previewDividerBorder, margin: String(previewDividerSpacing) + 'px 0' }} /></div>;
                      }

                      var fieldCellStyle: React.CSSProperties = previewGrid
                        ? (field.startNewRow === true
                          ? { gridColumn: '1 / -1' }
                          : { gridColumn: 'span ' + String(Math.max(1, field.columnSpan || 1)) })
                        : { marginBottom: String(fieldVerticalSpacing) + 'px' };

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

                      // Label style: always resolved (own default when unset) so the preview label
                      // never inherits the container's font styling.
                      var labelStyle: React.CSSProperties = { fontWeight: 600, fontSize: 14, fontFamily: DEFAULT_THEME_FONT_FAMILY, color: '#000000' };
                      if (field.labelFontSize) { labelStyle.fontSize = field.labelFontSize; }
                      if (field.labelFontFamily) { labelStyle.fontFamily = field.labelFontFamily; }
                      if (field.labelFontWeight) { labelStyle.fontWeight = field.labelFontWeight as any; }
                      if (field.labelColor) { labelStyle.color = field.labelColor; }

                      // Input/value style: always resolved (own default when unset)
                      var inputStyle: React.CSSProperties = { fontSize: '12px', fontFamily: DEFAULT_THEME_FONT_FAMILY, fontWeight: 'normal', color: '#605e5c', marginTop: '4px' };
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
          <button type="button" className={styles.designerPaletteButton} onClick={this.handleAddCustomImage}>
            <span>Image</span>
            <span className={styles.designerPaletteMeta}>From SharePoint library</span>
          </button>
          <button type="button" className={styles.designerPaletteButton} onClick={this.handleAddDivider}>
            <span>Divider</span>
            <span className={styles.designerPaletteMeta}>Layout</span>
          </button>
          <button type="button" className={styles.designerPaletteButton} onClick={this.handleAddGridControl}>
            <span>Grid Control</span>
            <span className={styles.designerPaletteMeta}>Editable child records</span>
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
          <button type="button" className={styles.designerTab} onClick={() => this.setState({ activeDesignerTab: 'conditional', selectedFieldId: null })}>Conditional styling</button>
          <button type="button" className={styles.designerActionButton} onClick={() => this.addStep()}>{strings.DesignerAddStep}</button>
          <button
            type="button"
            className={styles.designerInlineButton}
            title={strings.DesignerMoveStepLeft}
            aria-label={strings.DesignerMoveStepLeft}
            onClick={() => this.moveStep(-1)}
            disabled={this.state.selectedStepIndex <= 0}
          >
            {strings.DesignerMoveStepLeftLabel}
          </button>
          <button
            type="button"
            className={styles.designerInlineButton}
            title={strings.DesignerMoveStepRight}
            aria-label={strings.DesignerMoveStepRight}
            onClick={() => this.moveStep(1)}
            disabled={this.state.selectedStepIndex >= this.props.schema.steps.length - 1}
          >
            {strings.DesignerMoveStepRightLabel}
          </button>
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

  private renderCustomPlacementEditor(field: FormField): JSX.Element {
    return (
      <div>
        <label className={styles.designerFormLabel}>{strings.PropertyPanelColumnSpan}</label>
        <input className={styles.designerInput} type="number" min={1} max={12} value={String(Math.max(1, field.columnSpan || 1))} onChange={(ev) => this.updateSelectedField(function(nextField) { var parsed = parseInt(ev.currentTarget.value, 10); nextField.columnSpan = isNaN(parsed) ? 1 : Math.max(1, parsed); return nextField; })} />
        <label className={styles.designerCheckboxRow}><input type="checkbox" checked={field.startNewRow === true} onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.startNewRow = ev.currentTarget.checked; return nextField; })} /><span>{strings.DesignerStartNewRow}</span></label>
      </div>
    );
  }

  private renderFieldEditor(field: FormField): JSX.Element {
    if (field.type === 'customimage') {
      var imageConfig = field.config || {};
      return (
        <div className={styles.designerPanel}><div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Image</div>
          <div className={styles.designerPanelHint}>Add an image from a SharePoint library. This is a visual element and is not saved to a list field.</div>
          <label className={styles.designerFormLabel}>Image library</label>
          <select className={styles.designerInput} value={this.state.selectedImageLibrary} onChange={(ev) => this.loadImageFiles(ev.currentTarget.value)}><option value="">Select a library</option>{this.state.imageLibraries.map(function(option) { return <option key={option.key} value={option.key}>{option.text}</option>; })}</select>
          <label className={styles.designerFormLabel}>Image file</label>
          <select className={styles.designerInput} value={imageConfig.imageUrl || ''} disabled={!this.state.selectedImageLibrary || this.state.imageFilesLoading} onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.config = Object.assign({}, nextField.config || {}, { imageUrl: ev.currentTarget.value || undefined }); return nextField; })}><option value="">{this.state.imageFilesLoading ? 'Loading images...' : 'Select an image'}</option>{this.state.imageFiles.map(function(option) { return <option key={option.key} value={option.key}>{option.text}</option>; })}</select>
          {this.state.imageBrowserError && <div className={styles.designerPanelHint}>{this.state.imageBrowserError}</div>}
          <label className={styles.designerFormLabel}>Image URL</label>
          <input className={styles.designerInput} type="text" value={imageConfig.imageUrl || ''} placeholder="/sites/site/Library/logo.png" onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.config = Object.assign({}, nextField.config || {}, { imageUrl: ev.currentTarget.value || undefined }); return nextField; })} />
          <label className={styles.designerFormLabel}>Alternative text</label>
          <input className={styles.designerInput} type="text" value={imageConfig.imageAltText || ''} onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.config = Object.assign({}, nextField.config || {}, { imageAltText: ev.currentTarget.value || undefined }); return nextField; })} />
          <label className={styles.designerFormLabel}>Width (px, blank for natural size)</label>
          <input className={styles.designerInput} type="number" min={1} value={imageConfig.imageWidth === undefined ? '' : String(imageConfig.imageWidth)} onChange={(ev) => this.updateSelectedField(function(nextField) { var parsed = parseInt(ev.currentTarget.value, 10); nextField.config = Object.assign({}, nextField.config || {}, { imageWidth: isNaN(parsed) ? undefined : Math.max(1, parsed) }); return nextField; })} />
          <label className={styles.designerFormLabel}>Height (px, blank for natural size)</label>
          <input className={styles.designerInput} type="number" min={1} value={imageConfig.imageHeight === undefined ? '' : String(imageConfig.imageHeight)} onChange={(ev) => this.updateSelectedField(function(nextField) { var parsed = parseInt(ev.currentTarget.value, 10); nextField.config = Object.assign({}, nextField.config || {}, { imageHeight: isNaN(parsed) ? undefined : Math.max(1, parsed) }); return nextField; })} />
          <label className={styles.designerFormLabel}>Fit</label>
          <select className={styles.designerInput} value={imageConfig.imageFit || 'contain'} onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.config = Object.assign({}, nextField.config || {}, { imageFit: ev.currentTarget.value as 'contain' | 'cover' }); return nextField; })}><option value="contain">Contain</option><option value="cover">Cover</option></select>
          <label className={styles.designerFormLabel}>Alignment</label>
          <select className={styles.designerInput} value={imageConfig.imageAlignment || 'left'} onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.config = Object.assign({}, nextField.config || {}, { imageAlignment: ev.currentTarget.value as 'left' | 'center' | 'right' }); return nextField; })}><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></select>
          {this.renderCustomPlacementEditor(field)}
        </div></div>
      );
    }

    if (field.type === 'divider') {
      var dividerConfig = field.config || {};
      return (
        <div className={styles.designerPanel}><div className={styles.designerPanelSection}>
          <div className={styles.designerPanelTitle}>Divider</div>
          <div className={styles.designerPanelHint}>Separate fields within this container. This is a visual element and is not saved to a list field.</div>
          <label className={styles.designerFormLabel}>Orientation</label><select className={styles.designerInput} value={dividerConfig.dividerOrientation || 'horizontal'} onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.config = Object.assign({}, nextField.config || {}, { dividerOrientation: ev.currentTarget.value as 'horizontal' | 'vertical' }); return nextField; })}><option value="horizontal">Horizontal</option><option value="vertical">Vertical</option></select>
          <label className={styles.designerFormLabel}>Color</label><input type="color" value={dividerConfig.dividerColor || '#c8c6c4'} onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.config = Object.assign({}, nextField.config || {}, { dividerColor: ev.currentTarget.value }); return nextField; })} />
          <label className={styles.designerFormLabel}>Thickness (px)</label><input className={styles.designerInput} type="number" min={1} max={20} value={String(dividerConfig.dividerThickness || 1)} onChange={(ev) => this.updateSelectedField(function(nextField) { var parsed = parseInt(ev.currentTarget.value, 10); nextField.config = Object.assign({}, nextField.config || {}, { dividerThickness: isNaN(parsed) ? 1 : Math.max(1, parsed) }); return nextField; })} />
          <label className={styles.designerFormLabel}>Line style</label><select className={styles.designerInput} value={dividerConfig.dividerStyle || 'solid'} onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.config = Object.assign({}, nextField.config || {}, { dividerStyle: ev.currentTarget.value as 'solid' | 'dashed' | 'dotted' }); return nextField; })}><option value="solid">Solid</option><option value="dashed">Dashed</option><option value="dotted">Dotted</option></select>
          {dividerConfig.dividerOrientation === 'vertical' && <div><label className={styles.designerFormLabel}>Length (px)</label><input className={styles.designerInput} type="number" min={1} max={1000} value={String(dividerConfig.dividerLength || 80)} onChange={(ev) => this.updateSelectedField(function(nextField) { var parsed = parseInt(ev.currentTarget.value, 10); nextField.config = Object.assign({}, nextField.config || {}, { dividerLength: isNaN(parsed) ? 80 : Math.max(1, parsed) }); return nextField; })} /></div>}
          <label className={styles.designerFormLabel}>{dividerConfig.dividerOrientation === 'vertical' ? 'Horizontal spacing (px)' : 'Vertical spacing (px)'}</label><input className={styles.designerInput} type="number" min={0} max={100} value={String(dividerConfig.dividerSpacing === undefined ? 16 : dividerConfig.dividerSpacing)} onChange={(ev) => this.updateSelectedField(function(nextField) { var parsed = parseInt(ev.currentTarget.value, 10); nextField.config = Object.assign({}, nextField.config || {}, { dividerSpacing: isNaN(parsed) ? 16 : Math.max(0, parsed) }); return nextField; })} />
          {this.renderCustomPlacementEditor(field)}
        </div></div>
      );
    }

    if (field.type === 'gridcontrol') {
      var gridConfig = field.config || {};
      var selectedGridSourceId = gridConfig.gridControlSourceId || '';
      var selectedGridSourceAvailable = this.state.gridControlSources.some(function(option) { return option.key === selectedGridSourceId; });
      var gridSourceOptions = this.state.gridControlSources.slice(0);
      if (selectedGridSourceId && !selectedGridSourceAvailable) {
        gridSourceOptions.push({ key: selectedGridSourceId, text: (gridConfig.gridControlSourceName || selectedGridSourceId) + ' (unavailable)' });
      }
      var parentFieldOptions = this.state.reportListFields.slice(0);
      var selectedParentField = gridConfig.gridControlFilterSourceField || '';
      var selectedGridTarget = gridConfig.gridControlFilterTargetField || '';
      var selectedGridOperator = gridConfig.gridControlFilterOperator || 'eq';
      return (
        <div className={styles.designerPanel}>
          <div className={styles.designerPanelSection}>
            <div className={styles.designerPanelTitle}>Grid Control</div>
            <div className={styles.designerPanelHint}>Embed a Grid Control for child records. The parent list value filters existing rows and becomes the default for the mapped column on new rows.</div>

            <label className={styles.designerFormLabel}>{strings.PropertyPanelFieldName}</label>
            <input className={styles.designerInput} type="text" value={field.label} onChange={(ev) => this.updateSelectedField(function(nextField) { nextField.label = ev.currentTarget.value; return nextField; })} />

            <label className={styles.designerFormLabel}>Grid Control source</label>
            <select
              className={styles.designerInput}
              value={selectedGridSourceId}
              onChange={(ev) => {
                var sourceId = ev.currentTarget.value;
                var selectedSource = gridSourceOptions.filter(function(option) { return option.key === sourceId; })[0];
                this.updateSelectedField(function(nextField) {
                  nextField.config = Object.assign({}, nextField.config || {}, {
                    gridControlSourceId: sourceId || undefined,
                    gridControlSourceName: selectedSource ? selectedSource.text : undefined,
                    gridControlSourceListName: selectedSource && selectedSource.listName ? selectedSource.listName : undefined,
                    gridControlFilterTargetField: undefined
                  });
                  return nextField;
                });
                if (selectedSource && selectedSource.listName) {
                  this.loadGridControlFields(selectedSource.listName);
                } else {
                  this.setState({ gridControlFields: [], gridControlFieldsListName: '', gridControlFieldsLoading: false, gridControlFieldsError: sourceId ? 'The selected Grid Control did not provide its configured list name.' : null });
                }
              }}
            >
              <option value="">Select a Grid Control</option>
              {gridSourceOptions.map(function(option) { return <option key={option.key} value={option.key}>{option.text}</option>; })}
            </select>
            <button type="button" className={styles.designerInlineButton} onClick={() => this.loadGridControlSources()}>Refresh Grid Controls</button>

            <label className={styles.designerFormLabel}>Parent list field</label>
            <select className={styles.designerInput} value={selectedParentField} onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.config = Object.assign({}, nextField.config || {}, { gridControlFilterSourceField: ev.currentTarget.value || undefined });
              return nextField;
            })}>
              <option value="">No parent-child mapping</option>
              {parentFieldOptions.map(function(option) { return <option key={option.key} value={option.key}>{option.text}</option>; })}
            </select>

            <label className={styles.designerFormLabel}>Grid Control child column</label>
            <select className={styles.designerInput} value={selectedGridTarget} disabled={!selectedGridSourceId || this.state.gridControlFieldsLoading} onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.config = Object.assign({}, nextField.config || {}, { gridControlFilterTargetField: ev.currentTarget.value || undefined });
              return nextField;
            })}>
              <option value="">{this.state.gridControlFieldsLoading ? 'Loading columns...' : 'Select a child column'}</option>
              {this.state.gridControlFields.map(function(option) { return <option key={option.key} value={option.key}>{option.text}</option>; })}
            </select>
            {this.state.gridControlFieldsError && <div className={styles.designerPanelHint}>{this.state.gridControlFieldsError}</div>}

            <label className={styles.designerFormLabel}>Filter operator</label>
            <select className={styles.designerInput} value={selectedGridOperator} onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.config = Object.assign({}, nextField.config || {}, { gridControlFilterOperator: ev.currentTarget.value || 'eq' });
              return nextField;
            })}>
              <option value="eq">Equals</option>
              <option value="ne">Does not equal</option>
              <option value="contains">Contains</option>
              <option value="notcontains">Does not contain</option>
              <option value="startswith">Starts with</option>
              <option value="endswith">Ends with</option>
              <option value="gt">Greater than</option>
              <option value="ge">Greater than or equal</option>
              <option value="lt">Less than</option>
              <option value="le">Less than or equal</option>
            </select>

            {this.renderCustomPlacementEditor(field)}
          </div>

          <div className={styles.designerPanelSection}>
            <div className={styles.designerPanelTitle}>Title Font Settings</div>

            <label className={styles.designerFormLabel}>Font Size (px)</label>
            <input className={styles.designerInput} type="number" min={8} max={72} value={String(field.labelFontSize || 14)} onChange={(ev) => this.updateSelectedField(function(nextField) {
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextField.labelFontSize = isNaN(parsed) ? undefined : parsed;
              return nextField;
            })} />

            <label className={styles.designerFormLabel}>Font Family</label>
            <select className={styles.designerInput} value={field.labelFontFamily || 'inherit'} onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.labelFontFamily = ev.currentTarget.value === 'inherit' ? undefined : ev.currentTarget.value;
              return nextField;
            })}>
              <option value="inherit">Inherit</option>
              <option value="Arial">Arial</option>
              <option value="'Segoe UI'">Segoe UI</option>
              <option value="Verdana">Verdana</option>
              <option value="Georgia">Georgia</option>
              <option value="Courier New">Courier New</option>
              <option value="Times New Roman">Times New Roman</option>
            </select>

            <label className={styles.designerFormLabel}>Font Weight</label>
            <select className={styles.designerInput} value={field.labelFontWeight || 'normal'} onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.labelFontWeight = ev.currentTarget.value as 'normal' | 'bold' | '500' | '600' | '700' || undefined;
              return nextField;
            })}>
              <option value="normal">Normal</option>
              <option value="500">Medium (500)</option>
              <option value="600">Semi-Bold (600)</option>
              <option value="bold">Bold</option>
              <option value="700">Bold (700)</option>
            </select>

            <label className={styles.designerFormLabel}>Title Color</label>
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <input type="color" value={field.labelColor || '#000000'} onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.labelColor = ev.currentTarget.value || undefined;
                return nextField;
              })} style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }} />
              <input className={styles.designerInput} type="text" value={field.labelColor || ''} placeholder="#000000" onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.labelColor = ev.currentTarget.value || undefined;
                return nextField;
              })} style={{ flex: 1 }} />
            </div>
          </div>

          <div className={styles.designerPanelSection}>
            <div className={styles.designerPanelTitle}>Field Background &amp; Border</div>

            <label className={styles.designerFormLabel}>Background Color</label>
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <input type="color" value={field.fieldBackgroundColor || '#ffffff'} onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.fieldBackgroundColor = ev.currentTarget.value || undefined;
                return nextField;
              })} style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }} />
              <input className={styles.designerInput} type="text" value={field.fieldBackgroundColor || ''} placeholder="#ffffff" onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.fieldBackgroundColor = ev.currentTarget.value || undefined;
                return nextField;
              })} style={{ flex: 1 }} />
            </div>

            <label className={styles.designerFormLabel}>Border Style</label>
            <select className={styles.designerInput} value={field.fieldBorderStyle || 'square'} onChange={(ev) => this.updateSelectedField(function(nextField) {
              nextField.fieldBorderStyle = ev.currentTarget.value as 'square' | 'rounded' || undefined;
              return nextField;
            })}>
              <option value="square">Square Corners</option>
              <option value="rounded">Rounded Corners</option>
            </select>

            {(field.fieldBorderStyle || 'square') === 'rounded' && (
              <div>
                <label className={styles.designerFormLabel}>Corner Radius (px)</label>
                <input className={styles.designerInput} type="number" min={0} max={40} value={String(field.fieldBorderRadius || 8)} onChange={(ev) => this.updateSelectedField(function(nextField) {
                  var parsedRadius = parseInt(ev.currentTarget.value, 10);
                  nextField.fieldBorderRadius = isNaN(parsedRadius) ? 8 : parsedRadius;
                  return nextField;
                })} />
              </div>
            )}

            <label className={styles.designerFormLabel}>Border Color</label>
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <input type="color" value={field.fieldBorderColor || '#cccccc'} onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.fieldBorderColor = ev.currentTarget.value || undefined;
                return nextField;
              })} style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }} />
              <input className={styles.designerInput} type="text" value={field.fieldBorderColor || ''} placeholder="#cccccc" onChange={(ev) => this.updateSelectedField(function(nextField) {
                nextField.fieldBorderColor = ev.currentTarget.value || undefined;
                return nextField;
              })} style={{ flex: 1 }} />
            </div>

            <label className={styles.designerFormLabel}>Border Width (px)</label>
            <input className={styles.designerInput} type="number" min={0} max={10} value={String(field.fieldBorderWidth || 1)} onChange={(ev) => this.updateSelectedField(function(nextField) {
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextField.fieldBorderWidth = isNaN(parsed) ? 1 : parsed;
              return nextField;
            })} />
          </div>
        </div>
      );
    }

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
          <div className={styles.designerPanelTitle}>Container Description Font Settings</div>

          <label className={styles.designerFormLabel}>Font Size (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={8}
            max={72}
            value={String(theme.descriptionFontSize || 14)}
            onChange={(ev) => this.updateCurrentStep(function(nextStep) {
              if (!nextStep.theme) nextStep.theme = {};
              var parsed = parseInt(ev.currentTarget.value, 10);
              nextStep.theme.descriptionFontSize = isNaN(parsed) ? undefined : parsed;
              return nextStep;
            })}
          />

          <label className={styles.designerFormLabel}>Font Family</label>
          <select
            className={styles.designerInput}
            value={theme.descriptionFontFamily || 'inherit'}
            onChange={(ev) => this.updateCurrentStep(function(nextStep) {
              if (!nextStep.theme) nextStep.theme = {};
              nextStep.theme.descriptionFontFamily = ev.currentTarget.value === 'inherit' ? undefined : ev.currentTarget.value;
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
            value={theme.descriptionFontWeight || 'normal'}
            onChange={(ev) => this.updateCurrentStep(function(nextStep) {
              if (!nextStep.theme) nextStep.theme = {};
              nextStep.theme.descriptionFontWeight = ev.currentTarget.value as 'normal' | 'bold' | '500' | '600' | '700' || undefined;
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
              value={theme.descriptionColor || '#666666'}
              onChange={(ev) => this.updateCurrentStep(function(nextStep) {
                if (!nextStep.theme) nextStep.theme = {};
                nextStep.theme.descriptionColor = ev.currentTarget.value || undefined;
                return nextStep;
              })}
              style={{ width: '50px', height: '40px', cursor: 'pointer', border: '1px solid #ccc' }}
            />
            <input
              className={styles.designerInput}
              type="text"
              value={theme.descriptionColor || ''}
              placeholder="#666666"
              onChange={(ev) => this.updateCurrentStep(function(nextStep) {
                if (!nextStep.theme) nextStep.theme = {};
                nextStep.theme.descriptionColor = ev.currentTarget.value || undefined;
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
          <button
            className={activeTab === 'conditional' ? styles.designerTabActive : styles.designerTab}
            onClick={() => this.setState({ activeDesignerTab: 'conditional' })}
          >
            Conditional styling
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

          <label className={styles.designerFormLabel}>Form name alignment</label>
          <select className={styles.designerInput} value={this.props.schema.nameAlignment || 'left'} onChange={(ev) => this.updateForm(function(schema) { schema.nameAlignment = ev.currentTarget.value as 'left' | 'center' | 'right'; return schema; })}>
            <option value="left">Left</option><option value="center">Center</option><option value="right">Right</option>
          </select>

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

          <label className={styles.designerFormLabel}>Form description alignment</label>
          <select className={styles.designerInput} value={this.props.schema.descriptionAlignment || 'left'} onChange={(ev) => this.updateForm(function(schema) { schema.descriptionAlignment = ev.currentTarget.value as 'left' | 'center' | 'right'; return schema; })}>
            <option value="left">Left</option><option value="center">Center</option><option value="right">Right</option>
          </select>

          <label className={styles.designerFormLabel}>Logo image library</label>
          <select className={styles.designerInput} value={this.state.selectedImageLibrary} onChange={(ev) => this.loadImageFiles(ev.currentTarget.value)}>
            <option value="">Select a library</option>{this.state.imageLibraries.map(function(option) { return <option key={option.key} value={option.key}>{option.text}</option>; })}
          </select>
          <label className={styles.designerFormLabel}>Logo image</label>
          <select className={styles.designerInput} value={this.props.schema.logoUrl || ''} disabled={!this.state.selectedImageLibrary || this.state.imageFilesLoading} onChange={(ev) => this.updateForm(function(schema) { schema.logoUrl = ev.currentTarget.value || undefined; return schema; })}>
            <option value="">{this.state.imageFilesLoading ? 'Loading images...' : 'Select an image'}</option>{this.state.imageFiles.map(function(option) { return <option key={option.key} value={option.key}>{option.text}</option>; })}
          </select>
          {this.state.imageBrowserError && <div className={styles.designerPanelHint}>{this.state.imageBrowserError}</div>}
          <label className={styles.designerFormLabel}>Logo URL</label>
          <input className={styles.designerInput} type="text" value={this.props.schema.logoUrl || ''} placeholder="/sites/site/Library/logo.png" onChange={(ev) => this.updateForm(function(schema) { schema.logoUrl = ev.currentTarget.value || undefined; return schema; })} />
          <label className={styles.designerFormLabel}>Logo alternative text</label>
          <input className={styles.designerInput} type="text" value={this.props.schema.logoAltText || ''} onChange={(ev) => this.updateForm(function(schema) { schema.logoAltText = ev.currentTarget.value || undefined; return schema; })} />
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

          <label className={styles.designerFormLabel}>Vertical space between fields (px)</label>
          <input
            className={styles.designerInput}
            type="number"
            min={0}
            max={100}
            value={String(this.props.schema.theme && this.props.schema.theme.fieldVerticalSpacing !== undefined ? this.props.schema.theme.fieldVerticalSpacing : 16)}
            onChange={(ev) => this.updateForm(function(schema) {
              if (!schema.theme) schema.theme = {};
              var parsed = parseInt(ev.currentTarget.value, 10);
              schema.theme.fieldVerticalSpacing = isNaN(parsed) ? 16 : Math.max(0, Math.min(100, parsed));
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

          <label className={styles.designerFormLabel}>Container name alignment</label>
          <select className={styles.designerInput} value={currentStep.titleAlignment || 'left'} onChange={(ev) => this.updateCurrentStep(function(step) { step.titleAlignment = ev.currentTarget.value as 'left' | 'center' | 'right'; return step; })}>
            <option value="left">Left</option><option value="center">Center</option><option value="right">Right</option>
          </select>

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

          <label className={styles.designerFormLabel}>Container description alignment</label>
          <select className={styles.designerInput} value={currentStep.descriptionAlignment || 'left'} onChange={(ev) => this.updateCurrentStep(function(step) { step.descriptionAlignment = ev.currentTarget.value as 'left' | 'center' | 'right'; return step; })}>
            <option value="left">Left</option><option value="center">Center</option><option value="right">Right</option>
          </select>

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
        {activeTab === 'conditional' && this.renderConditionalRules()}
      </div>
    );
  }

  public render(): JSX.Element {
    var selectedField = this.getSelectedField();
    var conditionalMode = this.state.activeDesignerTab === 'conditional';

    return (
      <div className={styles.designerWorkspace}>
        {this.renderPalette()}
        {conditionalMode ? this.renderConditionalRulesCanvas() : this.renderCanvas()}
        {conditionalMode ? this.renderConditionalRules() : (selectedField ? this.renderFieldEditor(selectedField) : this.renderStepEditor())}
      </div>
    );
  }
}