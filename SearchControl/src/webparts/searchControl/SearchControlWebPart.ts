import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import * as strings from 'SearchControlWebPartStrings';
import { applyOverrideCss, PropertyPaneOverrideCss } from '../shared/overrideCss';
import './SearchControl.css';

var packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface IDynamicDataPropertyDefinitionCompat {
  id: string;
  title: string;
}

interface IDynamicDataSourceCompat {
  id: string;
  metadata: { alias?: string; title?: string };
  getPropertyValue: (propertyId: string) => any;
}

export interface ISearchControlWebPartProps {
  title: string;
  targetInstanceId: string;
  defaultMode: string;
  showModeToggle: boolean;
  buttonDisplayMode?: string;
  simplePlaceholder: string;
  searchOnType: boolean;
  minimumCharacters: number;
  debounceMilliseconds: number;
  timeDisplayFormat: string;
  timeMinuteIncrement: number;
  surfaceBackgroundColor: string;
  surfaceTextColor: string;
  accentColor: string;
  borderColor: string;
  borderWidth: number;
  cornerRadius: number;
  fontFamily: string;
  fontSize: number;
  excludeFromTabs: boolean;
  overrideCssUrl?: string;
  enableDiagnostics: boolean;
}

interface ISearchFieldMetadata {
  internalName: string;
  title: string;
  typeAsString: string;
  dateOnly?: boolean;
  choices?: string[];
  lookupList?: string;
}

interface ISearchMetadata {
  instanceId: string;
  instanceName: string;
  controlType: string;
  listName: string;
  fields: ISearchFieldMetadata[];
}

interface ISearchCondition {
  field: string;
  operator: string;
  logical: string;
  value: any;
}

interface IOption {
  key: string;
  text: string;
}

interface ITarget {
  source: IDynamicDataSourceCompat;
  metadata: ISearchMetadata;
}

export default class SearchControlWebPart extends BaseClientSideWebPart<ISearchControlWebPartProps> {
  private _mode: string = 'simple';
  private _simpleField: string = '*';
  private _simpleValue: string = '';
  private _conditions: ISearchCondition[] = [];
  private _searchState: any = { targetInstanceId: '', filterJson: '', mode: 'simple', query: '', conditions: [] };
  private _targetSourceId: string = '';
  private _targetSourceChangedHandler: () => void;
  private _lookupOptions: { [fieldName: string]: IOption[] } = {};
  private _lookupLoading: { [fieldName: string]: boolean } = {};
  private _lookupErrors: { [fieldName: string]: string } = {};
  private _dateOnlyByTarget: { [targetId: string]: { [fieldName: string]: boolean } } = {};
  private _dateMetadataLoading: { [targetId: string]: boolean } = {};
  private _typeTimer: number = 0;
  private _targetDiscoveryTimer: number = 0;
  private _targetDiscoveryAttempts: number = 0;
  private _statusMessage: string = '';
  private _statusIsError: boolean = false;
  private _dynamicDataSourceManager: any;
  private _dynamicDataProvider: any;

  protected onInit(): Promise<void> {
    this._mode = this.properties.defaultMode === 'advanced' ? 'advanced' : 'simple';
    this._dynamicDataSourceManager = (this.context as any).dynamicDataSourceManager || (this.context as any)._dynamicDataSourceManager;
    this._dynamicDataProvider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
    if (!this._dynamicDataSourceManager) {
      return Promise.reject(new Error('SearchControlWebPart: SPFx Dynamic Data source manager is unavailable.'));
    }
    if (this._dynamicDataSourceManager.initializeSource) {
      this._dynamicDataSourceManager.initializeSource(this);
    } else if (this._dynamicDataSourceManager.registerSource) {
      this._dynamicDataSourceManager.registerSource(this);
    } else {
      return Promise.reject(new Error('SearchControlWebPart: SPFx Dynamic Data source registration is unavailable.'));
    }
    this.logDiagnostic('Initialized Dynamic Data source.');
    return Promise.resolve();
  }

  protected onDispose(): void {
    if (this._typeTimer) {
      window.clearTimeout(this._typeTimer);
    }
    if (this._targetDiscoveryTimer) {
      window.clearTimeout(this._targetDiscoveryTimer);
    }
    this.unsubscribeTarget();
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinitionCompat> {
    return [{ id: 'searchState', title: 'Search state' }];
  }

  public getPropertyValue(propertyId: string): any {
    if (propertyId === 'searchState') {
      return this._searchState;
    }
    throw new Error('SearchControlWebPart: unsupported Dynamic Data property "' + propertyId + '".');
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }
    console.log('[SearchControlWebPart] ' + message);
  }

  private getWebPartVersion(): string {
    var solution: any = packageSolutionConfig && packageSolutionConfig.solution;
    return solution ? String(solution.version || '') : String(this.context.manifest.version || 'Unknown');
  }

  private normalizeId(value: string): string {
    return String(value || '').replace(/[{}]/g, '').toLowerCase();
  }

  private getTargets(): ITarget[] {
    if (!this._dynamicDataProvider || !this._dynamicDataProvider.getAvailableSources) {
      return [];
    }
    var sources: IDynamicDataSourceCompat[] = this._dynamicDataProvider.getAvailableSources();
    var targets: ITarget[] = [];
    for (var i = 0; i < sources.length; i += 1) {
      var source = sources[i];
      if (source.metadata.alias !== 'ListControlWebPart' && source.metadata.alias !== 'GridControlWebPart') {
        continue;
      }
      try {
        var metadata = source.getPropertyValue('searchMetadata') as ISearchMetadata;
        if (metadata && metadata.instanceId) {
          targets.push({ source: source, metadata: metadata });
        }
      } catch (error) {
        this.logDiagnostic('Unable to read search metadata from ' + source.id + ': ' + this.errorMessage(error));
      }
    }
    return targets;
  }

  private getSelectedTarget(): ITarget | undefined {
    var targetId = this.normalizeId(this.properties.targetInstanceId);
    var targets = this.getTargets();
    for (var i = 0; i < targets.length; i += 1) {
      if (this.normalizeId(targets[i].metadata.instanceId) === targetId) {
        this.applyDateOnlyMetadata(targets[i]);
        this.ensureDateOnlyMetadata(targets[i]);
        this.subscribeTarget(targets[i].source);
        return targets[i];
      }
    }
    this.unsubscribeTarget();
    return undefined;
  }

  private applyDateOnlyMetadata(target: ITarget): void {
    var formats = this._dateOnlyByTarget[target.source.id];
    if (!formats) {
      return;
    }
    var fields = target.metadata.fields || [];
    for (var index = 0; index < fields.length; index += 1) {
      var internalName = String(fields[index].internalName || '').toLowerCase();
      if (Object.prototype.hasOwnProperty.call(formats, internalName)) {
        fields[index].dateOnly = formats[internalName];
      }
    }
  }

  private ensureDateOnlyMetadata(target: ITarget): void {
    var fields = target.metadata.fields || [];
    var dateFields = fields.filter(function(field: ISearchFieldMetadata): boolean {
      return String(field.typeAsString || '').toLowerCase().indexOf('datetime') >= 0;
    });
    var hasCompleteMetadata = dateFields.length > 0 && dateFields.every(function(field: ISearchFieldMetadata): boolean {
      return typeof field.dateOnly === 'boolean';
    });
    if (hasCompleteMetadata || this._dateOnlyByTarget[target.source.id] || this._dateMetadataLoading[target.source.id]
      || !target.metadata.listName) {
      return;
    }

    this._dateMetadataLoading[target.source.id] = true;
    var escapedListName = String(target.metadata.listName).replace(/'/g, "''");
    var url = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '')
      + "/_api/web/lists/getByTitle('" + escapedListName + "')/fields?$select=InternalName,TypeAsString,DisplayFormat&$filter=Hidden eq false";
    this.logDiagnostic('REST request: GET ' + url + ' (date field metadata)');
    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        this.logDiagnostic('REST response: GET ' + url + ' -> HTTP '
          + String(response.status) + ' ' + response.statusText);
        if (!response.ok) {
          throw new Error('SharePoint returned HTTP ' + String(response.status) + '.');
        }
        return response.json();
      })
      .then((data: any) => {
        var rows = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
        var formats: { [fieldName: string]: boolean } = {};
        rows.forEach(function(field: any): void {
          if (String(field.TypeAsString || '').toLowerCase().indexOf('datetime') >= 0) {
            formats[String(field.InternalName || '').toLowerCase()] = Number(field.DisplayFormat) === 0;
          }
        });
        this._dateOnlyByTarget[target.source.id] = formats;
        this._dateMetadataLoading[target.source.id] = false;
        this.render();
      })
      .catch((error: any) => {
        this._dateMetadataLoading[target.source.id] = false;
        console.error('[SearchControlWebPart] Unable to load date field metadata: ' + this.errorMessage(error));
      });
  }

  private subscribeTarget(source: IDynamicDataSourceCompat): void {
    if (this._targetSourceId === source.id) {
      return;
    }
    this.unsubscribeTarget();
    this._targetSourceId = source.id;
    this._targetSourceChangedHandler = () => {
      this.updateResultCount(source);
      this.render();
    };
    this._dynamicDataProvider.registerSourceChanged(source.id, this._targetSourceChangedHandler);
  }

  private updateResultCount(source: IDynamicDataSourceCompat): void {
    if (!this._searchState.filterJson) {
      return;
    }
    try {
      var count = Number(source.getPropertyValue('searchResultCount'));
      if (!isNaN(count) && count >= 0) {
        this._statusMessage = String(count) + (count === 1 ? ' item found.' : ' items found.');
        this._statusIsError = false;
      }
    } catch (error) {
      this.logDiagnostic('Connected control does not expose a search result count: ' + this.errorMessage(error));
    }
  }

  private unsubscribeTarget(): void {
    if (this._targetSourceId && this._targetSourceChangedHandler) {
      this._dynamicDataProvider.unregisterSourceChanged(this._targetSourceId, this._targetSourceChangedHandler);
    }
    this._targetSourceId = '';
    this._targetSourceChangedHandler = undefined;
  }

  public render(): void {
    this.domElement.setAttribute('data-spse-exclude-from-tabs', String(this.properties.excludeFromTabs === true));
    applyOverrideCss(this.properties.overrideCssUrl || '', this.context.instanceId);
    var target = this.getSelectedTarget();
    this.domElement.innerHTML = '';

    var root = document.createElement('div');
    root.className = 'sc-root';
    root.style.backgroundColor = this.properties.surfaceBackgroundColor || '#ffffff';
    root.style.color = this.properties.surfaceTextColor || '#201f1e';
    root.style.border = String(this.properties.borderWidth === undefined ? 1 : this.properties.borderWidth) + 'px solid ' + (this.properties.borderColor || '#d2d0ce');
    root.style.borderRadius = String(this.properties.cornerRadius === undefined ? 4 : this.properties.cornerRadius) + 'px';
    root.style.fontFamily = this.properties.fontFamily || 'Segoe UI';
    root.style.fontSize = String(this.properties.fontSize || 14) + 'px';
    this.domElement.appendChild(root);

    var header = document.createElement('div');
    header.className = 'sc-header';
    root.appendChild(header);

    if (this.properties.title) {
      var heading = document.createElement('h2');
      heading.className = 'sc-title';
      heading.textContent = this.properties.title;
      header.appendChild(heading);
    }

    if (!target) {
      this.scheduleTargetDiscovery();
      this.appendMessage(root, this.displayMode === DisplayMode.Edit
        ? 'Select an SPS List Control or Grid Control in the property pane. The target must be on this page.'
        : 'This Search Control is not connected to an available List Control or Grid Control.', true);
      return;
    }
    this._targetDiscoveryAttempts = 0;
    if (this._targetDiscoveryTimer) {
      window.clearTimeout(this._targetDiscoveryTimer);
      this._targetDiscoveryTimer = 0;
    }

    var fields = target.metadata.fields || [];
    if (fields.length === 0) {
      this.appendMessage(root, 'The connected control is still loading its searchable field metadata.', false);
      return;
    }

    if (this.properties.showModeToggle !== false) {
      var modeGroup = document.createElement('fieldset');
      modeGroup.className = 'sc-mode-group';
      var modeLegend = document.createElement('legend');
      modeLegend.textContent = 'Mode';
      modeGroup.appendChild(modeLegend);
      modeGroup.appendChild(this.createModeRadio('simple', 'Simple'));
      modeGroup.appendChild(this.createModeRadio('advanced', 'Advanced'));
      header.appendChild(modeGroup);
    }

    if (this._mode === 'advanced') {
      this.renderAdvanced(root, fields);
    } else {
      this.renderSimple(root, fields);
    }

    if (this._statusMessage) {
      this.appendMessage(root, this._statusMessage, this._statusIsError);
    }
  }

  private renderSimple(root: HTMLElement, fields: ISearchFieldMetadata[]): void {
    var toolbar = document.createElement('div');
    toolbar.className = 'sc-toolbar';
    var fieldOptions: IOption[] = [{ key: '*', text: 'All compatible fields' }];
    for (var i = 0; i < fields.length; i += 1) {
      fieldOptions.push({ key: fields[i].internalName, text: fields[i].title });
    }
    if (this._simpleField !== '*' && !this.findField(fields, this._simpleField)) {
      this._simpleField = '*';
    }
    var fieldSelect = this.createSelect(fieldOptions, this._simpleField, 'Search field');
    fieldSelect.onchange = () => {
      this._simpleField = fieldSelect.value;
      this._simpleValue = '';
      this.render();
    };
    toolbar.appendChild(this.wrapField('Field', fieldSelect));

    var field = this.findField(fields, this._simpleField);
    this.ensureRemoteOptions(field);
    var valueControl = field
      ? this.createValueControl(field, this._simpleValue, (value: any) => {
        this._simpleValue = value;
        this.scheduleTypeSearch();
      })
      : this.createTextInput(this._simpleValue, this.properties.simplePlaceholder || 'Search...', (value: string) => {
        this._simpleValue = value;
        this.scheduleTypeSearch();
      });
    toolbar.appendChild(this.wrapField('Search for', valueControl));
    root.appendChild(toolbar);

    var actions = this.createActions();
    actions.appendChild(this.createButton('Search', 'Search', true, () => this.applySimple(fields)));
    actions.appendChild(this.createButton('Clear', 'ClearFilter', false, () => this.clearSearch()));
    root.appendChild(actions);
  }

  private createModeRadio(value: string, labelText: string): HTMLLabelElement {
    var label = document.createElement('label');
    label.className = 'sc-mode-option';
    var radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'search-mode-' + this.context.instanceId;
    radio.value = value;
    radio.checked = this._mode === value;
    radio.onchange = () => {
      if (radio.checked) {
        this._mode = value;
        this.render();
      }
    };
    label.appendChild(radio);
    label.appendChild(document.createTextNode(labelText));
    return label;
  }

  private renderAdvanced(root: HTMLElement, fields: ISearchFieldMetadata[]): void {
    if (this._conditions.length === 0) {
      this._conditions.push({
        field: fields[0].internalName,
        operator: this.getDefaultOperator(fields[0]),
        logical: 'and',
        value: ''
      });
    }

    for (var i = 0; i < this._conditions.length; i += 1) {
      this.renderCondition(root, fields, i);
    }

    var actions = this.createActions();
    actions.appendChild(this.createButton('Add condition', 'Add', false, () => {
      this._conditions.push({
        field: fields[0].internalName,
        operator: this.getDefaultOperator(fields[0]),
        logical: 'and',
        value: ''
      });
      this.render();
    }));
    actions.appendChild(this.createButton('Apply filters', 'CheckMark', true, () => this.applyAdvanced(fields)));
    actions.appendChild(this.createButton('Clear', 'ClearFilter', false, () => this.clearSearch()));
    root.appendChild(actions);
  }

  private renderCondition(root: HTMLElement, fields: ISearchFieldMetadata[], index: number): void {
    var condition = this._conditions[index];
    var field = this.findField(fields, condition.field) || fields[0];
    condition.field = field.internalName;
    var row = document.createElement('div');
    row.className = 'sc-condition';

    if (index > 0) {
      var logical = this.createSelect([
        { key: 'and', text: 'AND' },
        { key: 'or', text: 'OR' }
      ], condition.logical, 'Connector');
      logical.onchange = () => { condition.logical = logical.value; };
      row.appendChild(this.wrapField('Connector', logical, 'sc-field-connector'));
    }

    var fieldOptions: IOption[] = fields.map(function(item: ISearchFieldMetadata): IOption {
      return { key: item.internalName, text: item.title };
    });
    var fieldSelect = this.createSelect(fieldOptions, field.internalName, 'Field');
    fieldSelect.onchange = () => {
      var selected = this.findField(fields, fieldSelect.value);
      condition.field = fieldSelect.value;
      condition.operator = this.getDefaultOperator(selected);
      condition.value = '';
      this.ensureRemoteOptions(selected);
      this.render();
    };
    row.appendChild(this.wrapField('Field', fieldSelect));

    var operatorSelect = this.createSelect(this.getOperatorOptions(field), condition.operator, 'Operator');
    operatorSelect.onchange = () => { condition.operator = operatorSelect.value; };
    row.appendChild(this.wrapField('Operator', operatorSelect, 'sc-field-operator'));

    this.ensureRemoteOptions(field);
    row.appendChild(this.wrapField('Value', this.createValueControl(field, condition.value, (value: any) => {
      condition.value = value;
    })));

    row.appendChild(this.createButton('Remove', 'Delete', false, () => {
      this._conditions.splice(index, 1);
      this.render();
    }));
    root.appendChild(row);
  }

  private applySimple(fields: ISearchFieldMetadata[]): void {
    var value = String(this._simpleValue === undefined || this._simpleValue === null ? '' : this._simpleValue).trim();
    if (!value) {
      this.clearSearch();
      return;
    }
    var conditions: any[] = [];
    if (this._simpleField === '*') {
      var compatibleFields = fields.filter((field: ISearchFieldMetadata) => this.isSimpleCompatible(field));
      for (var i = 0; i < compatibleFields.length; i += 1) {
        conditions.push({
          field: compatibleFields[i].internalName,
          operator: 'contains',
          logical: i === 0 ? 'and' : 'or',
          valueType: 'static',
          value: value
        });
      }
      if (conditions.length === 0) {
        this.setStatus('The connected control does not expose any fields compatible with all-field search.', true);
        return;
      }
    } else {
      var field = this.findField(fields, this._simpleField);
      if (!field) {
        this.setStatus('The selected search field is no longer available.', true);
        return;
      }
      conditions.push({
        field: field.internalName,
        operator: this.getDefaultOperator(field),
        logical: 'and',
        valueType: 'static',
        value: this._simpleValue
      });
    }
    this.publish(conditions, value);
  }

  private applyAdvanced(fields: ISearchFieldMetadata[]): void {
    var published: any[] = [];
    for (var i = 0; i < this._conditions.length; i += 1) {
      var condition = this._conditions[i];
      var field = this.findField(fields, condition.field);
      if (!field) {
        this.setStatus('One or more selected fields are no longer available.', true);
        return;
      }
      if (condition.value === undefined || condition.value === null || String(condition.value).trim() === '') {
        this.setStatus('Enter a value for every advanced condition.', true);
        return;
      }
      published.push({
        field: condition.field,
        operator: condition.operator,
        logical: i === 0 ? 'and' : condition.logical,
        valueType: 'static',
        value: condition.value
      });
    }
    this.publish(published, '');
  }

  private publish(conditions: any[], query: string): void {
    this._searchState = {
      targetInstanceId: this.properties.targetInstanceId,
      filterJson: JSON.stringify(conditions),
      mode: this._mode,
      query: query,
      conditions: conditions
    };
    this._statusMessage = 'Searching...';
    this._statusIsError = false;
    this.notifySearchStateChanged();
    this.logDiagnostic('Published ' + String(conditions.length) + ' filter condition(s) for target ' + this.properties.targetInstanceId + '.');
    this.render();
  }

  private clearSearch(): void {
    this._simpleValue = '';
    for (var i = 0; i < this._conditions.length; i += 1) {
      this._conditions[i].value = '';
    }
    this._searchState = {
      targetInstanceId: this.properties.targetInstanceId,
      filterJson: '',
      mode: this._mode,
      query: '',
      conditions: []
    };
    this._statusMessage = 'Search filters cleared.';
    this._statusIsError = false;
    this.notifySearchStateChanged();
    this.render();
  }

  private scheduleTypeSearch(): void {
    if (!this.properties.searchOnType || this._simpleField !== '*') {
      return;
    }
    if (this._typeTimer) {
      window.clearTimeout(this._typeTimer);
    }
    var minimum = this.properties.minimumCharacters === undefined ? 2 : this.properties.minimumCharacters;
    if (this._simpleValue.length > 0 && this._simpleValue.length < minimum) {
      return;
    }
    this._typeTimer = window.setTimeout(() => {
      var target = this.getSelectedTarget();
      if (target) {
        this.applySimple(target.metadata.fields || []);
      }
    }, this.properties.debounceMilliseconds === undefined ? 300 : this.properties.debounceMilliseconds);
  }

  private scheduleTargetDiscovery(): void {
    if (!this.properties.targetInstanceId || this._targetDiscoveryTimer || this._targetDiscoveryAttempts >= 20) {
      return;
    }
    this._targetDiscoveryAttempts += 1;
    this._targetDiscoveryTimer = window.setTimeout(() => {
      this._targetDiscoveryTimer = 0;
      this.render();
    }, 500);
  }

  private ensureRemoteOptions(field: ISearchFieldMetadata): void {
    if (!field || !this.isRemoteField(field) || this._lookupOptions[field.internalName] || this._lookupLoading[field.internalName]) {
      return;
    }
    this._lookupLoading[field.internalName] = true;
    delete this._lookupErrors[field.internalName];
    var isUser = this.isUserField(field);
    var url = isUser
      ? this.context.pageContext.web.absoluteUrl + '/_api/web/siteusers?$select=Id,Title,Email&$top=5000'
      : this.context.pageContext.web.absoluteUrl + '/_api/web/lists(guid%27' + encodeURIComponent(String(field.lookupList || '').replace(/[{}]/g, '')) + '%27)/items?$select=Id,Title&$top=5000';
    this.logDiagnostic('REST request: GET ' + url + ' (field=' + field.internalName + ')');
    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        this.logDiagnostic('REST response: GET ' + url + ' -> HTTP '
          + String(response.status) + ' ' + response.statusText);
        if (!response.ok) {
          throw new Error('SharePoint returned HTTP ' + String(response.status) + '.');
        }
        return response.json();
      })
      .then((data: any) => {
        var rows = data && data.value ? data.value : [];
        this.logDiagnostic('REST payload summary: GET ' + url + ' -> items=' + String(rows.length));
        this._lookupOptions[field.internalName] = rows.map(function(row: any): IOption {
          var text = String(row.Title || row.Email || row.Id);
          if (row.Email && row.Title) {
            text += ' (' + String(row.Email) + ')';
          }
          return { key: String(row.Id), text: text };
        });
        this._lookupLoading[field.internalName] = false;
        this.render();
      })
      .catch((error: any) => {
        this._lookupLoading[field.internalName] = false;
        this._lookupErrors[field.internalName] = this.errorMessage(error);
        this.logDiagnostic('Unable to load options for ' + field.internalName + ': ' + this._lookupErrors[field.internalName]);
        this.render();
      });
  }

  private createValueControl(field: ISearchFieldMetadata, value: any, changed: (value: any) => void): HTMLElement {
    var type = this.fieldType(field);
    if (this._lookupLoading[field.internalName]) {
      var loading = document.createElement('span');
      loading.textContent = 'Loading options...';
      return loading;
    }
    if (this._lookupErrors[field.internalName]) {
      var error = document.createElement('button');
      error.type = 'button';
      error.className = 'sc-button';
      error.textContent = 'Retry loading options';
      error.title = this._lookupErrors[field.internalName];
      error.onclick = () => {
        delete this._lookupErrors[field.internalName];
        this.ensureRemoteOptions(field);
        this.render();
      };
      return error;
    }
    if (type === 'boolean') {
      var booleanSelect = this.createSelect([
        { key: '', text: 'Select...' },
        { key: 'Yes', text: 'Yes' },
        { key: 'No', text: 'No' }
      ], String(value || ''), field.title);
      booleanSelect.onchange = () => changed(booleanSelect.value);
      return booleanSelect;
    }
    if (type === 'choice' || this.isRemoteField(field)) {
      var options = type === 'choice'
        ? (field.choices || []).map(function(choice: string): IOption { return { key: choice, text: choice }; })
        : (this._lookupOptions[field.internalName] || []);
      options.unshift({ key: '', text: 'Select...' });
      var optionSelect = this.createSelect(options, String(value || ''), field.title);
      optionSelect.onchange = () => changed(optionSelect.value);
      return optionSelect;
    }
    if (type === 'datetime' && this.properties.timeDisplayFormat !== '12hour') {
      return this.create24HourDateTimeControl(String(value || ''), field.title, changed);
    }
    var input = document.createElement('input');
    input.value = value === undefined || value === null ? '' : String(value);
    input.setAttribute('aria-label', field.title);
    input.type = type === 'number' ? 'number' : (type === 'date' ? 'date' : (type === 'datetime' ? 'datetime-local' : 'text'));
    if (type === 'datetime') {
      input.lang = this.properties.timeDisplayFormat === '12hour' ? 'en-US' : 'en-GB';
      input.step = String((this.properties.timeMinuteIncrement || 5) * 60);
    }
    input.oninput = () => changed(input.value);
    return input;
  }

  private create24HourDateTimeControl(value: string, label: string, changed: (value: any) => void): HTMLElement {
    var text = String(value || '');
    var separatorIndex = text.indexOf('T');
    var dateValue = text.substring(0, 10);
    var timeValue = separatorIndex >= 0 ? text.substring(separatorIndex + 1) : '';
    var hourValue = /^\d{2}:\d{2}/.test(timeValue) ? timeValue.substring(0, 2) : '';
    var minuteValue = /^\d{2}:\d{2}/.test(timeValue) ? timeValue.substring(3, 5) : '';
    var increment = [1, 5, 10, 15].indexOf(Number(this.properties.timeMinuteIncrement)) >= 0
      ? Number(this.properties.timeMinuteIncrement) : 5;
    var container = document.createElement('div');
    container.className = 'sc-24hour-input';
    var dateInput = document.createElement('input');
    dateInput.type = 'date';
    dateInput.value = dateValue;
    dateInput.setAttribute('aria-label', label + ' date');
    var hourOptions: IOption[] = [{ key: '', text: 'HH' }];
    var minuteOptions: IOption[] = [{ key: '', text: 'mm' }];
    var index: number;
    for (index = 0; index < 24; index += 1) {
      var hourText = index < 10 ? '0' + String(index) : String(index);
      hourOptions.push({ key: hourText, text: hourText });
    }
    for (index = 0; index < 60; index += increment) {
      var minuteText = index < 10 ? '0' + String(index) : String(index);
      minuteOptions.push({ key: minuteText, text: minuteText });
    }
    if (minuteValue && !minuteOptions.some(function(option: IOption): boolean { return option.key === minuteValue; })) {
      minuteOptions.push({ key: minuteValue, text: minuteValue });
    }
    var hourSelect = this.createSelect(hourOptions, hourValue, label + ' hour (24-hour)');
    var minuteSelect = this.createSelect(minuteOptions, minuteValue, label + ' minute');
    var emit = (): void => {
      changed(dateInput.value
        ? dateInput.value + 'T' + (hourSelect.value || '00') + ':' + (minuteSelect.value || '00')
        : '');
    };
    dateInput.oninput = emit;
    hourSelect.onchange = emit;
    minuteSelect.onchange = emit;
    var separator = document.createElement('span');
    separator.textContent = ':';
    separator.setAttribute('aria-hidden', 'true');
    container.appendChild(dateInput);
    container.appendChild(hourSelect);
    container.appendChild(separator);
    container.appendChild(minuteSelect);
    return container;
  }

  private createTextInput(value: string, placeholder: string, changed: (value: string) => void): HTMLInputElement {
    var input = document.createElement('input');
    input.type = 'search';
    input.value = value || '';
    input.placeholder = placeholder;
    input.setAttribute('aria-label', 'Search value');
    input.oninput = () => changed(input.value);
    input.onkeydown = (event: KeyboardEvent) => {
      if (event.keyCode === 13) {
        var target = this.getSelectedTarget();
        if (target) {
          this.applySimple(target.metadata.fields || []);
        }
      }
    };
    return input;
  }

  private createSelect(options: IOption[], selected: string, label: string): HTMLSelectElement {
    var select = document.createElement('select');
    select.setAttribute('aria-label', label);
    for (var i = 0; i < options.length; i += 1) {
      var option = document.createElement('option');
      option.value = options[i].key;
      option.textContent = options[i].text;
      option.selected = options[i].key === selected;
      select.appendChild(option);
    }
    return select;
  }

  private createButton(text: string, iconName: string, primary: boolean, action: () => void): HTMLButtonElement {
    var button = document.createElement('button');
    var displayMode = this.properties.buttonDisplayMode === 'text' || this.properties.buttonDisplayMode === 'iconText'
      ? this.properties.buttonDisplayMode : 'icon';
    button.type = 'button';
    button.className = (primary ? 'sc-button sc-button-primary' : 'sc-button')
      + (displayMode === 'icon' ? ' sc-button-icon-only' : '');
    button.title = text;
    button.setAttribute('aria-label', text);
    if (displayMode === 'text') {
      button.textContent = text;
    } else {
      var content = document.createElement('span');
      content.className = 'sc-command-content';
      var icon = document.createElement('i');
      icon.className = 'ms-Icon ms-Icon--' + iconName;
      icon.setAttribute('aria-hidden', 'true');
      var label = document.createElement('span');
      label.className = displayMode === 'icon' ? 'sc-visually-hidden' : '';
      label.textContent = text;
      content.appendChild(icon);
      content.appendChild(label);
      button.appendChild(content);
    }
    if (primary) {
      button.style.backgroundColor = this.properties.accentColor || '#0078d4';
      button.style.borderColor = this.properties.accentColor || '#0078d4';
    }
    button.onclick = action;
    return button;
  }

  private createActions(): HTMLElement {
    var actions = document.createElement('div');
    actions.className = 'sc-actions';
    return actions;
  }

  private wrapField(labelText: string, control: HTMLElement, className?: string): HTMLElement {
    var wrapper = document.createElement('div');
    wrapper.className = 'sc-field' + (className ? ' ' + className : '');
    var label = document.createElement('label');
    label.textContent = labelText;
    wrapper.appendChild(label);
    wrapper.appendChild(control);
    return wrapper;
  }

  private appendMessage(parent: HTMLElement, message: string, isError: boolean): void {
    var element = document.createElement('div');
    element.className = isError ? 'sc-message sc-error' : 'sc-message';
    element.setAttribute('role', isError ? 'alert' : 'status');
    element.textContent = message;
    parent.appendChild(element);
  }

  private setStatus(message: string, isError: boolean): void {
    this._statusMessage = message;
    this._statusIsError = isError;
    this.render();
  }

  private findField(fields: ISearchFieldMetadata[], name: string): ISearchFieldMetadata | undefined {
    var normalized = String(name || '').toLowerCase();
    for (var i = 0; i < fields.length; i += 1) {
      if (String(fields[i].internalName || '').toLowerCase() === normalized) {
        return fields[i];
      }
    }
    return undefined;
  }

  private fieldType(field: ISearchFieldMetadata): string {
    var type = String(field && field.typeAsString || '').toLowerCase();
    if (type.indexOf('datetime') >= 0) {
      return field && field.dateOnly === true ? 'date' : 'datetime';
    }
    if (type.indexOf('boolean') >= 0) { return 'boolean'; }
    if (type.indexOf('choice') >= 0) { return 'choice'; }
    if (type.indexOf('number') >= 0 || type.indexOf('currency') >= 0 || type.indexOf('integer') >= 0) { return 'number'; }
    return 'text';
  }

  private isUserField(field: ISearchFieldMetadata): boolean {
    return String(field && field.typeAsString || '').toLowerCase().indexOf('user') >= 0;
  }

  private isRemoteField(field: ISearchFieldMetadata): boolean {
    var type = String(field && field.typeAsString || '').toLowerCase();
    return type.indexOf('lookup') >= 0 || type.indexOf('user') >= 0;
  }

  private isSimpleCompatible(field: ISearchFieldMetadata): boolean {
    var type = String(field && field.typeAsString || '').toLowerCase();
    return type === 'text' || type === 'note' || type === 'url' || type.indexOf('choice') >= 0;
  }

  private getDefaultOperator(field: ISearchFieldMetadata): string {
    return this.fieldType(field) === 'text' && !this.isRemoteField(field) ? 'contains' : 'eq';
  }

  private getOperatorOptions(field: ISearchFieldMetadata): IOption[] {
    var type = this.fieldType(field);
    if (type === 'text' && !this.isRemoteField(field)) {
      return [
        { key: 'contains', text: 'Contains' },
        { key: 'notcontains', text: 'Does not contain' },
        { key: 'eq', text: 'Equals' },
        { key: 'ne', text: 'Does not equal' },
        { key: 'startswith', text: 'Starts with' },
        { key: 'endswith', text: 'Ends with' }
      ];
    }
    var options: IOption[] = [
      { key: 'eq', text: 'Equals' },
      { key: 'ne', text: 'Does not equal' }
    ];
    if (type === 'number' || type === 'date' || type === 'datetime') {
      options.push(
        { key: 'gt', text: 'Greater than' },
        { key: 'ge', text: 'Greater than or equal' },
        { key: 'lt', text: 'Less than' },
        { key: 'le', text: 'Less than or equal' }
      );
    }
    return options;
  }

  private errorMessage(error: any): string {
    return error && error.message ? String(error.message) : String(error || 'Unknown error');
  }

  private notifySearchStateChanged(): void {
    if (!this._dynamicDataSourceManager) {
      throw new Error('SearchControlWebPart: cannot publish because the Dynamic Data source manager is unavailable.');
    }
    var notified = false;
    if (this._dynamicDataSourceManager.notifyPropertyChanged) {
      this._dynamicDataSourceManager.notifyPropertyChanged('searchState');
      notified = true;
    }
    if (this._dynamicDataSourceManager.notifySourceChanged) {
      this._dynamicDataSourceManager.notifySourceChanged();
      notified = true;
    }
    if (!notified && this._dynamicDataSourceManager.notifyDataChanged) {
      this._dynamicDataSourceManager.notifyDataChanged();
      notified = true;
    }
    if (!notified) {
      throw new Error('SearchControlWebPart: Dynamic Data change notification is unavailable.');
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'targetInstanceId' && oldValue !== newValue) {
      if (oldValue && this._searchState.filterJson) {
        this._searchState = {
          targetInstanceId: oldValue,
          filterJson: '',
          mode: this._mode,
          query: '',
          conditions: []
        };
        this.notifySearchStateChanged();
      }
      this.unsubscribeTarget();
      this._lookupOptions = {};
      this._lookupErrors = {};
      this._conditions = [];
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private handleColorPropertyChange(propertyPath: string, oldValue: any, newValue: any): void {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.context.propertyPane.refresh();
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    var targets = this.getTargets();
    var targetOptions = targets.map(function(target: ITarget): IOption {
      var metadata = target.metadata;
      var typeName = metadata.controlType === 'grid' ? 'Grid' : 'List';
      return {
        key: metadata.instanceId,
        text: (metadata.instanceName || typeName) + ' - ' + typeName + (metadata.listName ? ' (' + metadata.listName + ')' : '')
      };
    });
    if (targetOptions.length === 0) {
      targetOptions.push({ key: '', text: 'No SPS List/Grid controls found' });
    }

    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: 'Connection',
          groupFields: [
            PropertyPaneDropdown('targetInstanceId', { label: 'Target control', options: targetOptions }),
            PropertyPaneDropdown('defaultMode', {
              label: 'Default mode',
              options: [{ key: 'simple', text: 'Simple' }, { key: 'advanced', text: 'Advanced' }]
            }),
            PropertyPaneCheckbox('showModeToggle', { text: 'Allow users to switch modes' }),
            PropertyPaneTextField('title', { label: 'Title' }),
            PropertyPaneTextField('simplePlaceholder', { label: 'Search placeholder' })
          ]
        }, {
          groupName: 'Buttons',
          groupFields: [
            PropertyPaneDropdown('buttonDisplayMode', {
              label: 'Button display',
              options: [
                { key: 'text', text: 'Names' },
                { key: 'iconText', text: 'Icons and names' },
                { key: 'icon', text: 'Icons' }
              ],
              selectedKey: this.properties.buttonDisplayMode === 'text' || this.properties.buttonDisplayMode === 'iconText'
                ? this.properties.buttonDisplayMode : 'icon'
            })
          ]
        }, {
          groupName: 'Simple search',
          groupFields: [
            PropertyPaneCheckbox('searchOnType', { text: 'Search while typing (all fields)' }),
            PropertyPaneSlider('minimumCharacters', { label: 'Minimum characters', min: 0, max: 10, step: 1 }),
            PropertyPaneSlider('debounceMilliseconds', { label: 'Typing delay (milliseconds)', min: 100, max: 1500, step: 100 })
          ]
        }, {
          groupName: 'Date and time input',
          groupFields: [
            PropertyPaneDropdown('timeDisplayFormat', {
              label: 'Time format',
              options: [{ key: '24hour', text: '24-hour' }, { key: '12hour', text: '12-hour' }]
            }),
            PropertyPaneDropdown('timeMinuteIncrement', {
              label: 'Minute increment',
              options: [{ key: 1, text: '1 minute' }, { key: 5, text: '5 minutes' }, { key: 10, text: '10 minutes' }, { key: 15, text: '15 minutes' }]
            })
          ]
        }, {
          groupName: 'Appearance',
          groupFields: [
            PropertyFieldColorPicker('surfaceBackgroundColor', {
              label: 'Background color',
              selectedColor: this.properties.surfaceBackgroundColor || '#ffffff',
              onPropertyChange: this.handleColorPropertyChange.bind(this),
              properties: this.properties,
              style: PropertyFieldColorPickerStyle.Inline,
              key: 'surfaceBackgroundColorField'
            }),
            PropertyFieldColorPicker('surfaceTextColor', {
              label: 'Text color',
              selectedColor: this.properties.surfaceTextColor || '#201f1e',
              onPropertyChange: this.handleColorPropertyChange.bind(this),
              properties: this.properties,
              style: PropertyFieldColorPickerStyle.Inline,
              key: 'surfaceTextColorField'
            }),
            PropertyFieldColorPicker('accentColor', {
              label: 'Accent color',
              selectedColor: this.properties.accentColor || '#0078d4',
              onPropertyChange: this.handleColorPropertyChange.bind(this),
              properties: this.properties,
              style: PropertyFieldColorPickerStyle.Inline,
              key: 'accentColorField'
            }),
            PropertyFieldColorPicker('borderColor', {
              label: 'Border color',
              selectedColor: this.properties.borderColor || '#d2d0ce',
              onPropertyChange: this.handleColorPropertyChange.bind(this),
              properties: this.properties,
              style: PropertyFieldColorPickerStyle.Inline,
              key: 'borderColorField'
            }),
            PropertyPaneSlider('borderWidth', { label: 'Border width', min: 0, max: 8, step: 1 }),
            PropertyPaneSlider('cornerRadius', { label: 'Corner radius', min: 0, max: 24, step: 1 }),
            PropertyPaneTextField('fontFamily', { label: 'Font family' }),
            PropertyPaneSlider('fontSize', { label: 'Font size', min: 10, max: 30, step: 1 })
          ]
        }, {
          groupName: 'Advanced',
          groupFields: [
            PropertyPaneOverrideCss('overrideCssUrl', this.properties.overrideCssUrl || '', this.context, (newValue: string): void => {
              var oldValue = this.properties.overrideCssUrl || '';
              this.properties.overrideCssUrl = newValue;
              this.onPropertyPaneFieldChanged('overrideCssUrl', oldValue, newValue);
              this.render();
            }),
            PropertyPaneCheckbox('excludeFromTabs', { text: 'Exclude this web part from SPS Tabs' }),
            PropertyPaneCheckbox('enableDiagnostics', { text: 'Enable diagnostic console logging' }),
            PropertyPaneLabel('versionInfo', { text: 'SPS Search Control version: ' + this.getWebPartVersion() })
          ]
        }]
      }]
    };
  }
}
