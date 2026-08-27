import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import * as strings from 'ListControlWebPartStrings';
import './ListControl.css';

export interface IListControlViewOption {
  key: string;
  text: string;
  isDefault?: boolean;
}

export interface IListControlColumnConfiguration {
  fieldName: string;
  displayName: string;
  width?: string;
}

export interface IListControlProps {
  context: any;
  listName: string;
  defaultViewId: string;
  views: IListControlViewOption[];
  viewColumns: IListControlColumnConfiguration[];
  pageSize: number;
  isEditMode: boolean;
  showViewSelector: boolean;
  showRefresh: boolean;
  showAdd: boolean;
  showEdit: boolean;
  showView: boolean;
  showDelete: boolean;
  showLinkToItem: boolean;
  linkTargetPageUrl: string;
  linkTargetIdParam: string;
  includeReturnUrlParam: boolean;
  enableDiagnostics: boolean;
  bodyTextColor: string;
  bodyFontFamily: string;
  bodyFontSize: string;
  bodyFontStyle: string;
  bodyFontBold: boolean;
  bodyTextAlign: string;
  dateDisplayFormat: string;
  timeDisplayFormat: string;
  selectedTextColor: string;
  selectedBackgroundColor: string;
  selectedFontStyle: string;
  selectedFontBold: boolean;
  headerTextColor: string;
  headerBackgroundColor: string;
  headerFontFamily: string;
  headerFontSize: string;
  headerFontStyle: string;
  headerFontBold: boolean;
  headerTextAlign: string;
  tableBackgroundColor: string;
  tableBorderColor: string;
  tableBorderWidth: string;
  tableCornerStyle: string;
  tableCornerRadius: number;
  tableRowLineWidth: number;
  cornerStyle: string;
  cornerRadius: number;
  alternateRowShading: boolean;
  alternateRowShadingColor: string;
  buttonTextColor: string;
  buttonBackgroundColor: string;
  buttonFontFamily: string;
  buttonFontSize: string;
  buttonFontStyle: string;
  buttonFontBold: boolean;
  buttonCornerStyle: string;
  buttonCornerRadius: number;
  webpartBackgroundColor: string;
  webpartBorderColor: string;
  webpartBorderWidth: number;
  filterJson?: string;
  conditionalStyleJson?: string;
  onSelectionChange: (itemId: number, mode: string) => void;
}

export interface IListFieldDefinition {
  Name: string;
  RealFieldName?: string;
  DisplayName?: string;
  TypeAsString?: string;
  DisplayFormat?: number;
  Hidden?: string | boolean;
  ConfiguredWidth?: string;
}

export interface IListControlState {
  selectedViewId: string;
  fields: IListFieldDefinition[];
  rows: any[];
  loading: boolean;
  error: string | null;
  selectedItemId: number;
  selectedMode: string;
  deleting: boolean;
  sortFieldName: string;
  sortDirection: 'asc' | 'desc' | '';
  activeFilterFieldName: string;
  filterPopoverStyle: any;
  activeFilterIsDate: boolean;
  draftFilterOperator: FilterOperator;
  draftFilterValue: string;
  draftFilterEndValue: string;
  datePickerTarget: 'start' | 'end' | '';
  datePickerMonth: string;
  columnFilters: { [fieldName: string]: IColumnFilter };
  currentPage: number;
  displayFormUrl: string;
  displayFormLoading: boolean;
  displayFormError: string;
  embeddedByReportForms: boolean;
  runtimeFilterJson: string;
  runtimeConfigOwner: string;
}

export type FilterOperator = 'eq' | 'ne' | 'contains' | 'notcontains' | 'startswith' | 'endswith' | 'gt' | 'ge' | 'lt' | 'le';

export interface IColumnFilter {
  operator: FilterOperator;
  value: any;
  endValue?: string;
  compareDateOnly?: boolean;
}

interface IPresetFilterCondition {
  field: string;
  operator?: string;
  logical?: string;
  valueType?: string;
  value: any;
}

type ConditionalScope = 'row' | 'column';

interface IConditionalStyleDefinition {
  backgroundColor?: string;
  color?: string;
  fontFamily?: string;
  fontSize?: string;
  fontStyle?: string;
  fontWeight?: string;
  textAlign?: string;
}

interface IConditionalStyleCondition {
  field: string;
  operator: FilterOperator;
  logical: 'and' | 'or';
  valueType?: string;
  value: string;
}

interface IConditionalStyleRule {
  key: string;
  scope: ConditionalScope;
  applyField: string;
  enabled: boolean;
  priority: number;
  sourceOrder: number;
  style: IConditionalStyleDefinition;
  conditions: IConditionalStyleCondition[];
}

interface IResolvedConditionalStyle {
  rowStyle: React.CSSProperties;
  columnStylesByFieldKey: { [fieldKey: string]: React.CSSProperties };
}

var LIST_CONTROL_REFRESH_EVENT = 'spse:listcontrol-refresh';
var LIST_CONTROL_RUNTIME_CONFIG_EVENT = 'spse:listcontrol-runtime-config';

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

function toPositiveInt(value: any): number {
  var parsed = parseInt(String(value || ''), 10);
  return !isNaN(parsed) && parsed > 0 ? parsed : 0;
}

function formatString(template: string, value: string | number): string {
  return template.replace('{0}', String(value));
}

function tryParseObject(value: any): any {
  if (!value || typeof value !== 'string') {
    return value;
  }

  try {
    return JSON.parse(value);
  } catch (_error) {
    return value;
  }
}

function toArray(value: any): any[] {
  var parsed = tryParseObject(value);

  if (!parsed) {
    return [];
  }

  if (Array.isArray(parsed)) {
    return parsed;
  }

  if (Array.isArray(parsed.results)) {
    return parsed.results;
  }

  return [];
}

function trimGuidBraces(value: string): string {
  return String(value || '').replace(/^[{]/, '').replace(/[}]$/, '');
}

function appendQueryParam(url: string, key: string, value: string): string {
  var hash = '';
  var base = url;
  var hashIndex = url.indexOf('#');
  if (hashIndex >= 0) {
    base = url.substring(0, hashIndex);
    hash = url.substring(hashIndex);
  }

  var separator = base.indexOf('?') >= 0 ? '&' : '?';
  return base + separator + encodeURIComponent(key) + '=' + encodeURIComponent(value) + hash;
}

function isWorkbenchUrl(url: string): boolean {
  return /\/_layouts\/15\/workbench\.aspx/i.test(String(url || ''));
}

function getQueryParam(url: string, paramName: string): string {
  var queryIndex = url.indexOf('?');
  if (queryIndex < 0) {
    return '';
  }

  var hashIndex = url.indexOf('#', queryIndex);
  var query = hashIndex >= 0 ? url.substring(queryIndex + 1, hashIndex) : url.substring(queryIndex + 1);
  var parts = query.split('&');
  for (var i = 0; i < parts.length; i += 1) {
    var pair = parts[i].split('=');
    if (pair.length < 1) {
      continue;
    }

    var key = decodeURIComponent(String(pair[0] || '')).toLowerCase();
    if (key === paramName.toLowerCase()) {
      return decodeURIComponent(String(pair[1] || ''));
    }
  }

  return '';
}

function stripQueryAndHash(url: string): string {
  var result = String(url || '');
  var hashIndex = result.indexOf('#');
  if (hashIndex >= 0) {
    result = result.substring(0, hashIndex);
  }
  var queryIndex = result.indexOf('?');
  if (queryIndex >= 0) {
    result = result.substring(0, queryIndex);
  }
  return result;
}

function resolveEmbeddedPageUrl(url: string): string {
  var source = String(url || '').trim();
  if (!source) {
    return '';
  }

  // First choice: explicit Source/source parameter.
  var fromSource = getQueryParam(source, 'source') || getQueryParam(source, 'Source');
  if (fromSource && /\.aspx/i.test(fromSource)) {
    return fromSource;
  }

  // Second choice: id parameter used by SharePoint wrapper pages.
  var fromId = getQueryParam(source, 'id') || getQueryParam(source, 'Id') || getQueryParam(source, 'ID');
  if (fromId && /\.aspx/i.test(fromId)) {
    return fromId;
  }

  return '';
}

function isSharePointWrapperUrl(url: string): boolean {
  return /\/_layouts\/15\/(sharepoint|onedrive|doc)\.aspx/i.test(String(url || ''));
}

function extractRenderRowsAndFields(data: any): { rows: any[]; fields: IListFieldDefinition[] } {
  var directData = data && data.d && data.d.RenderListDataAsStream
    ? data.d.RenderListDataAsStream : data;
  var responseData = tryParseObject(directData) || {};
  var schema = tryParseObject(responseData.ListSchema) || tryParseObject(responseData.Schema) || {};
  var listData = tryParseObject(responseData.ListData) || responseData;
  var rows = toArray(responseData.Row);

  if (rows.length === 0) {
    rows = toArray(responseData.Rows);
  }

  if (rows.length === 0) {
    rows = toArray(listData.Row);
  }

  if (rows.length === 0) {
    rows = toArray(listData.Rows);
  }

  if (rows.length === 0) {
    rows = toArray(data && data.d && data.d.Row);
  }

  var fields = toArray(schema.Field);
  if (fields.length === 0) {
    fields = toArray(responseData.Field);
  }

  return {
    rows: rows,
    fields: fields
  };
}

function isSystemItemKey(key: string): boolean {
  if (!key) {
    return true;
  }

  return key.indexOf('__') === 0 || key === 'odata.editLink' || key === 'odata.id' || key === 'odata.type' || key === 'GUID';
}

function normalizeViewFieldNames(value: any): string[] {
  var items = toArray(value);
  var names: string[] = [];

  for (var i = 0; i < items.length; i += 1) {
    var candidate = String(items[i] || '');
    if (!candidate) {
      continue;
    }

    if (names.indexOf(candidate) < 0) {
      names.push(candidate);
    }
  }

  return names;
}

function isExpandableFieldType(typeAsString: string): boolean {
  var normalized = String(typeAsString || '').trim();
  return normalized === 'User' || normalized === 'UserMulti' || normalized === 'Lookup' || normalized === 'LookupMulti';
}

function joinClassNames(classes: string[]): string {
  return classes.filter(function(className: string) {
    return !!className;
  }).join(' ');
}

function normalizeFilterOperator(value: any): FilterOperator {
  var normalized = String(value || '').trim().toLowerCase().replace(/\s+/g, '');
  if (normalized === 'eq' || normalized === 'ne' || normalized === 'contains' || normalized === 'notcontains' || normalized === 'startswith' || normalized === 'endswith' || normalized === 'gt' || normalized === 'ge' || normalized === 'lt' || normalized === 'le') {
    return normalized as FilterOperator;
  }

  return 'contains';
}

function normalizeFilterLogical(value: any): 'and' | 'or' {
  return String(value || '').trim().toLowerCase() === 'or' ? 'or' : 'and';
}

function formatLocalDate(value: Date): string {
  var month = String(value.getMonth() + 1);
  var day = String(value.getDate());
  return String(value.getFullYear()) + '-' + (month.length < 2 ? '0' + month : month) + '-' + (day.length < 2 ? '0' + day : day);
}

function toCssString(value: any): string {
  return String(value === undefined || value === null ? '' : value).trim();
}

function mergeStyleObjects(baseStyle: React.CSSProperties, overrideStyle: React.CSSProperties): React.CSSProperties {
  var merged: React.CSSProperties = {};
  var key: string;

  for (key in baseStyle) {
    if (Object.prototype.hasOwnProperty.call(baseStyle, key)) {
      (merged as any)[key] = (baseStyle as any)[key];
    }
  }

  for (key in overrideStyle) {
    if (Object.prototype.hasOwnProperty.call(overrideStyle, key)) {
      (merged as any)[key] = (overrideStyle as any)[key];
    }
  }

  return merged;
}

function toPriorityNumber(value: any, fallback: number): number {
  var parsed = parseInt(String(value === undefined || value === null ? '' : value), 10);
  if (isNaN(parsed)) {
    return fallback;
  }

  return parsed;
}

export class ListControl extends React.Component<IListControlProps, IListControlState> {
  private _refreshEventHandler: any;
  private _runtimeConfigEventHandler: any;
  private _fieldDisplayFormatMap: { [internalName: string]: number } = {};
  private _loadRowsRequestId: number = 0;

  public constructor(props: IListControlProps) {
    super(props);

    this.state = {
      selectedViewId: props.defaultViewId || this.getInitialViewId(props.views),
      fields: [],
      rows: [],
      loading: true,
      error: null,
      selectedItemId: 0,
      selectedMode: 'view',
      deleting: false,
      sortFieldName: '',
      sortDirection: '',
      activeFilterFieldName: '',
      filterPopoverStyle: {},
      activeFilterIsDate: false,
      draftFilterOperator: 'contains',
      draftFilterValue: '',
      draftFilterEndValue: '',
      datePickerTarget: '',
      datePickerMonth: '',
      columnFilters: {},
      currentPage: 0,
      displayFormUrl: '',
      displayFormLoading: false,
      displayFormError: '',
      embeddedByReportForms: false,
      runtimeFilterJson: '',
      runtimeConfigOwner: '',
    };

    this._refreshEventHandler = this.handleExternalRefresh.bind(this);
    this._runtimeConfigEventHandler = this.handleRuntimeConfig.bind(this);
  }

  public componentDidMount(): void {
    this.logDiagnostic('Component mounted. listName=' + String(this.props.listName || '(none)') + ', defaultViewId=' + String(this.props.defaultViewId || '(none)'));
    if (typeof window !== 'undefined' && window.addEventListener) {
      window.addEventListener(LIST_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
      window.addEventListener(LIST_CONTROL_RUNTIME_CONFIG_EVENT, this._runtimeConfigEventHandler);
    }
    this.loadRows();
  }

  public componentWillUnmount(): void {
    this._loadRowsRequestId += 1;
    if (typeof window !== 'undefined' && window.removeEventListener) {
      window.removeEventListener(LIST_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
      window.removeEventListener(LIST_CONTROL_RUNTIME_CONFIG_EVENT, this._runtimeConfigEventHandler);
    }
  }

  private async openDefaultDisplayForm(row: any): Promise<void> {
    var itemId = this.getRowItemId(row);
    if (itemId <= 0) {
      return;
    }
    this.setState({ displayFormLoading: true, displayFormError: '' });
    try {
      var metadataUrl = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName)
        + "')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder";
      var response = await this.getJsonWithFallback(metadataUrl);
      if (!response.ok) {
        throw new Error('HTTP ' + String(response.status) + ' ' + response.statusText);
      }
      var payload = await response.json();
      var list = payload && payload.d ? payload.d : payload;
      var rootFolderUrl = String(list && list.RootFolder && list.RootFolder.ServerRelativeUrl || '').replace(/\/$/, '');
      var webUrl = this.getWebUrl();
      var originMatch = webUrl.match(/^https?:\/\/[^/]+/i);
      var formUrl = String(originMatch ? originMatch[0] : '') + rootFolderUrl + '/DispForm.aspx';
      formUrl = appendQueryParam(formUrl, 'ID', String(itemId));
      formUrl = appendQueryParam(formUrl, 'IsDlg', '1');
      this.setState({ displayFormUrl: formUrl, displayFormLoading: false });
    } catch (error) {
      this.setState({
        displayFormLoading: false,
        displayFormError: String(error && error.message ? error.message : error)
      });
    }
  }

  private closeDefaultDisplayForm(): void {
    this.setState({ displayFormUrl: '', displayFormLoading: false, displayFormError: '' }, () => this.loadRows());
  }

  private setDisplayFormFrameRef = (frame: HTMLIFrameElement): void => {
    if (!frame) {
      return;
    }
    var dialogFrame: any = frame;
    dialogFrame.cancelPopUp = () => this.closeDefaultDisplayForm();
    dialogFrame.commitPopup = () => this.closeDefaultDisplayForm();
    dialogFrame.commonModalDialogClose = () => this.closeDefaultDisplayForm();
  }

  public componentDidUpdate(prevProps: IListControlProps, prevState: IListControlState): void {
    if (prevProps.listName !== this.props.listName || prevProps.defaultViewId !== this.props.defaultViewId) {
      this.logDiagnostic('Props changed; resetting selection and reloading rows. listName=' + String(this.props.listName || '(none)') + ', viewId=' + String(this.props.defaultViewId || '(none)'));
      var nextSelectedViewId = this.props.defaultViewId || this.getInitialViewId(this.props.views);
      var selectedViewWillChange = nextSelectedViewId !== this.state.selectedViewId;
      this.setState({
        selectedViewId: nextSelectedViewId,
        selectedItemId: 0,
        selectedMode: 'view',
        sortFieldName: '',
        sortDirection: '',
        activeFilterFieldName: '',
        draftFilterOperator: 'contains',
        draftFilterValue: '',
        columnFilters: {},
        currentPage: 0,
      }, () => {
        this.props.onSelectionChange(0, 'view');
        if (!selectedViewWillChange) {
          this.loadRows();
        }
      });
      return;
    }

    if (prevState.selectedViewId !== this.state.selectedViewId) {
      if (this.state.currentPage !== 0) {
        this.setState({ currentPage: 0 });
      }
      this.loadRows();
      return;
    }

    if (prevProps.pageSize !== this.props.pageSize || prevProps.filterJson !== this.props.filterJson) {
      this.setState({ currentPage: 0 });
    }
  }

  private getInitialViewId(views: IListControlViewOption[]): string {
    for (var i = 0; i < views.length; i += 1) {
      if (views[i].isDefault) {
        return String(views[i].key);
      }
    }
    return views.length > 0 ? String(views[0].key) : '';
  }

  private handleExternalRefresh(event: any): void {
    var detail = event && event.detail ? event.detail : {};
    var sourceListName = String(detail.listName || '');
    if (!sourceListName || !this.props.listName) {
      return;
    }

    if (sourceListName.toLowerCase() !== String(this.props.listName).toLowerCase()) {
      return;
    }

    this.logDiagnostic('Received external refresh event for list: ' + sourceListName);
    this.loadRows();
  }

  private handleRuntimeConfig(event: any): void {
    var detail = event && event.detail ? event.detail : {};
    var targetInstanceId = String(detail.instanceId || '').toLowerCase();
    var currentInstanceId = String(this.props.context && this.props.context.instanceId || '').toLowerCase();
    if (!targetInstanceId || targetInstanceId !== currentInstanceId) {
      return;
    }

    var owner = String(detail.owner || '');
    if (detail.active === false) {
      if (this.state.runtimeConfigOwner && owner && this.state.runtimeConfigOwner !== owner) {
        return;
      }
      this.setState({ embeddedByReportForms: false, runtimeFilterJson: '', runtimeConfigOwner: '' });
      return;
    }

    this.setState({
      embeddedByReportForms: true,
      runtimeFilterJson: String(detail.filterJson || ''),
      runtimeConfigOwner: owner
    });
  }

  private getWebUrl(): string {
    return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
  }

  private async getJsonWithFallback(url: string): Promise<any> {
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

  private async postJsonWithFallback(url: string, body: any): Promise<any> {
    var payload = JSON.stringify(body || {});
    var response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json; charset=utf-8'
      },
      body: payload
    });

    if (!response.ok) {
      response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        body: payload
      });
    }

    if (!response.ok) {
      response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=minimalmetadata',
          'Content-Type': 'application/json;odata=minimalmetadata'
        },
        body: payload
      });
    }

    if (!response.ok) {
      response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: payload
      });
    }

    return response;
  }

  private buildViewIdCandidates(selectedViewId: string): string[] {
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

  private buildMinimalViewXml(viewQuery: string, viewFieldNames: string[], rowLimit: number, scope: any): string {
    var queryText = String(viewQuery || '').trim();
    var queryDocument = new DOMParser().parseFromString(
      /^<Query(?:\s|>)/i.test(queryText) ? queryText : '<Query>' + queryText + '</Query>',
      'text/xml'
    );
    if (queryDocument.getElementsByTagName('parsererror').length > 0 || queryDocument.getElementsByTagName('Query').length === 0) {
      throw new Error('The selected SharePoint view query is invalid.');
    }

    var xmlDocument = new DOMParser().parseFromString('<View><Query/><ViewFields/></View>', 'text/xml');
    var viewElement = xmlDocument.getElementsByTagName('View')[0];
    var existingQuery = xmlDocument.getElementsByTagName('Query')[0];
    var queryElement = xmlDocument.importNode(queryDocument.getElementsByTagName('Query')[0], true);
    viewElement.replaceChild(queryElement, existingQuery);

    var viewFieldsElement = xmlDocument.getElementsByTagName('ViewFields')[0];
    for (var fieldIndex = 0; fieldIndex < viewFieldNames.length; fieldIndex += 1) {
      var fieldName = String(viewFieldNames[fieldIndex] || '');
      if (fieldName) {
        var fieldRef = xmlDocument.createElement('FieldRef');
        fieldRef.setAttribute('Name', fieldName);
        viewFieldsElement.appendChild(fieldRef);
      }
    }

    var normalizedScope = String(scope === undefined || scope === null ? '' : scope).toLowerCase();
    var scopeNames: { [key: string]: string } = {
      '1': 'Recursive',
      '2': 'RecursiveAll',
      '3': 'FilesOnly',
      'recursive': 'Recursive',
      'recursiveall': 'RecursiveAll',
      'filesonly': 'FilesOnly'
    };
    if (scopeNames[normalizedScope]) {
      viewElement.setAttribute('Scope', scopeNames[normalizedScope]);
    }

    if (rowLimit > 0) {
      var rowLimitElement = xmlDocument.createElement('RowLimit');
      rowLimitElement.setAttribute('Paged', 'TRUE');
      rowLimitElement.appendChild(xmlDocument.createTextNode(String(rowLimit)));
      viewElement.appendChild(rowLimitElement);
    }

    return new XMLSerializer().serializeToString(xmlDocument);
  }

  private async loadSelectedViewXml(selectedViewId: string, viewFieldNames: string[]): Promise<string> {
    if (!selectedViewId) {
      return '';
    }

    var webUrl = this.getWebUrl();
    var listPath = "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')";
    var viewIds = this.buildViewIdCandidates(selectedViewId);
    var urls: string[] = [];

    for (var i = 0; i < viewIds.length; i += 1) {
      var encoded = encodeURIComponent(viewIds[i]);
      var normalized = encodeURIComponent(trimGuidBraces(viewIds[i]));
      urls.push(webUrl + listPath + "/views/getById('" + encoded + "')?$select=ViewQuery,RowLimit,Scope");
      urls.push(webUrl + listPath + "/views(guid'" + normalized + "')?$select=ViewQuery,RowLimit,Scope");
    }

    for (var j = 0; j < urls.length; j += 1) {
      try {
        var response = await this.getJsonWithFallback(urls[j]);
        if (!response.ok) {
          continue;
        }

        var data = await response.json();
        var viewData = data && data.d ? data.d : data;
        if (viewData && viewData.ViewQuery !== undefined && viewData.ViewQuery !== null) {
          var rowLimit = parseInt(String(viewData.RowLimit || ''), 10);
          var viewXml = this.buildMinimalViewXml(String(viewData.ViewQuery), viewFieldNames, isNaN(rowLimit) ? 0 : rowLimit, viewData.Scope);
          this.logDiagnostic('Loaded minimal selected view CAML. HasFilter=' + String(/<Where(?:\s|>)/i.test(viewXml)) + ', hasSort=' + String(/<OrderBy(?:\s|>)/i.test(viewXml)) + ', fields=' + String(viewFieldNames.length) + '.');
          return viewXml;
        }
      } catch (viewXmlError) {
        this.logDiagnostic('loadSelectedViewXml: Attempt failed for url=' + urls[j] + ': ' + (viewXmlError && viewXmlError.message ? viewXmlError.message : String(viewXmlError)));
      }
    }

    throw new Error('Failed to load the selected SharePoint view definition.');
  }

  private async loadSelectedViewFieldNames(selectedViewId: string): Promise<string[]> {
    if (!selectedViewId) {
      return [];
    }

    var webUrl = this.getWebUrl();
    var listPath = "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')";
    var viewIds = this.buildViewIdCandidates(selectedViewId);
    var urls: string[] = [];

    for (var i = 0; i < viewIds.length; i += 1) {
      var encoded = encodeURIComponent(viewIds[i]);
      var normalized = encodeURIComponent(trimGuidBraces(viewIds[i]));
      urls.push(webUrl + listPath + "/views/getById('" + encoded + "')/ViewFields");
      urls.push(webUrl + listPath + "/views(guid'" + normalized + "')/ViewFields");
    }

    for (var j = 0; j < urls.length; j += 1) {
      try {
        var response = await this.getJsonWithFallback(urls[j]);

        if (!response.ok) {
          continue;
        }

        var data = await response.json();
        var names = normalizeViewFieldNames(data.value);
        if (names.length === 0) {
          names = normalizeViewFieldNames(data.Items);
        }
        if (names.length === 0) {
          names = normalizeViewFieldNames(data && data.d && data.d.Items);
        }
        if (names.length > 0) {
          return names;
        }
      } catch (viewFieldsError) {
        this.logDiagnostic('loadSelectedViewFieldNames: Attempt failed for url=' + urls[j] + ': ' + (viewFieldsError && viewFieldsError.message ? viewFieldsError.message : String(viewFieldsError)));
      }
    }

    this.logDiagnostic('loadSelectedViewFieldNames: No view field names resolved for viewId=' + String(selectedViewId));
    return [];
  }

  private getDisplayFields(): IListFieldDefinition[] {
    if (this.state.selectedViewId !== this.props.defaultViewId || !this.props.viewColumns || this.props.viewColumns.length === 0) {
      return this.state.fields;
    }

    var byName: { [fieldName: string]: IListFieldDefinition } = {};
    for (var i = 0; i < this.state.fields.length; i += 1) {
      var field = this.state.fields[i];
      var fieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
      var responseName = String(field.Name || '').toLowerCase();
      if (fieldName) { byName[fieldName] = field; }
      if (responseName) { byName[responseName] = field; }
    }

    var configured: IListFieldDefinition[] = [];
    for (var j = 0; j < this.props.viewColumns.length; j += 1) {
      var column = this.props.viewColumns[j];
      var match = byName[String(column.fieldName || '').toLowerCase()];
      if (match) {
        configured.push(Object.assign({}, match, {
          DisplayName: column.displayName || match.DisplayName || match.Name,
          ConfiguredWidth: column.width || ''
        }));
      }
    }
    return configured;
  }

  private getConfiguredColumnStyle(field: IListFieldDefinition): React.CSSProperties {
    var width = parseInt(String(field.ConfiguredWidth || ''), 10);
    if (!isNaN(width) && width > 0) {
      return { width: width + 'px', minWidth: width + 'px', maxWidth: width + 'px' };
    }
    return {};
  }

  private getFieldsForConsumption(rawFields: IListFieldDefinition[], viewFieldNames: string[]): IListFieldDefinition[] {
    if (viewFieldNames.length === 0) {
      return rawFields.filter(function(field: IListFieldDefinition) {
        return !(field.Hidden === true || field.Hidden === 'TRUE');
      });
    }

    var byName: { [name: string]: IListFieldDefinition } = {};
    for (var i = 0; i < rawFields.length; i += 1) {
      var field = rawFields[i];
      var key = String(field.Name || field.RealFieldName || '');
      var real = String(field.RealFieldName || '');
      if (key) {
        byName[key] = field;
      }
      if (real) {
        byName[real] = field;
      }
    }

    var mapped: IListFieldDefinition[] = [];
    for (var j = 0; j < viewFieldNames.length; j += 1) {
      var viewFieldName = String(viewFieldNames[j] || '');
      if (!viewFieldName) {
        continue;
      }

      var match = byName[viewFieldName];
      if (match) {
        mapped.push(match);
      } else if (!isSystemItemKey(viewFieldName)) {
        mapped.push({ Name: viewFieldName, DisplayName: viewFieldName });
      }
    }

    return mapped;
  }

  private async loadListFieldTypeMap(viewFieldNames: string[]): Promise<{ [internalName: string]: string }> {
    if (!viewFieldNames || viewFieldNames.length === 0) {
      return {};
    }

    try {
      var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=InternalName,TypeAsString,DisplayFormat";
      var response = await this.getJsonWithFallback(endpoint);
      if (!response.ok) {
        return {};
      }

      var data = await response.json();
      var fields = toArray(data.value);
      if (fields.length === 0) {
        fields = toArray(data && data.d && data.d.results);
      }

      var requested: { [name: string]: boolean } = {};
      for (var i = 0; i < viewFieldNames.length; i += 1) {
        requested[String(viewFieldNames[i] || '')] = true;
      }

      var map: { [internalName: string]: string } = {};
      for (var j = 0; j < fields.length; j += 1) {
        var field = fields[j] || {};
        var internalName = String(field.InternalName || '');
        if (!internalName || !requested[internalName]) {
          continue;
        }
        map[internalName] = String(field.TypeAsString || '');
        map[internalName.toLowerCase()] = String(field.TypeAsString || '');
        if (String(field.TypeAsString || '').toLowerCase() === 'datetime') {
          var displayFormat = parseInt(String(field.DisplayFormat), 10);
          this._fieldDisplayFormatMap[internalName.toLowerCase()] = isNaN(displayFormat) ? 1 : displayFormat;
        }
      }

      return map;
    } catch (error) {
      this.logDiagnostic('loadListFieldTypeMap failed: ' + (error && error.message ? error.message : String(error)));
      return {};
    }
  }

  private async loadListFieldTitleMap(): Promise<{ [internalName: string]: string }> {
    try {
      var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=InternalName,Title";
      var response = await this.getJsonWithFallback(endpoint);
      if (!response.ok) {
        return {};
      }

      var data = await response.json();
      var fields = toArray(data.value);
      if (fields.length === 0) {
        fields = toArray(data && data.d && data.d.results);
      }

      var map: { [internalName: string]: string } = {};
      for (var i = 0; i < fields.length; i += 1) {
        var field = fields[i] || {};
        var internalName = String(field.InternalName || '');
        var title = String(field.Title || '');
        if (internalName && title) {
          map[internalName.toLowerCase()] = title;
        }
      }

      return map;
    } catch (error) {
      this.logDiagnostic('loadListFieldTitleMap failed: ' + (error && error.message ? error.message : String(error)));
      return {};
    }
  }

  private applyFieldDisplayNames(fields: IListFieldDefinition[], titleMap: { [internalName: string]: string }): IListFieldDefinition[] {
    return fields.map(function(field: IListFieldDefinition) {
      var internalName = String(field.RealFieldName || field.Name || '');
      var displayName = titleMap[internalName.toLowerCase()] || field.DisplayName || internalName;
      return {
        ...field,
        DisplayName: displayName
      };
    });
  }

  private applyFieldTypes(fields: IListFieldDefinition[], typeMap: { [internalName: string]: string }): IListFieldDefinition[] {
    return fields.map(function(field: IListFieldDefinition) {
      var internalName = String(field.RealFieldName || field.Name || '');
      return {
        ...field,
        TypeAsString: typeMap[internalName] || typeMap[internalName.toLowerCase()] || field.TypeAsString || '',
        DisplayFormat: this._fieldDisplayFormatMap[internalName.toLowerCase()]
      };
    }.bind(this));
  }

  private getRowFieldValue(row: any, field: IListFieldDefinition): any {
    var fieldName = field.Name || field.RealFieldName || '';
    var value = row[fieldName];
    if ((value === undefined || value === null || value === '') && field.RealFieldName) {
      value = row[field.RealFieldName];
    }
    return value;
  }

  private getUrlCellValue(row: any, field: IListFieldDefinition): { href: string; text: string } | null {
    var fieldType = String(field.TypeAsString || '').toLowerCase();
    if (fieldType !== 'url' && fieldType !== 'hyperlink') { return null; }

    var rawValue = this.getRowFieldValue(row, field);
    if (rawValue === undefined || rawValue === null || rawValue === '') { return null; }

    var href = '';
    var text = '';
    if (typeof rawValue === 'object') {
      href = String((rawValue as any).Url || (rawValue as any).url || '').trim();
      text = String((rawValue as any).Description || (rawValue as any).description || '').trim();
    } else {
      var rawText = String(rawValue).trim();
      if (rawText.indexOf('<') >= 0 && rawText.indexOf('>') >= 0) {
        var container = document.createElement('div');
        container.innerHTML = rawText;
        var anchor = container.querySelector('a');
        if (anchor) {
          href = String(anchor.getAttribute('href') || '').trim();
          text = String(anchor.textContent || '').trim();
        }
      }
      if (!href) {
        var descriptionSeparator = rawText.indexOf(', ');
        href = descriptionSeparator > 0 ? rawText.substring(0, descriptionSeparator).trim() : rawText;
        text = descriptionSeparator > 0 ? rawText.substring(descriptionSeparator + 2).trim() : '';
      }
    }

    if (!href || /^\s*(?:javascript|data|vbscript):/i.test(href)) { return null; }
    return { href: href, text: text || href };
  }

  private stringifyCellValue(value: any): string {
    if (value === undefined || value === null || value === '') {
      return '';
    }

    if (Array.isArray(value)) {
      var listValues: string[] = [];
      for (var i = 0; i < value.length; i += 1) {
        var itemText = this.stringifyCellValue(value[i]);
        if (itemText) {
          listValues.push(itemText);
        }
      }
      return listValues.join('; ');
    }

    if (typeof value === 'object') {
      if (Array.isArray((value as any).results)) {
        return this.stringifyCellValue((value as any).results);
      }
      if ((value as any).Title) {
        return String((value as any).Title);
      }
      if ((value as any).LookupValue) {
        return String((value as any).LookupValue);
      }
      if ((value as any).Name) {
        return String((value as any).Name);
      }
      if ((value as any).Email) {
        return String((value as any).Email);
      }
      if ((value as any).Url) {
        return String((value as any).Url);
      }
      if ((value as any).Id !== undefined && (value as any).Id !== null) {
        return String((value as any).Id);
      }
      try {
        return JSON.stringify(value);
      } catch (_jsonError) {
        return String(value);
      }
    }

    return String(value);
  }

  private formatDateCellValue(value: any, field: IListFieldDefinition): string {
    var rawValue = this.stringifyCellValue(value);
    if (String(field.TypeAsString || '').toLowerCase() !== 'datetime') {
      return rawValue;
    }

    var dateOnlyMatch = field.DisplayFormat === 0 ? /^(\d{4})-(\d{2})-(\d{2})/.exec(rawValue) : null;
    var dateValue = new Date(rawValue);
    if (isNaN(dateValue.getTime())) {
      return rawValue;
    }

    var month = dateOnlyMatch ? dateOnlyMatch[2] : String(dateValue.getMonth() + 1);
    var day = dateOnlyMatch ? dateOnlyMatch[3] : String(dateValue.getDate());
    var year = dateOnlyMatch ? dateOnlyMatch[1] : String(dateValue.getFullYear());
    month = month.length < 2 ? '0' + month : month;
    day = day.length < 2 ? '0' + day : day;

    var dateText = month + '/' + day + '/' + year;
    if (this.props.dateDisplayFormat === 'dmy') {
      dateText = day + '/' + month + '/' + year;
    } else if (this.props.dateDisplayFormat === 'ymd') {
      dateText = year + '-' + month + '-' + day;
    }

    if (field.DisplayFormat === 0) {
      return dateText;
    }

    var hours = dateValue.getHours();
    var minutes = String(dateValue.getMinutes());
    minutes = minutes.length < 2 ? '0' + minutes : minutes;
    if (this.props.timeDisplayFormat === '12hour') {
      var period = hours >= 12 ? 'PM' : 'AM';
      var twelveHour = hours % 12 || 12;
      return dateText + ' ' + (twelveHour < 10 ? '0' : '') + String(twelveHour) + ':' + minutes + ' ' + period;
    }

    return dateText + ' ' + (hours < 10 ? '0' : '') + String(hours) + ':' + minutes;
  }

  private isMeaningfulCellValue(value: any): boolean {
    if (value === undefined || value === null) {
      return false;
    }

    var text = this.stringifyCellValue(value).replace(/<[^>]*>/g, '').replace(/&nbsp;/gi, '').trim();
    return text.length > 0;
  }

  private filterRenderableRows(rows: any[], visibleFields: IListFieldDefinition[]): any[] {
    var filtered: any[] = [];

    for (var i = 0; i < rows.length; i += 1) {
      var row = rows[i] || {};
      if (this.getRowItemId(row) > 0) {
        filtered.push(row);
        continue;
      }
      var hasVisibleValue = false;
      for (var j = 0; j < visibleFields.length; j += 1) {
        var field = visibleFields[j];
        var fieldName = String(field.Name || field.RealFieldName || '');
        var realFieldName = String(field.RealFieldName || '');
        var candidateValue = row[fieldName];

        if ((candidateValue === undefined || candidateValue === null || candidateValue === '') && realFieldName) {
          candidateValue = row[realFieldName];
        }

        if (this.isMeaningfulCellValue(candidateValue)) {
          hasVisibleValue = true;
          break;
        }
      }

      // Rows without a standard item ID still require visible content.
      if (hasVisibleValue) {
        filtered.push(row);
      }
    }

    return filtered;
  }

  private async loadRowsFromItemsEndpoint(viewFieldNames: string[]): Promise<{ rows: any[]; fields: IListFieldDefinition[] }> {
    var fieldTypeMap = await this.loadListFieldTypeMap(viewFieldNames);
    var selectFields = ['ID'];
    var expandFields: string[] = [];
    for (var i = 0; i < viewFieldNames.length; i += 1) {
      var fieldName = String(viewFieldNames[i] || '');
      if (!fieldName) {
        continue;
      }

      var fieldType = String(fieldTypeMap[fieldName] || '');
      if (isExpandableFieldType(fieldType)) {
        if (selectFields.indexOf(fieldName + '/Id') < 0) {
          selectFields.push(fieldName + '/Id');
        }
        if (selectFields.indexOf(fieldName + '/Title') < 0) {
          selectFields.push(fieldName + '/Title');
        }
        if (expandFields.indexOf(fieldName) < 0) {
          expandFields.push(fieldName);
        }
      } else if (selectFields.indexOf(fieldName) < 0) {
        selectFields.push(fieldName);
      }
    }

    var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items?$top=200";
    if (selectFields.length > 0) {
      endpoint += '&$select=' + encodeURIComponent(selectFields.join(','));
    }
    if (expandFields.length > 0) {
      endpoint += '&$expand=' + encodeURIComponent(expandFields.join(','));
    }

    var response = await this.getJsonWithFallback(endpoint);
    this.logDiagnostic('Items endpoint fallback executed. Url=' + endpoint + ', ok=' + String(response.ok));

    if (!response.ok) {
      return { rows: [], fields: [] };
    }

    var data = await response.json();
    var rows = toArray(data.value);
    if (rows.length === 0) {
      rows = toArray(data && data.d && data.d.results);
    }

    var fields: IListFieldDefinition[] = [];
    if (viewFieldNames.length > 0) {
      for (var v = 0; v < viewFieldNames.length; v += 1) {
        var viewFieldName = String(viewFieldNames[v] || '');
        if (viewFieldName && !isSystemItemKey(viewFieldName)) {
          fields.push({ Name: viewFieldName, DisplayName: viewFieldName });
        }
      }
    } else if (rows.length > 0) {
      var firstRow = rows[0] || {};
      for (var key in firstRow) {
        if (firstRow.hasOwnProperty(key) && !isSystemItemKey(key)) {
          fields.push({ Name: key, DisplayName: key });
        }
      }
    }

    return {
      rows: rows,
      fields: fields
    };
  }

  private async loadRows(): Promise<void> {
    if (!this.props.listName) {
      this.setState({ loading: false, error: null, fields: [], rows: [] });
      return;
    }

    var requestId = ++this._loadRowsRequestId;
    this.logDiagnostic('Starting loadRows. requestId=' + String(requestId) + ', listName=' + String(this.props.listName) + ', selectedViewId=' + String(this.state.selectedViewId || '(none)'));
    this.setState({ loading: true, error: null });

    try {
      var baseEndpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/RenderListDataAsStream";
      var selectedViewId = this.state.selectedViewId;
      var viewFieldNames = await this.loadSelectedViewFieldNames(selectedViewId);
      var selectedViewXml = await this.loadSelectedViewXml(selectedViewId, viewFieldNames);

      var body: any = {
        parameters: {
          RenderOptions: 7
        }
      };
      if (selectedViewXml) {
        body.parameters.ViewXml = selectedViewXml;
      }

      var selectedRows: any[] = [];
      var selectedFields: IListFieldDefinition[] = [];
      var lastError: string = '';
      var hadSuccessfulResponse = false;
      var requestUrls = [baseEndpoint];

      for (var requestIndex = 0; requestIndex < requestUrls.length; requestIndex += 1) {
        var requestUrl = requestUrls[requestIndex];
        var response = await this.postJsonWithFallback(requestUrl, body);

        if (!response.ok) {
          var errorText = '';
          try {
            errorText = await response.text();
          } catch (_readError) {
            errorText = '';
          }
          lastError = 'Failed to load list view data. HTTP ' + String(response.status) + ' ' + response.statusText + (errorText ? (': ' + errorText) : '');
          continue;
        }

        hadSuccessfulResponse = true;

        var data = await response.json();
        var extracted = extractRenderRowsAndFields(data);
        this.logDiagnostic('RenderListDataAsStream parsed. rows=' + String(extracted.rows.length) + ', fields=' + String(extracted.fields.length));

        if (selectedFields.length === 0 && extracted.fields.length > 0) {
          selectedFields = extracted.fields;
        }

        if (extracted.rows.length > 0 || requestIndex === requestUrls.length - 1) {
          selectedRows = extracted.rows;
          if (selectedFields.length === 0) {
            selectedFields = extracted.fields;
          }
          break;
        }
      }

      if (!hadSuccessfulResponse) {
        throw new Error(lastError || 'Failed to load list view data.');
      }

      var rows = selectedRows;
      var fields = selectedFields;
      if (rows.length === 0 && !selectedViewId) {
        var itemsFallback = await this.loadRowsFromItemsEndpoint(viewFieldNames);
        rows = itemsFallback.rows;

        if (fields.length === 0) {
          fields = itemsFallback.fields;
        }
      }

      var visibleFields = this.getFieldsForConsumption(fields, viewFieldNames);
      var fieldTitleMap = await this.loadListFieldTitleMap();
      var visibleFieldNames = visibleFields.map((field: IListFieldDefinition) => this.getFieldKey(field));
      var fieldTypeMap = await this.loadListFieldTypeMap(visibleFieldNames);
      visibleFields = this.applyFieldDisplayNames(visibleFields, fieldTitleMap);
      visibleFields = this.applyFieldTypes(visibleFields, fieldTypeMap);
      var renderableRows = this.filterRenderableRows(rows, visibleFields);

      if (requestId !== this._loadRowsRequestId) {
        this.logDiagnostic('Ignoring stale loadRows result. requestId=' + String(requestId) + ', latestRequestId=' + String(this._loadRowsRequestId));
        return;
      }

      this.setState({
        fields: visibleFields,
        rows: renderableRows,
        loading: false,
        error: null
      });
      this.logDiagnostic('loadRows completed. visibleFields=' + String(visibleFields.length) + ', renderableRows=' + String(renderableRows.length));
    } catch (error) {
      var loadError: any = error as any;
      if (requestId !== this._loadRowsRequestId) {
        this.logDiagnostic('Ignoring stale loadRows failure. requestId=' + String(requestId) + ', latestRequestId=' + String(this._loadRowsRequestId));
        return;
      }
      this.setState({
        loading: false,
        error: loadError && loadError.message ? loadError.message : 'Failed to load data.',
        fields: [],
        rows: []
      });
      this.logDiagnostic('loadRows failed: ' + (loadError && loadError.message ? loadError.message : String(loadError)));
    }
  }

  private getRowItemId(row: any): number {
    return toPositiveInt(row.ID || row.Id || row.id);
  }

  private selectRow(row: any): void {
    var itemId = this.getRowItemId(row);
    this.setState({ selectedItemId: itemId, selectedMode: 'view' });
    this.logDiagnostic('Row selected. itemId=' + String(itemId));
    this.props.onSelectionChange(itemId, 'view');
  }

  private async deleteSelected(): Promise<void> {
    if (this.state.selectedItemId <= 0) {
      return;
    }

    if (typeof window !== 'undefined' && !window.confirm(strings.RuntimeDeleteConfirm)) {
      return;
    }

    this.logDiagnostic('Deleting item. itemId=' + String(this.state.selectedItemId));
    this.setState({ deleting: true, error: null });

    try {
      var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items(" + this.state.selectedItemId + ")";
      var response = await this.props.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }
      );

      if (!response.ok) {
        throw new Error(strings.RuntimeDeleteFailed);
      }

      this.setState({ selectedItemId: 0, selectedMode: 'view', deleting: false });
      this.props.onSelectionChange(0, 'view');
      this.logDiagnostic('Item deleted successfully. itemId=' + String(this.state.selectedItemId));
      await this.loadRows();
    } catch (error) {
      var deleteError: any = error as any;
      this.setState({
        deleting: false,
        error: deleteError && deleteError.message ? deleteError.message : strings.RuntimeDeleteFailed
      });
      this.logDiagnostic('Delete failed: ' + (deleteError && deleteError.message ? deleteError.message : String(deleteError)));
    }
  }

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }

    console.log('[ListControl] ' + message);
  }

  private getCellMarkup(row: any, field: IListFieldDefinition): { __html: string } | null {
    var value = this.getRowFieldValue(row, field);

    if (value === undefined || value === null || value === '') {
      return null;
    }

    var text = this.formatDateCellValue(value, field);
    if (text.indexOf('<') >= 0 && text.indexOf('>') >= 0) {
      return { __html: text };
    }

    return { __html: text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') };
  }

  private isTitleField(field: IListFieldDefinition): boolean {
    var fieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
    var displayName = String(field.DisplayName || '').toLowerCase();

    // RenderListDataAsStream often returns link-capable title/name aliases.
    return fieldName === 'title'
      || fieldName === 'linktitle'
      || fieldName === 'linktitlenomenu'
      || fieldName === 'linkfilename'
      || fieldName === 'linkfilename2'
      || fieldName === 'fileleafref'
      || displayName === 'title'
      || displayName === 'name';
  }

  private getCellPlainText(row: any, field: IListFieldDefinition): string {
    var value = this.getRowFieldValue(row, field);

    if (value === undefined || value === null || value === '') {
      return '';
    }

    return this.stringifyCellValue(value).replace(/<[^>]*>/g, '').trim();
  }

  private getItemLinkUrl(row: any): string {
    var itemId = this.getRowItemId(row);
    var targetPageUrl = String(this.props.linkTargetPageUrl || '').trim();
    var targetIdParam = String(this.props.linkTargetIdParam || 'itemid').trim() || 'itemid';

    if (targetPageUrl && itemId > 0) {
      var normalizedTargetUrl = targetPageUrl;
      if (!/^https?:\/\//i.test(normalizedTargetUrl)) {
        normalizedTargetUrl = normalizedTargetUrl.charAt(0) === '/'
          ? this.getWebUrl() + normalizedTargetUrl
          : this.getWebUrl() + '/' + normalizedTargetUrl;
      }

      var embeddedPageUrl = resolveEmbeddedPageUrl(normalizedTargetUrl);
      if (embeddedPageUrl) {
        if (/^https?:\/\//i.test(embeddedPageUrl)) {
          normalizedTargetUrl = embeddedPageUrl;
        } else if (embeddedPageUrl.charAt(0) === '/') {
          normalizedTargetUrl = this.getWebUrl() + embeddedPageUrl;
        } else {
          normalizedTargetUrl = this.getWebUrl() + '/' + embeddedPageUrl;
        }
      }

      if (isWorkbenchUrl(normalizedTargetUrl) || isSharePointWrapperUrl(normalizedTargetUrl)) {
        var sourceUrl = resolveEmbeddedPageUrl(normalizedTargetUrl);
        if (sourceUrl && !isWorkbenchUrl(sourceUrl) && !isSharePointWrapperUrl(sourceUrl)) {
          normalizedTargetUrl = sourceUrl;
        } else if (typeof window !== 'undefined' && window.location && !isWorkbenchUrl(window.location.pathname)) {
          normalizedTargetUrl = window.location.protocol + '//' + window.location.host + stripQueryAndHash(window.location.pathname);
        } else {
          normalizedTargetUrl = '';
        }
      }

      if (normalizedTargetUrl) {
        var targetUrlWithItem = appendQueryParam(normalizedTargetUrl, targetIdParam, String(itemId));
        if (this.props.includeReturnUrlParam === true && typeof window !== 'undefined' && window.location && window.location.href) {
          return appendQueryParam(targetUrlWithItem, 'return', window.location.href);
        }
        return targetUrlWithItem;
      }
    }

    var fileRef = String(row.FileRef || '').trim();

    if (fileRef) {
      if (/^https?:\/\//i.test(fileRef)) {
        return fileRef;
      }

      // List item pseudo files usually end with _.000; avoid those and use form URL instead.
      if (fileRef.charAt(0) === '/' && !/\/\d+_\.000$/i.test(fileRef)) {
        return this.getWebUrl() + fileRef;
      }
    }

    if (itemId > 0) {
      return this.getWebUrl() + '/Lists/' + encodeURIComponent(this.props.listName) + '/DispForm.aspx?ID=' + itemId;
    }

    return '';
  }

  private getFieldKey(field: IListFieldDefinition): string {
    return String(field.Name || field.RealFieldName || '').trim();
  }

  private getFilterOperatorOptions(): Array<{ key: FilterOperator; label: string }> {
    return [
      { key: 'eq', label: strings.RuntimeFilterOperatorEquals },
      { key: 'ne', label: strings.RuntimeFilterOperatorNotEquals },
      { key: 'contains', label: strings.RuntimeFilterOperatorContains },
      { key: 'notcontains', label: strings.RuntimeFilterOperatorNotContains },
      { key: 'startswith', label: strings.RuntimeFilterOperatorStartsWith },
      { key: 'endswith', label: strings.RuntimeFilterOperatorEndsWith },
      { key: 'gt', label: strings.RuntimeFilterOperatorGreaterThan },
      { key: 'ge', label: strings.RuntimeFilterOperatorGreaterThanOrEqual },
      { key: 'lt', label: strings.RuntimeFilterOperatorLessThan },
      { key: 'le', label: strings.RuntimeFilterOperatorLessThanOrEqual }
    ];
  }

  private toggleSort(field: IListFieldDefinition): void {
    var fieldKey = this.getFieldKey(field);
    if (!fieldKey) {
      return;
    }

    var nextDirection: 'asc' | 'desc' | '' = 'asc';
    if (this.state.sortFieldName === fieldKey && this.state.sortDirection === 'asc') {
      nextDirection = 'desc';
    } else if (this.state.sortFieldName === fieldKey && this.state.sortDirection === 'desc') {
      nextDirection = '';
    }

    this.setState({
      sortFieldName: nextDirection ? fieldKey : '',
      sortDirection: nextDirection,
      currentPage: 0
    });
  }

  private openFilter(field: IListFieldDefinition, anchorElement: HTMLElement): void {
    var fieldKey = this.getFieldKey(field);
    if (!fieldKey) {
      return;
    }

    var anchorRect = anchorElement.getBoundingClientRect();
    var viewportWidth = document.documentElement.clientWidth || window.innerWidth;
    var viewportHeight = document.documentElement.clientHeight || window.innerHeight;
    var popoverWidth = 220;
    var estimatedPopoverHeight = 190;
    var viewportMargin = 8;
    var popoverGap = 6;
    var left = anchorRect.right - popoverWidth;
    left = Math.max(viewportMargin, Math.min(left, viewportWidth - popoverWidth - viewportMargin));
    var availableBelow = viewportHeight - anchorRect.bottom - popoverGap - viewportMargin;
    var availableAbove = anchorRect.top - popoverGap - viewportMargin;
    var popoverStyle: any = {
      left: left,
      right: 'auto'
    };

    if (availableBelow < estimatedPopoverHeight && availableAbove > availableBelow) {
      popoverStyle.top = 'auto';
      popoverStyle.bottom = viewportHeight - anchorRect.top + popoverGap;
      popoverStyle.maxHeight = Math.max(120, availableAbove);
    } else {
      popoverStyle.top = anchorRect.bottom + popoverGap;
      popoverStyle.bottom = 'auto';
      popoverStyle.maxHeight = Math.max(120, availableBelow);
    }

    var existing = this.state.columnFilters[fieldKey];
    var isDateField = String(field.TypeAsString || '').toLowerCase() === 'datetime';
    this.setState({
      activeFilterFieldName: fieldKey,
      filterPopoverStyle: popoverStyle,
      activeFilterIsDate: isDateField,
      draftFilterOperator: isDateField ? 'eq' : (existing ? existing.operator : 'contains'),
      draftFilterValue: existing ? existing.value : '',
      draftFilterEndValue: existing ? String(existing.endValue || '') : '',
      datePickerTarget: '',
      datePickerMonth: ''
    });
  }

  private closeFilter(): void {
    this.setState({ activeFilterFieldName: '', datePickerTarget: '' });
  }

  private getIsoDate(date: Date): string {
    var month = String(date.getMonth() + 1);
    var day = String(date.getDate());
    return String(date.getFullYear()) + '-' + (month.length < 2 ? '0' + month : month) + '-' + (day.length < 2 ? '0' + day : day);
  }

  private toggleDatePicker(target: 'start' | 'end'): void {
    if (this.state.datePickerTarget === target) {
      this.setState({ datePickerTarget: '' });
      return;
    }

    var value = target === 'start' ? this.state.draftFilterValue : this.state.draftFilterEndValue;
    var month = /^\d{4}-\d{2}-\d{2}$/.test(value) ? value.substr(0, 7) : this.getIsoDate(new Date()).substr(0, 7);
    this.setState({ datePickerTarget: target, datePickerMonth: month });
  }

  private changeDatePickerMonth(offset: number): void {
    var parts = this.state.datePickerMonth.split('-');
    var monthDate = new Date(Number(parts[0]), Number(parts[1]) - 1 + offset, 1);
    this.setState({ datePickerMonth: this.getIsoDate(monthDate).substr(0, 7) });
  }

  private selectFilterDate(value: string): void {
    if (this.state.datePickerTarget === 'end') {
      this.setState({ draftFilterEndValue: value, datePickerTarget: '' });
    } else {
      this.setState({ draftFilterValue: value, datePickerTarget: '' });
    }
  }

  private renderDatePicker(): React.ReactNode {
    if (!this.state.datePickerTarget || !/^\d{4}-\d{2}$/.test(this.state.datePickerMonth)) {
      return null;
    }

    var parts = this.state.datePickerMonth.split('-');
    var year = Number(parts[0]);
    var monthIndex = Number(parts[1]) - 1;
    var firstWeekday = new Date(year, monthIndex, 1).getDay();
    var dayCount = new Date(year, monthIndex + 1, 0).getDate();
    var selectedValue = this.state.datePickerTarget === 'start' ? this.state.draftFilterValue : this.state.draftFilterEndValue;
    var cells: React.ReactNode[] = [];
    var index: number;
    for (index = 0; index < firstWeekday; index += 1) {
      cells.push(<span key={'blank-' + index} className="lc-date-picker-blank" />);
    }
    for (index = 1; index <= dayCount; index += 1) {
      let dateValue = this.getIsoDate(new Date(year, monthIndex, index));
      let disabled = this.state.datePickerTarget === 'end' && !!this.state.draftFilterValue && dateValue < this.state.draftFilterValue;
      cells.push(
        <button key={dateValue} type="button" className={dateValue === selectedValue ? 'lc-date-picker-day lc-date-picker-selected' : 'lc-date-picker-day'} disabled={disabled} onClick={() => this.selectFilterDate(dateValue)}>
          {index}
        </button>
      );
    }

    return (
      <div className="lc-date-picker" role="dialog" aria-label="Choose date">
        <div className="lc-date-picker-header">
          <button type="button" title="Previous month" aria-label="Previous month" onClick={() => this.changeDatePickerMonth(-1)}>&#8249;</button>
          <strong>{new Date(year, monthIndex, 1).toLocaleString(undefined, { month: 'long', year: 'numeric' })}</strong>
          <button type="button" title="Next month" aria-label="Next month" onClick={() => this.changeDatePickerMonth(1)}>&#8250;</button>
        </div>
        <div className="lc-date-picker-weekdays"><span>Su</span><span>Mo</span><span>Tu</span><span>We</span><span>Th</span><span>Fr</span><span>Sa</span></div>
        <div className="lc-date-picker-days">{cells}</div>
      </div>
    );
  }

  private applyActiveFilter(): void {
    var fieldKey = String(this.state.activeFilterFieldName || '');
    if (!fieldKey) {
      return;
    }

    var nextFilters: { [fieldName: string]: IColumnFilter } = {};
    var key: string;
    for (key in this.state.columnFilters) {
      if (Object.prototype.hasOwnProperty.call(this.state.columnFilters, key)) {
        nextFilters[key] = this.state.columnFilters[key];
      }
    }

    var value = String(this.state.draftFilterValue || '').trim();
    var endValue = String(this.state.draftFilterEndValue || '').trim();
    if (!value) {
      delete nextFilters[fieldKey];
    } else if (this.state.activeFilterIsDate) {
      if (endValue && endValue < value) {
        var originalValue = value;
        value = endValue;
        endValue = originalValue;
      }
      nextFilters[fieldKey] = {
        operator: 'eq',
        value: value,
        endValue: endValue,
        compareDateOnly: true
      };
    } else {
      nextFilters[fieldKey] = {
        operator: this.state.draftFilterOperator,
        value: value
      };
    }

    this.setState({
      columnFilters: nextFilters,
      activeFilterFieldName: '',
      currentPage: 0
    });
  }

  private clearActiveFilter(): void {
    var fieldKey = String(this.state.activeFilterFieldName || '');
    if (!fieldKey) {
      return;
    }

    var nextFilters: { [fieldName: string]: IColumnFilter } = {};
    var key: string;
    for (key in this.state.columnFilters) {
      if (Object.prototype.hasOwnProperty.call(this.state.columnFilters, key) && key !== fieldKey) {
        nextFilters[key] = this.state.columnFilters[key];
      }
    }

    this.setState({
      columnFilters: nextFilters,
      draftFilterOperator: 'contains',
      draftFilterValue: '',
      draftFilterEndValue: '',
      activeFilterFieldName: '',
      currentPage: 0
    });
  }

  private compareComparableValues(leftValue: string, rightValue: string): number {
    var leftNumber = Number(leftValue);
    var rightNumber = Number(rightValue);
    if (!isNaN(leftNumber) && !isNaN(rightNumber)) {
      if (leftNumber > rightNumber) {
        return 1;
      }
      if (leftNumber < rightNumber) {
        return -1;
      }
      return 0;
    }

    var leftDate = Date.parse(leftValue);
    var rightDate = Date.parse(rightValue);
    if (!isNaN(leftDate) && !isNaN(rightDate)) {
      if (leftDate > rightDate) {
        return 1;
      }
      if (leftDate < rightDate) {
        return -1;
      }
      return 0;
    }

    var leftText = leftValue.toLowerCase();
    var rightText = rightValue.toLowerCase();
    if (leftText > rightText) {
      return 1;
    }
    if (leftText < rightText) {
      return -1;
    }
    return 0;
  }

  private rowMatchesFilter(row: any, field: IListFieldDefinition, filter: IColumnFilter): boolean {
    if (Array.isArray(filter.value)) {
      var values = filter.value.filter(function(value: any) { return value !== undefined && value !== null && String(value).trim() !== ''; });
      if (values.length === 0) { return true; }
      var isNegative = filter.operator === 'ne' || filter.operator === 'notcontains';
      var scalarOperator: FilterOperator = filter.operator === 'ne' ? 'eq' : (filter.operator === 'notcontains' ? 'contains' : filter.operator);
      var scalarMatches = values.map((value: any) => this.rowMatchesFilter(row, field, { operator: scalarOperator, value: value, compareDateOnly: filter.compareDateOnly }));
      return isNegative ? scalarMatches.every(function(matches: boolean) { return !matches; }) : scalarMatches.some(function(matches: boolean) { return matches; });
    }
    var valueText = this.getFilterCellText(row, field);
    var candidate = String(valueText || '').trim();
    var query = String(filter.value || '').trim();

    if (!query) {
      return true;
    }

    var normalizedCandidate = candidate.toLowerCase();
    var normalizedQuery = query.toLowerCase();
    var compareResult = this.compareComparableValues(candidate, query);
    var fieldType = String(field.TypeAsString || '').toLowerCase();
    var lookupCandidates = fieldType.indexOf('lookup') >= 0
      ? normalizedCandidate.split(';').map(function(value) { return value.trim(); }).filter(function(value) { return !!value; })
      : [];

    if (filter.compareDateOnly) {
      var candidateDate = new Date(candidate);
      if (isNaN(candidateDate.getTime())) {
        return false;
      }
      normalizedCandidate = formatLocalDate(candidateDate).toLowerCase();
      compareResult = this.compareComparableValues(normalizedCandidate, query);
      if (filter.endValue) {
        return normalizedCandidate >= normalizedQuery && normalizedCandidate <= String(filter.endValue).toLowerCase();
      }
    }

    switch (filter.operator) {
      case 'eq':
        return lookupCandidates.length > 0 ? lookupCandidates.indexOf(normalizedQuery) >= 0 : normalizedCandidate === normalizedQuery;
      case 'ne':
        return lookupCandidates.length > 0 ? lookupCandidates.indexOf(normalizedQuery) < 0 : normalizedCandidate !== normalizedQuery;
      case 'contains':
        return normalizedCandidate.indexOf(normalizedQuery) >= 0;
      case 'notcontains':
        return normalizedCandidate.indexOf(normalizedQuery) < 0;
      case 'startswith':
        return normalizedCandidate.indexOf(normalizedQuery) === 0;
      case 'endswith':
        return normalizedCandidate.lastIndexOf(normalizedQuery) === normalizedCandidate.length - normalizedQuery.length;
      case 'gt':
        return compareResult > 0;
      case 'ge':
        return compareResult >= 0;
      case 'lt':
        return compareResult < 0;
      case 'le':
        return compareResult <= 0;
      default:
        return true;
    }
  }

  private getFilterCellText(row: any, field: IListFieldDefinition): string {
    var fieldType = String(field.TypeAsString || '').toLowerCase();
    if (fieldType.indexOf('lookup') < 0) {
      return this.getCellPlainText(row, field);
    }

    var rawValue = this.getRowFieldValue(row, field);
    var lookupIds: string[] = [];
    var collectLookupIds = function(value: any): void {
      if (value === undefined || value === null || value === '') {
        return;
      }
      if (Array.isArray(value)) {
        for (var valueIndex = 0; valueIndex < value.length; valueIndex += 1) {
          collectLookupIds(value[valueIndex]);
        }
        return;
      }
      if (typeof value === 'object') {
        var lookupId = value.LookupId !== undefined ? value.LookupId
          : value.lookupId !== undefined ? value.lookupId
            : value.Id !== undefined ? value.Id
              : value.ID;
        if (lookupId !== undefined && lookupId !== null && String(lookupId).trim()) {
          lookupIds.push(String(lookupId).trim());
        }
      }
    };
    collectLookupIds(rawValue);

    if (lookupIds.length === 0) {
      var lookupFieldNames = [String(field.RealFieldName || ''), String(field.Name || '')];
      for (var fieldIndex = 0; fieldIndex < lookupFieldNames.length; fieldIndex += 1) {
        var fieldName = lookupFieldNames[fieldIndex];
        if (!fieldName) {
          continue;
        }
        var companionValues = [row[fieldName + 'Id'], row[fieldName + '.lookupId']];
        for (var companionIndex = 0; companionIndex < companionValues.length; companionIndex += 1) {
          var companionValue = companionValues[companionIndex];
          if (Array.isArray(companionValue)) {
            for (var lookupIndex = 0; lookupIndex < companionValue.length; lookupIndex += 1) {
              lookupIds.push(String(companionValue[lookupIndex]).trim());
            }
          } else if (companionValue !== undefined && companionValue !== null && String(companionValue).trim()) {
            lookupIds.push(String(companionValue).trim());
          }
        }
      }
    }
    return lookupIds.length > 0 ? lookupIds.join('; ') : this.getCellPlainText(row, field);
  }

  private parsePresetFilterConditions(): IPresetFilterCondition[] {
    var sources = [String(this.props.filterJson || '').trim(), String(this.state.runtimeFilterJson || '').trim()];
    var conditions: IPresetFilterCondition[] = [];
    for (var sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
      var source = sources[sourceIndex];
      if (!source) {
        continue;
      }
      try {
        var parsed = JSON.parse(source);
        if (!Array.isArray(parsed)) {
          continue;
        }
      for (var i = 0; i < parsed.length; i += 1) {
        var item = parsed[i] || {};
        var field = String(item.field || '').trim();
        if (!field) {
          continue;
        }

        conditions.push({
          field: field,
          operator: normalizeFilterOperator(item.operator),
          logical: normalizeFilterLogical(item.logical),
          valueType: String(item.valueType || '').toLowerCase() === 'expression' ? 'expression' : (String(item.valueType || '').toLowerCase() === 'fieldvalue' ? 'fieldValue' : 'static'),
          value: item.value
        });
      }
      } catch (_parseError) {
        this.logDiagnostic('Preset filter JSON is invalid; skipping that preset filter source.');
      }
    }
    return conditions;
  }

  private resolveFieldByReference(fieldsByKey: { [key: string]: IListFieldDefinition }, fieldRef: string): IListFieldDefinition | undefined {
    var direct = fieldsByKey[fieldRef];
    if (direct) {
      return direct;
    }

    var normalizedRef = String(fieldRef || '').trim().toLowerCase();
    if (!normalizedRef) {
      return undefined;
    }

    var key: string;
    for (key in fieldsByKey) {
      if (!Object.prototype.hasOwnProperty.call(fieldsByKey, key)) {
        continue;
      }

      var candidate = fieldsByKey[key];
      if (!candidate) {
        continue;
      }

      var byName = String(candidate.Name || '').toLowerCase();
      var byRealName = String(candidate.RealFieldName || '').toLowerCase();
      var byDisplay = String(candidate.DisplayName || '').toLowerCase();
      if (normalizedRef === byName || normalizedRef === byRealName || normalizedRef === byDisplay) {
        return candidate;
      }
    }

    return undefined;
  }

  private rowMatchesPresetConditions(row: any, fieldsByKey: { [key: string]: IListFieldDefinition }, conditions: IPresetFilterCondition[]): boolean {
    if (conditions.length === 0) {
      return true;
    }

    var aggregate: boolean | null = null;

    for (var i = 0; i < conditions.length; i += 1) {
      var condition = conditions[i];
      var field = this.resolveFieldByReference(fieldsByKey, condition.field);
      if (!field) {
        continue;
      }

      var resolvedValue = this.resolvePresetFilterValue(condition);

      var matches = this.rowMatchesFilter(row, field, {
        operator: normalizeFilterOperator(condition.operator),
        value: resolvedValue.value,
        compareDateOnly: resolvedValue.compareDateOnly
      });

      if (aggregate === null) {
        aggregate = matches;
      } else if (normalizeFilterLogical(condition.logical) === 'or') {
        aggregate = aggregate || matches;
      } else {
        aggregate = aggregate && matches;
      }
    }

    return aggregate === null ? true : aggregate;
  }

  private resolvePresetFilterValue(condition: { value: any; valueType?: string }): { value: any; compareDateOnly: boolean } {
    var rawValue: any = Array.isArray(condition.value) ? condition.value : (condition.value === undefined || condition.value === null ? '' : String(condition.value).trim());
    if (String(condition.valueType || '').toLowerCase() !== 'expression') {
      return { value: rawValue, compareDateOnly: false };
    }

    var expression = rawValue.toLowerCase();
    var pageContext = this.props.context && this.props.context.pageContext;
    var user = pageContext && pageContext.user;
    var legacyContext = pageContext && (pageContext as any).legacyPageContext;

    if (expression === 'today') {
      return { value: formatLocalDate(new Date()), compareDateOnly: true };
    }
    if (expression === 'now') {
      return { value: new Date().toISOString(), compareDateOnly: false };
    }

    var dateMatch = /^date\(\s*([+-]?\d+)\s*\)$/.exec(expression);
    if (dateMatch) {
      var dateValue = new Date();
      dateValue.setDate(dateValue.getDate() + parseInt(dateMatch[1], 10));
      return { value: formatLocalDate(dateValue), compareDateOnly: true };
    }

    if (expression === 'me.email') {
      return { value: String(user && user.email || ''), compareDateOnly: false };
    }
    if (expression === 'me.login') {
      return { value: String(user && user.loginName || ''), compareDateOnly: false };
    }
    if (expression === 'me.id') {
      return { value: String(legacyContext && legacyContext.userId || ''), compareDateOnly: false };
    }
    if (expression === 'me') {
      return { value: String(user && (user.displayName || user.loginName) || ''), compareDateOnly: false };
    }

    return { value: rawValue, compareDateOnly: false };
  }

  private getProcessedRows(): any[] {
    var rows = this.state.rows.slice(0);
    var fieldsByKey: { [key: string]: IListFieldDefinition } = {};
    for (var i = 0; i < this.state.fields.length; i += 1) {
      var field = this.state.fields[i];
      var fieldKey = this.getFieldKey(field);
      if (fieldKey) {
        fieldsByKey[fieldKey] = field;
      }
    }

    var presetConditions = this.parsePresetFilterConditions();
    if (presetConditions.length > 0) {
      rows = rows.filter((row: any) => {
        return this.rowMatchesPresetConditions(row, fieldsByKey, presetConditions);
      });
    }

    var filterKeys = Object.keys(this.state.columnFilters || {});
    if (filterKeys.length > 0) {
      rows = rows.filter((row: any) => {
        for (var index = 0; index < filterKeys.length; index += 1) {
          var key = filterKeys[index];
          var filter = this.state.columnFilters[key];
          var field = fieldsByKey[key];
          if (!filter || !field) {
            continue;
          }
          if (!this.rowMatchesFilter(row, field, filter)) {
            return false;
          }
        }
        return true;
      });
    }

    if (this.state.sortFieldName && this.state.sortDirection) {
      var sortField = fieldsByKey[this.state.sortFieldName];
      if (sortField) {
        rows.sort((leftRow: any, rightRow: any) => {
          var left = this.getCellPlainText(leftRow, sortField);
          var right = this.getCellPlainText(rightRow, sortField);
          var comparison = this.compareComparableValues(left, right);
          return this.state.sortDirection === 'asc' ? comparison : (comparison * -1);
        });
      }
    }

    return rows;
  }

  private parseConditionalStyleRules(): IConditionalStyleRule[] {
    var source = String(this.props.conditionalStyleJson || '').trim();
    if (!source) {
      return [];
    }

    try {
      var parsed = JSON.parse(source);
      if (!Array.isArray(parsed)) {
        return [];
      }

      var groupedRules: { [ruleKey: string]: IConditionalStyleRule } = {};
      var orderedRuleKeys: string[] = [];

      for (var i = 0; i < parsed.length; i += 1) {
        var item = parsed[i] || {};
        var ruleKey = toCssString(item.rule || item.ruleName || 'rule-' + String(i + 1));
        if (!ruleKey) {
          ruleKey = 'rule-' + String(i + 1);
        }

        var scope: ConditionalScope = String(item.scope || 'row').trim().toLowerCase() === 'column' ? 'column' : 'row';
        var applyField = toCssString(item.applyField);
        var conditionField = toCssString(item.conditionField || item.field);
        if (!conditionField) {
          continue;
        }

        var conditionOperator = normalizeFilterOperator(item.operator || item.conditionOperator);
        var conditionLogical = normalizeFilterLogical(item.logical || item.conditionLogical) as 'and' | 'or';
        var conditionValue = toCssString(item.value || item.conditionValue);

        var styleDefinition: IConditionalStyleDefinition = {
          backgroundColor: toCssString(item.backgroundColor || (item.style && item.style.backgroundColor) || ''),
          color: toCssString(item.color || item.foregroundColor || (item.style && item.style.color) || ''),
          fontFamily: toCssString(item.fontFamily || (item.style && item.style.fontFamily) || ''),
          fontSize: toCssString(item.fontSize || (item.style && item.style.fontSize) || ''),
          fontStyle: toCssString(item.fontStyle || (item.style && item.style.fontStyle) || ''),
          fontWeight: toCssString(item.fontWeight || (item.style && item.style.fontWeight) || ''),
          textAlign: toCssString(item.textAlign || (item.style && item.style.textAlign) || '')
        };

        if (!groupedRules[ruleKey]) {
          groupedRules[ruleKey] = {
            key: ruleKey,
            scope: scope,
            applyField: applyField,
            enabled: item.enabled !== false,
            priority: toPriorityNumber(item.priority, 100),
            sourceOrder: i,
            style: styleDefinition,
            conditions: []
          };
          orderedRuleKeys.push(ruleKey);
        }

        groupedRules[ruleKey].scope = scope;
        groupedRules[ruleKey].applyField = applyField;
        groupedRules[ruleKey].enabled = item.enabled !== false;
        groupedRules[ruleKey].priority = toPriorityNumber(item.priority, groupedRules[ruleKey].priority);
        groupedRules[ruleKey].style = styleDefinition;
        groupedRules[ruleKey].conditions.push({
          field: conditionField,
          operator: conditionOperator,
          logical: conditionLogical,
          valueType: String(item.valueType || '').toLowerCase() === 'expression' ? 'expression' : 'static',
          value: conditionValue
        });
      }

      return orderedRuleKeys
        .map((ruleKey) => groupedRules[ruleKey])
        .filter((rule) => !!rule && rule.conditions.length > 0)
        .sort(function(leftRule: IConditionalStyleRule, rightRule: IConditionalStyleRule) {
          if (leftRule.priority !== rightRule.priority) {
            return leftRule.priority - rightRule.priority;
          }

          return leftRule.sourceOrder - rightRule.sourceOrder;
        });
    } catch (_parseError) {
      this.logDiagnostic('Conditional style JSON is invalid; skipping conditional styles.');
      return [];
    }
  }

  private toReactCssStyle(styleDefinition: IConditionalStyleDefinition): React.CSSProperties {
    var style: React.CSSProperties = {};
    var backgroundColor = toCssString(styleDefinition.backgroundColor);
    var color = toCssString(styleDefinition.color);
    var fontFamily = toCssString(styleDefinition.fontFamily);
    var fontSize = toCssString(styleDefinition.fontSize);
    var fontStyle = toCssString(styleDefinition.fontStyle);
    var fontWeight = toCssString(styleDefinition.fontWeight);
    var textAlign = toCssString(styleDefinition.textAlign);

    if (backgroundColor) {
      style.backgroundColor = backgroundColor;
    }
    if (color) {
      style.color = color;
    }
    if (fontFamily) {
      style.fontFamily = fontFamily;
    }
    if (fontSize) {
      style.fontSize = fontSize;
    }
    if (fontStyle) {
      style.fontStyle = fontStyle as any;
    }
    if (fontWeight) {
      style.fontWeight = fontWeight as any;
    }
    if (textAlign) {
      style.textAlign = textAlign as any;
    }

    return style;
  }

  private evaluateConditionalStyleRule(row: any, rule: IConditionalStyleRule, fieldsByKey: { [key: string]: IListFieldDefinition }): boolean {
    if (!rule || !rule.conditions || rule.conditions.length === 0) {
      return false;
    }

    var aggregate: boolean | null = null;
    for (var i = 0; i < rule.conditions.length; i += 1) {
      var condition = rule.conditions[i];
      var field = this.resolveFieldByReference(fieldsByKey, condition.field);
      if (!field) {
        continue;
      }

      var resolvedValue = this.resolvePresetFilterValue(condition);

      var matches = this.rowMatchesFilter(row, field, {
        operator: condition.operator,
        value: resolvedValue.value,
        compareDateOnly: resolvedValue.compareDateOnly
      });

      if (aggregate === null) {
        aggregate = matches;
      } else if (condition.logical === 'or') {
        aggregate = aggregate || matches;
      } else {
        aggregate = aggregate && matches;
      }
    }

    return aggregate === null ? false : aggregate;
  }

  private getConditionalStyleForRow(row: any, fieldsByKey: { [key: string]: IListFieldDefinition }, rules: IConditionalStyleRule[]): IResolvedConditionalStyle {
    var rowStyle: React.CSSProperties = {};
    var columnStylesByFieldKey: { [fieldKey: string]: React.CSSProperties } = {};

    for (var i = 0; i < rules.length; i += 1) {
      var rule = rules[i];
      if (!rule.enabled) {
        continue;
      }

      if (!this.evaluateConditionalStyleRule(row, rule, fieldsByKey)) {
        continue;
      }

      var overrideStyle = this.toReactCssStyle(rule.style);
      if (Object.keys(overrideStyle).length === 0) {
        continue;
      }

      if (rule.scope === 'column') {
        var applyFieldDefinition = this.resolveFieldByReference(fieldsByKey, rule.applyField);
        if (!applyFieldDefinition) {
          continue;
        }

        var applyFieldKey = this.getFieldKey(applyFieldDefinition);
        if (!applyFieldKey) {
          continue;
        }

        var columnStyle = columnStylesByFieldKey[applyFieldKey] || {};
        columnStylesByFieldKey[applyFieldKey] = mergeStyleObjects(columnStyle, overrideStyle);
      } else {
        rowStyle = mergeStyleObjects(rowStyle, overrideStyle);
      }
    }

    return {
      rowStyle: rowStyle,
      columnStylesByFieldKey: columnStylesByFieldKey
    };
  }

  private getMatchedConditionalRuleNames(row: any, fieldsByKey: { [key: string]: IListFieldDefinition }, rules: IConditionalStyleRule[]): string[] {
    var names: string[] = [];
    for (var i = 0; i < rules.length; i += 1) {
      var rule = rules[i];
      if (!rule.enabled) {
        continue;
      }

      if (this.evaluateConditionalStyleRule(row, rule, fieldsByKey)) {
        names.push(rule.key);
      }
    }

    return names;
  }

  public render(): JSX.Element {
    if (!this.props.listName) {
      return <div>{strings.RuntimeNoListSelected}</div>;
    }

    var canOperateOnSelection = this.state.selectedItemId > 0;
    var processedRows = this.getProcessedRows();
    var pageSize = this.props.pageSize > 0 ? this.props.pageSize : 0;
    var pageCount = pageSize > 0 ? Math.max(1, Math.ceil(processedRows.length / pageSize)) : 1;
    var currentPage = Math.min(this.state.currentPage, pageCount - 1);
    var visibleRows = pageSize > 0
      ? processedRows.slice(currentPage * pageSize, (currentPage + 1) * pageSize)
      : processedRows;
    var conditionalRules = this.parseConditionalStyleRules();
    var hasActiveFilters = Object.keys(this.state.columnFilters || {}).length > 0;
    var displayFields = this.getDisplayFields();
    var fieldsByKey: { [key: string]: IListFieldDefinition } = {};
    for (var i = 0; i < displayFields.length; i += 1) {
      var fieldsByKeyField = displayFields[i];
      var fieldsByKeyValue = this.getFieldKey(fieldsByKeyField);
      if (fieldsByKeyValue) {
        fieldsByKey[fieldsByKeyValue] = fieldsByKeyField;
      }
    }
    var selectedRow: any = null;
    var selectedMatchedRuleNames: string[] = [];
    if (this.props.isEditMode && this.state.selectedItemId > 0) {
      for (var selectedIndex = 0; selectedIndex < processedRows.length; selectedIndex += 1) {
        var selectedCandidate = processedRows[selectedIndex];
        if (this.getRowItemId(selectedCandidate) === this.state.selectedItemId) {
          selectedRow = selectedCandidate;
          break;
        }
      }

      if (selectedRow) {
        selectedMatchedRuleNames = this.getMatchedConditionalRuleNames(selectedRow, fieldsByKey, conditionalRules);
      }
    }
    var containerStyle: any = {
      '--lc-body-text-color': this.props.bodyTextColor || 'inherit',
      '--lc-body-font-family': this.props.bodyFontFamily || 'inherit',
      '--lc-body-font-size': this.props.bodyFontSize || '14px',
      '--lc-body-font-style': this.props.bodyFontStyle || 'normal',
      '--lc-body-font-weight': this.props.bodyFontBold ? 'bold' : 'normal',
      '--lc-body-text-align': this.props.bodyTextAlign || 'left',
      '--lc-selected-text-color': this.props.selectedTextColor || this.props.bodyTextColor || 'inherit',
      '--lc-selected-bg-color': this.props.selectedBackgroundColor || '#eef6ff',
      '--lc-selected-font-style': this.props.selectedFontStyle || this.props.bodyFontStyle || 'normal',
      '--lc-selected-font-weight': this.props.selectedFontBold ? 'bold' : (this.props.bodyFontBold ? 'bold' : 'normal'),
      '--lc-header-text-color': this.props.headerTextColor || 'inherit',
      '--lc-header-bg-color': this.props.headerBackgroundColor || 'transparent',
      '--lc-header-font-family': this.props.headerFontFamily || 'inherit',
      '--lc-header-font-size': this.props.headerFontSize || '14px',
      '--lc-header-font-style': this.props.headerFontStyle || 'normal',
      '--lc-header-font-weight': this.props.headerFontBold ? 'bold' : 'normal',
      '--lc-header-text-align': this.props.headerTextAlign || 'left',
      '--lc-table-bg-color': this.props.tableBackgroundColor || '#ffffff',
      '--lc-table-border-color': this.props.tableBorderColor || '#ccc',
      '--lc-table-border-width': this.props.tableBorderWidth || '1px',
      '--lc-table-corner-radius': this.props.tableCornerStyle === 'rounded'
        ? (typeof this.props.tableCornerRadius === 'number' ? this.props.tableCornerRadius : 8) + 'px'
        : '0px',
      '--lc-row-line-width': (typeof this.props.tableRowLineWidth === 'number' ? this.props.tableRowLineWidth : 1) + 'px',
      '--lc-alternate-row-color': this.props.alternateRowShading
        ? (this.props.alternateRowShadingColor || '#f9f9f9')
        : 'transparent',
      '--lc-button-text-color': this.props.buttonTextColor || '#000000',
      '--lc-button-bg-color': this.props.buttonBackgroundColor || '#f0f0f0',
      '--lc-button-font-family': this.props.buttonFontFamily || 'inherit',
      '--lc-button-font-size': this.props.buttonFontSize || '14px',
      '--lc-button-font-style': this.props.buttonFontStyle || 'normal',
      '--lc-button-font-weight': this.props.buttonFontBold ? 'bold' : 'normal',
      '--lc-button-corner-radius': this.props.buttonCornerStyle === 'rounded'
        ? (typeof this.props.buttonCornerRadius === 'number' ? this.props.buttonCornerRadius : 4) + 'px'
        : '0px',
      '--lc-webpart-bg-color': this.props.webpartBackgroundColor || '#ffffff',
      '--lc-webpart-border-color': this.props.webpartBorderColor || '#ccc',
      '--lc-webpart-border-width': (this.props.webpartBorderWidth || 0) + 'px',
      '--lc-webpart-corner-style': this.props.cornerStyle === 'rounded'
        ? (typeof this.props.cornerRadius === 'number' ? this.props.cornerRadius : 12) + 'px'
        : '0px'
    };

    return (
      <div className="lc-root" style={containerStyle}>
        {(this.state.displayFormUrl || this.state.displayFormLoading || this.state.displayFormError)
          && <div className="lc-form-dialog-backdrop" role="presentation" onClick={() => this.closeDefaultDisplayForm()}>
            <section className="lc-form-dialog" role="dialog" aria-modal="true" aria-label="Item details"
              onClick={(event) => event.stopPropagation()}>
              <button type="button" className="lc-form-dialog-close" aria-label="Close item details"
                title="Close item details" onClick={() => this.closeDefaultDisplayForm()}>&times;</button>
              {this.state.displayFormLoading && <div className="lc-form-dialog-message">Loading item...</div>}
              {this.state.displayFormError && <div className="lc-form-dialog-error">{this.state.displayFormError}</div>}
              {this.state.displayFormUrl && <iframe ref={this.setDisplayFormFrameRef} className="lc-form-dialog-frame"
                src={this.state.displayFormUrl} title="Item details" />}
            </section>
          </div>}
        <div className="lc-toolbar">
          {this.props.showViewSelector && !this.state.embeddedByReportForms && (
            <label>
              <span style={{ marginRight: '6px' }}>{strings.RuntimeViewLabel}</span>
              <select
                value={this.state.selectedViewId}
                onChange={(ev) => this.setState({ selectedViewId: ev.currentTarget.value })}
              >
                {this.props.views.map((view) => {
                  return <option key={view.key} value={view.key}>{view.text}</option>;
                })}
              </select>
            </label>
          )}
          {this.props.showRefresh && !this.state.embeddedByReportForms && <button type="button" onClick={() => this.loadRows()}>{strings.RuntimeRefresh}</button>}
          {this.props.showAdd && !this.state.embeddedByReportForms && (
            <button
              type="button"
              onClick={() => {
                this.setState({ selectedItemId: 0, selectedMode: 'new' });
                this.props.onSelectionChange(0, 'new');
              }}
            >
              {strings.RuntimeNew}
            </button>
          )}
          {this.props.showEdit && !this.state.embeddedByReportForms && (
            <button
              type="button"
              disabled={!canOperateOnSelection}
              onClick={() => {
                this.setState({ selectedMode: 'edit' });
                this.props.onSelectionChange(this.state.selectedItemId, 'edit');
              }}
            >
              {strings.RuntimeEdit}
            </button>
          )}
          {this.props.showView && !this.state.embeddedByReportForms && (
            <button
              type="button"
              disabled={!canOperateOnSelection}
              onClick={() => {
                this.setState({ selectedMode: 'view' });
                this.props.onSelectionChange(this.state.selectedItemId, 'view');
              }}
            >
              {strings.RuntimeView}
            </button>
          )}
          {this.props.showDelete && !this.state.embeddedByReportForms && (
            <button
              type="button"
              disabled={!canOperateOnSelection || this.state.deleting}
              onClick={() => this.deleteSelected()}
            >
              {strings.RuntimeDelete}
            </button>
          )}
        </div>

        {this.props.isEditMode && (
          <div className="lc-status">
            <div>{formatString(strings.RuntimeSelectedItem, this.state.selectedItemId || 0)}</div>
            <div>{formatString(strings.RuntimeSelectedMode, this.state.selectedMode)}</div>
            <div>{formatString(strings.RuntimeConditionalStyleMatches, selectedMatchedRuleNames.length > 0 ? selectedMatchedRuleNames.join(', ') : strings.RuntimeConditionalStyleMatchesNone)}</div>
            {!canOperateOnSelection && <div>{strings.RuntimeSelectRowPrompt}</div>}
          </div>
        )}

        {this.state.loading && <div>{strings.RuntimeLoading}</div>}
        {!!this.state.error && <div>{this.state.error}</div>}
        {!this.state.loading && !this.state.error && processedRows.length === 0 && (
          <div>{hasActiveFilters ? strings.RuntimeNoItemsAfterFilter : strings.RuntimeNoItems}</div>
        )}

        {!this.state.loading && !this.state.error && processedRows.length > 0 && (
          <div className="lc-table-wrap">
            <table className="lc-table">
              <thead>
                <tr>
                  {displayFields.map((field) => {
                    var fieldKey = this.getFieldKey(field);
                    var sortActive = this.state.sortFieldName === fieldKey;
                    var filterActive = !!this.state.columnFilters[fieldKey];
                    var sortIndicator = '';
                    if (sortActive && this.state.sortDirection === 'asc') {
                      sortIndicator = ' ▲';
                    } else if (sortActive && this.state.sortDirection === 'desc') {
                      sortIndicator = ' ▼';
                    }

                    return (
                      <th key={field.Name} style={this.getConfiguredColumnStyle(field)}>
                        <div className="lc-header-cell">
                          <button
                            type="button"
                            className="lc-header-title"
                            onClick={(ev) => {
                              ev.stopPropagation();
                              this.toggleSort(field);
                            }}
                            title={strings.RuntimeSortToggle}
                          >
                            {field.DisplayName || field.Name || strings.ColumnIdFallback}{sortIndicator}
                          </button>
                          <div className="lc-header-actions">
                            <button
                              type="button"
                              className={joinClassNames(['lc-header-action', filterActive ? 'lc-header-action-active' : ''])}
                              title={strings.RuntimeFilterTitle}
                              onClick={(ev) => {
                                ev.preventDefault();
                                ev.stopPropagation();
                                if (this.state.activeFilterFieldName === fieldKey) {
                                  this.closeFilter();
                                } else {
                                  this.openFilter(field, ev.currentTarget);
                                }
                              }}
                            >
                              {strings.RuntimeFilterIcon}
                            </button>
                          </div>
                          {this.state.activeFilterFieldName === fieldKey && (
                            <div
                              className="lc-filter-popover"
                              style={this.state.filterPopoverStyle}
                              onClick={(ev) => {
                                ev.preventDefault();
                                ev.stopPropagation();
                              }}
                            >
                              {this.state.activeFilterIsDate ? (
                                <div>
                                  <label className="lc-filter-label">{strings.RuntimeFilterDateLabel}</label>
                                  <div className="lc-date-input-row">
                                    <input className="lc-filter-input" type="text" placeholder="YYYY-MM-DD" value={this.state.draftFilterValue} onChange={(ev) => this.setState({ draftFilterValue: ev.currentTarget.value })} />
                                    <button type="button" className="lc-date-picker-button" title="Choose from date" aria-label="Choose from date" onClick={() => this.toggleDatePicker('start')}><span aria-hidden="true">&#128197;</span></button>
                                  </div>
                                  <label className="lc-filter-label">{strings.RuntimeFilterEndDateLabel}</label>
                                  <div className="lc-date-input-row">
                                    <input className="lc-filter-input" type="text" placeholder="YYYY-MM-DD" value={this.state.draftFilterEndValue} onChange={(ev) => this.setState({ draftFilterEndValue: ev.currentTarget.value })} />
                                    <button type="button" className="lc-date-picker-button" title="Choose to date" aria-label="Choose to date" onClick={() => this.toggleDatePicker('end')}><span aria-hidden="true">&#128197;</span></button>
                                  </div>
                                  {this.renderDatePicker()}
                                </div>
                              ) : (
                                <div>
                                  <label className="lc-filter-label">{strings.RuntimeFilterOperatorLabel}</label>
                                  <select className="lc-filter-select" value={this.state.draftFilterOperator} onChange={(ev) => this.setState({ draftFilterOperator: ev.currentTarget.value as FilterOperator })}>
                                    {this.getFilterOperatorOptions().map((option) => {
                                      return <option key={option.key} value={option.key}>{option.label}</option>;
                                    })}
                                  </select>
                                  <label className="lc-filter-label">{strings.RuntimeFilterValueLabel}</label>
                                  <input className="lc-filter-input" type="text" value={this.state.draftFilterValue} placeholder={strings.RuntimeFilterValuePlaceholder} onChange={(ev) => this.setState({ draftFilterValue: ev.currentTarget.value })} onKeyDown={(ev) => {
                                    if (ev.key === 'Enter') {
                                      this.applyActiveFilter();
                                    }
                                  }} />
                                </div>
                              )}
                              <div className="lc-filter-actions">
                                <button type="button" onClick={() => this.applyActiveFilter()}>{strings.RuntimeFilterApply}</button>
                                <button type="button" onClick={() => this.clearActiveFilter()}>{strings.RuntimeFilterClear}</button>
                                <button type="button" onClick={() => this.closeFilter()}>{strings.RuntimeFilterClose}</button>
                              </div>
                            </div>
                          )}
                        </div>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {visibleRows.map((row, index) => {
                  var rowItemId = this.getRowItemId(row);
                  var isSelected = rowItemId > 0 && rowItemId === this.state.selectedItemId;
                  var conditionalStyle = this.getConditionalStyleForRow(row, fieldsByKey, conditionalRules);
                  return (
                    <tr
                      key={rowItemId > 0 ? String(rowItemId) : String(index)}
                      onClick={() => this.selectRow(row)}
                      className={joinClassNames(['lc-row', isSelected ? 'lc-row-selected' : ''])}
                    >
                      {displayFields.map((field) => {
                        var markup = this.getCellMarkup(row, field);
                        var urlCell = this.getUrlCellValue(row, field);
                        var showItemLink = this.props.showLinkToItem && this.isTitleField(field);
                        var itemLinkUrl = showItemLink ? this.getItemLinkUrl(row) : '';
                        var itemLinkText = showItemLink ? this.getCellPlainText(row, field) : '';
                        var cellFieldKey = this.getFieldKey(field);
                        var columnStyle = conditionalStyle.columnStylesByFieldKey[cellFieldKey] || {};
                        var mergedCellStyle = mergeStyleObjects(
                          mergeStyleObjects(conditionalStyle.rowStyle, columnStyle),
                          this.getConfiguredColumnStyle(field)
                        );
                        return (
                          <td key={field.Name} style={mergedCellStyle}>
                            {urlCell ? (
                              <a
                                className="lc-item-link"
                                href={urlCell.href}
                                target="_blank"
                                rel="noopener noreferrer"
                                onClick={(ev) => ev.stopPropagation()}
                              >
                                {urlCell.text}
                              </a>
                            ) : showItemLink && itemLinkUrl ? (
                              <a
                                className="lc-item-link"
                                href={itemLinkUrl}
                                onClick={(ev) => {
                                  ev.stopPropagation();
                                  if (!String(this.props.linkTargetPageUrl || '').trim()) {
                                    ev.preventDefault();
                                    this.openDefaultDisplayForm(row);
                                  }
                                }}
                              >
                                {itemLinkText || strings.RuntimeView}
                              </a>
                            ) : (
                              markup ? <span dangerouslySetInnerHTML={markup} /> : null
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
        {!this.state.loading && !this.state.error && pageSize > 0 && pageCount > 1 && (
          <div className="lc-pagination">
            <button
              type="button"
              disabled={currentPage === 0}
              onClick={() => this.setState({ currentPage: Math.max(0, currentPage - 1) })}
            >
              {strings.RuntimePreviousPage}
            </button>
            <span>{strings.RuntimePageStatus.replace('{0}', String(currentPage + 1)).replace('{1}', String(pageCount))}</span>
            <button
              type="button"
              disabled={currentPage >= pageCount - 1}
              onClick={() => this.setState({ currentPage: Math.min(pageCount - 1, currentPage + 1) })}
            >
              {strings.RuntimeNextPage}
            </button>
          </div>
        )}
      </div>
    );
  }
}