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
  draftFilterOperator: FilterOperator;
  draftFilterValue: string;
  columnFilters: { [fieldName: string]: IColumnFilter };
  currentPage: number;
}

export type FilterOperator = 'eq' | 'ne' | 'contains' | 'notcontains' | 'startswith' | 'endswith' | 'gt' | 'ge' | 'lt' | 'le';

export interface IColumnFilter {
  operator: FilterOperator;
  value: string;
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

function appendQuery(url: string, query: string): string {
  return url + (url.indexOf('?') >= 0 ? '&' : '?') + query;
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

function buildViewRequestUrls(baseEndpoint: string, selectedViewId: string): string[] {
  var urls: string[] = [];
  var normalized = trimGuidBraces(selectedViewId);
  var candidates = [String(selectedViewId || ''), normalized, '{' + normalized + '}'];

  for (var i = 0; i < candidates.length; i += 1) {
    var candidate = String(candidates[i] || '');
    if (!candidate) {
      continue;
    }

    urls.push(appendQuery(baseEndpoint, 'View=' + encodeURIComponent(candidate)));
    urls.push(appendQuery(baseEndpoint, 'ViewId=' + encodeURIComponent(candidate)));
  }

  urls.push(baseEndpoint);

  var unique: string[] = [];
  for (var j = 0; j < urls.length; j += 1) {
    if (unique.indexOf(urls[j]) < 0) {
      unique.push(urls[j]);
    }
  }

  return unique;
}

function extractRenderRowsAndFields(data: any): { rows: any[]; fields: IListFieldDefinition[] } {
  var schema = tryParseObject(data.ListSchema) || tryParseObject(data.Schema) || {};
  var listData = tryParseObject(data.ListData) || {};
  var rows = toArray(data.Row);

  if (rows.length === 0) {
    rows = toArray(data.Rows);
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
    fields = toArray(data.Field);
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
      draftFilterOperator: 'contains',
      draftFilterValue: '',
      columnFilters: {},
      currentPage: 0,
    };

    this._refreshEventHandler = this.handleExternalRefresh.bind(this);
  }

  public componentDidMount(): void {
    this.logDiagnostic('Component mounted. listName=' + String(this.props.listName || '(none)') + ', defaultViewId=' + String(this.props.defaultViewId || '(none)'));
    if (typeof window !== 'undefined' && window.addEventListener) {
      window.addEventListener(LIST_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
    }
    this.loadRows();
  }

  public componentWillUnmount(): void {
    if (typeof window !== 'undefined' && window.removeEventListener) {
      window.removeEventListener(LIST_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
    }
  }

  public componentDidUpdate(prevProps: IListControlProps, prevState: IListControlState): void {
    if (prevProps.listName !== this.props.listName || prevProps.defaultViewId !== this.props.defaultViewId) {
      this.logDiagnostic('Props changed; resetting selection and reloading rows. listName=' + String(this.props.listName || '(none)') + ', viewId=' + String(this.props.defaultViewId || '(none)'));
      this.setState({
        selectedViewId: this.props.defaultViewId || this.getInitialViewId(this.props.views),
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
        this.loadRows();
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
      var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=InternalName,TypeAsString";
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

  private getRowFieldValue(row: any, field: IListFieldDefinition): any {
    var fieldName = field.Name || field.RealFieldName || '';
    var value = row[fieldName];
    if ((value === undefined || value === null || value === '') && field.RealFieldName) {
      value = row[field.RealFieldName];
    }
    return value;
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

      // Keep rows only when at least one visible field has meaningful content.
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

    this.logDiagnostic('Starting loadRows. listName=' + String(this.props.listName) + ', selectedViewId=' + String(this.state.selectedViewId || '(none)'));
    this.setState({ loading: true, error: null });

    try {
      var baseEndpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/RenderListDataAsStream";
      var selectedViewId = this.state.selectedViewId;

      var body: any = {
        parameters: {
          RenderOptions: 17
        }
      };

      var requestUrls = selectedViewId ? buildViewRequestUrls(baseEndpoint, selectedViewId) : [baseEndpoint];
      var selectedRows: any[] = [];
      var selectedFields: IListFieldDefinition[] = [];
      var lastError: string = '';
      var hadSuccessfulResponse = false;

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
      var viewFieldNames = await this.loadSelectedViewFieldNames(selectedViewId);

      if (rows.length === 0) {
        var itemsFallback = await this.loadRowsFromItemsEndpoint(viewFieldNames);
        rows = itemsFallback.rows;

        if (fields.length === 0) {
          fields = itemsFallback.fields;
        }
      }

      var visibleFields = this.getFieldsForConsumption(fields, viewFieldNames);
      var fieldTitleMap = await this.loadListFieldTitleMap();
      visibleFields = this.applyFieldDisplayNames(visibleFields, fieldTitleMap);
      var renderableRows = this.filterRenderableRows(rows, visibleFields);

      this.setState({
        fields: visibleFields,
        rows: renderableRows,
        loading: false,
        error: null
      });
      this.logDiagnostic('loadRows completed. visibleFields=' + String(visibleFields.length) + ', renderableRows=' + String(renderableRows.length));
    } catch (error) {
      var loadError: any = error as any;
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

    var text = this.stringifyCellValue(value);
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

  private openFilter(field: IListFieldDefinition): void {
    var fieldKey = this.getFieldKey(field);
    if (!fieldKey) {
      return;
    }

    var existing = this.state.columnFilters[fieldKey];
    this.setState({
      activeFilterFieldName: fieldKey,
      draftFilterOperator: existing ? existing.operator : 'contains',
      draftFilterValue: existing ? existing.value : ''
    });
  }

  private closeFilter(): void {
    this.setState({ activeFilterFieldName: '' });
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
    if (!value) {
      delete nextFilters[fieldKey];
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
    var valueText = this.getCellPlainText(row, field);
    var candidate = String(valueText || '').trim();
    var query = String(filter.value || '').trim();

    if (!query) {
      return true;
    }

    var normalizedCandidate = candidate.toLowerCase();
    var normalizedQuery = query.toLowerCase();
    var compareResult = this.compareComparableValues(candidate, query);

    if (filter.compareDateOnly) {
      var candidateDate = new Date(candidate);
      if (!isNaN(candidateDate.getTime())) {
        normalizedCandidate = formatLocalDate(candidateDate).toLowerCase();
        compareResult = this.compareComparableValues(normalizedCandidate, query);
      }
    }

    switch (filter.operator) {
      case 'eq':
        return normalizedCandidate === normalizedQuery;
      case 'ne':
        return normalizedCandidate !== normalizedQuery;
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

  private parsePresetFilterConditions(): IPresetFilterCondition[] {
    var source = String(this.props.filterJson || '').trim();
    if (!source) {
      return [];
    }

    try {
      var parsed = JSON.parse(source);
      if (!Array.isArray(parsed)) {
        return [];
      }

      var conditions: IPresetFilterCondition[] = [];
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
          valueType: String(item.valueType || '').toLowerCase() === 'expression' ? 'expression' : 'static',
          value: item.value
        });
      }

      return conditions;
    } catch (_parseError) {
      this.logDiagnostic('Preset filter JSON is invalid; skipping preset filters.');
      return [];
    }
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

  private resolvePresetFilterValue(condition: { value: any; valueType?: string }): { value: string; compareDateOnly: boolean } {
    var rawValue = condition.value === undefined || condition.value === null ? '' : String(condition.value).trim();
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
        <div className="lc-toolbar">
          {this.props.showViewSelector && (
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
          {this.props.showRefresh && <button type="button" onClick={() => this.loadRows()}>{strings.RuntimeRefresh}</button>}
          {this.props.showAdd && (
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
          {this.props.showEdit && (
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
          {this.props.showView && (
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
          {this.props.showDelete && (
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
                                  this.openFilter(field);
                                }
                              }}
                            >
                              {strings.RuntimeFilterIcon}
                            </button>
                          </div>
                          {this.state.activeFilterFieldName === fieldKey && (
                            <div
                              className="lc-filter-popover"
                              onClick={(ev) => {
                                ev.preventDefault();
                                ev.stopPropagation();
                              }}
                            >
                              <label className="lc-filter-label">{strings.RuntimeFilterOperatorLabel}</label>
                              <select
                                className="lc-filter-select"
                                value={this.state.draftFilterOperator}
                                onChange={(ev) => this.setState({ draftFilterOperator: ev.currentTarget.value as FilterOperator })}
                              >
                                {this.getFilterOperatorOptions().map((option) => {
                                  return <option key={option.key} value={option.key}>{option.label}</option>;
                                })}
                              </select>
                              <label className="lc-filter-label">{strings.RuntimeFilterValueLabel}</label>
                              <input
                                className="lc-filter-input"
                                type="text"
                                value={this.state.draftFilterValue}
                                placeholder={strings.RuntimeFilterValuePlaceholder}
                                onChange={(ev) => this.setState({ draftFilterValue: ev.currentTarget.value })}
                                onKeyDown={(ev) => {
                                  if (ev.key === 'Enter') {
                                    this.applyActiveFilter();
                                  }
                                }}
                              />
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
                            {showItemLink && itemLinkUrl ? (
                              <a
                                className="lc-item-link"
                                href={itemLinkUrl}
                                onClick={(ev) => ev.stopPropagation()}
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