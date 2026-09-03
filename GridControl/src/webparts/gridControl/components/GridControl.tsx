import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import * as strings from 'GridControlWebPartStrings';
import { evaluateGridValidationExpression, IGridAdvancedValidationRule, IGridValidationField } from './GridValidation';
import { GridRichTextEditor } from './GridRichTextEditor';
import './GridControl.css';

export interface IGridControlViewOption {
  key: string;
  text: string;
  isDefault?: boolean;
}

export interface IGridControlColumnConfiguration {
  fieldName: string;
  displayName: string;
  width?: string;
}

export interface IGridControlProps {
  context: any;
  listName: string;
  defaultViewId: string;
  views: IGridControlViewOption[];
  viewColumns: IGridControlColumnConfiguration[];
  gridSchemaJson: string;
  pageSize: number;
  fetchBatchSize: number;
  isEditMode: boolean;
  showViewSelector: boolean;
  showViewAsDropdown: boolean;
  showRefresh: boolean;
  showAdd: boolean;
  showDelete: boolean;
  showHistory: boolean;
  historyAvailable: boolean;
  actionButtonsPosition: string;
  buttonDisplayMode: string;
  readSecurityGroupId: string;
  editSecurityGroupId: string;
  readAccessDeniedMessage: string;
  editAccessDeniedMessage: string;
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
  dateCustomFormat?: string;
  dateCustomFormatCase?: string;
  timeDisplayFormat: string;
  timeCustomFormat?: string;
  timeCustomFormatCase?: string;
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

export interface IGridFieldMetadata {
  internalName: string;
  title: string;
  typeAsString: string;
  required: boolean;
  readOnly: boolean;
  hidden: boolean;
  description: string;
  choices: string[];
  displayFormat: number;
  lookupList: string;
  lookupField: string;
  allowMultiple: boolean;
  richText: boolean;
}

export interface IGridLookupOption {
  id: number;
  text: string;
}

interface IGridSchemaField {
  id?: string;
  fieldName: string;
  sharePointType?: string;
  type: string;
  label?: string;
  description?: string;
  visible?: any;
  required?: any;
  requiredMessage?: string;
  readOnly?: any;
  disabled?: boolean;
  defaultValue?: any;
  config?: any;
  validation?: any[];
  gridWidth?: string;
}

interface IGridGroupingConfig {
  enabled: boolean;
  field1: string;
  field2: string;
  collapsedByDefault: boolean;
  showCount: boolean;
}

interface IGridRowGroup {
  key: string;
  label: string;
  rows: any[];
  subGroups: IGridRowGroup[];
}

function isCompatibleGridControlType(sharePointType: string, controlType: string): boolean {
  var compatibleTypes: { [sharePointType: string]: string[] } = {
    Text: ['text', 'multiline', 'number'],
    Note: ['text', 'multiline', 'number'],
    Number: ['number'],
    Currency: ['number'],
    Integer: ['number'],
    Choice: ['dropdown'],
    MultiChoice: ['multiselect'],
    DateTime: ['datetime'],
    Boolean: ['boolean'],
    URL: ['url', 'text'],
    Hyperlink: ['url', 'text'],
    Lookup: ['lookup'],
    LookupMulti: ['lookup'],
    User: ['person'],
    UserMulti: ['person'],
    TaxonomyFieldType: ['taxonomy'],
    TaxonomyFieldTypeMulti: ['taxonomy'],
    Image: ['image'],
    Thumbnail: ['image'],
    Attachments: ['attachment']
  };
  var allowed = compatibleTypes[sharePointType];
  return !!allowed && allowed.indexOf(String(controlType || '').toLowerCase()) >= 0;
}

export interface IGridControlState {
  selectedViewId: string;
  fields: IListFieldDefinition[];
  fieldMetadataByName: { [fieldName: string]: IGridFieldMetadata };
  lookupOptionsByField: { [fieldName: string]: IGridLookupOption[] };
  rows: any[];
  loading: boolean;
  loadingMore: boolean;
  nextPageHref: string;
  error: string | null;
  selectedItemId: number;
  selectedItemIds: number[];
  selectedMode: string;
  deleting: boolean;
  deleteMessage: string;
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
  collapsedGroupKeys: { [groupKey: string]: boolean };
  editingItemId: number;
  editingValues: { [fieldName: string]: any };
  editingErrors: { [fieldName: string]: string };
  saving: boolean;
  embeddedByDynamicForms: boolean;
  runtimeFilterJson: string;
  runtimeDefaultField: string;
  runtimeDefaultValue: any;
  runtimeReadOnly: boolean;
  runtimeConfigOwner: string;
  accessLoading: boolean;
  hasReadAccess: boolean;
  hasEditAccess: boolean;
  attachmentCountsByItemId: { [itemId: number]: number };
  showScrollArrows: boolean;
  scrollArrowTop: number;
  scrollArrowLeft: number;
  scrollArrowRight: number;
  historyDialogOpen: boolean;
  historyDialogUrl: string;
  historyDialogLoading: boolean;
  historyDialogError: string;
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
  borderColor?: string;
  borderStyle?: string;
  borderWidth?: string;
  borderRadius?: string;
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

var GRID_CONTROL_REFRESH_EVENT = 'spse:gridcontrol-refresh';
var GRID_CONTROL_RUNTIME_CONFIG_EVENT = 'spse:gridcontrol-runtime-config';
var GRID_CONTROL_RUNTIME_CONFIG_REQUEST_EVENT = 'spse:gridcontrol-runtime-config-request';

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

function extractRenderRowsAndFields(data: any): { rows: any[]; fields: IListFieldDefinition[]; nextHref: string } {
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
    fields: fields,
    nextHref: String(listData.NextHref || responseData.NextHref || '')
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

function toCssString(value: any): string {
  return String(value === undefined || value === null ? '' : value).trim();
}

function decodeRichTextHtml(value: string): string {
  if (!/&lt;\/?[a-z][\s\S]*?&gt;/i.test(value) || typeof document === 'undefined') {
    return value;
  }
  var decoder = document.createElement('textarea');
  decoder.innerHTML = value;
  return decoder.value;
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

export class GridControl extends React.Component<IGridControlProps, IGridControlState> {
  private _refreshEventHandler: any;
  private _runtimeConfigEventHandler: any;
  private _loadRowsRequestId: number = 0;
  private _listItemEntityTypeName: string = '';
  private _listItemEntityTypeListName: string = '';
  private _listId: string = '';
  private _listIdListName: string = '';
  private _tableWrapEl: HTMLDivElement;
  private _tableHeadEl: HTMLTableSectionElement;
  private _stickyHeaderViewportEl: HTMLDivElement;
  private _stickyHeaderSourceHtml: string = '';
  private _stickyHeaderVisible: boolean = false;
  private _filterAnchorEl: HTMLElement;
  private _pagingEndpoint: string = '';
  private _pagingRequestBody: any = undefined;
  private _pagingSchemaFieldNames: string[] = [];
  private _scrollArrowResizeHandler: any;
  private _scrollArrowScrollHandler: any;

  public constructor(props: IGridControlProps) {
    super(props);

    this.state = {
      selectedViewId: props.defaultViewId || this.getInitialViewId(props.views),
      fields: [],
      fieldMetadataByName: {},
      lookupOptionsByField: {},
      rows: [],
      loading: true,
      loadingMore: false,
      nextPageHref: '',
      error: null,
      selectedItemId: 0,
      selectedItemIds: [],
      selectedMode: 'view',
      deleting: false,
      deleteMessage: '',
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
      collapsedGroupKeys: {},
      editingItemId: -1,
      editingValues: {},
      editingErrors: {},
      saving: false,
      embeddedByDynamicForms: false,
      runtimeFilterJson: '',
      runtimeDefaultField: '',
      runtimeDefaultValue: '',
      runtimeReadOnly: false,
      runtimeConfigOwner: '',
      accessLoading: true,
      hasReadAccess: false,
      hasEditAccess: false,
      attachmentCountsByItemId: {},
      showScrollArrows: false,
      scrollArrowTop: 0,
      scrollArrowLeft: 0,
      scrollArrowRight: 0,
      historyDialogOpen: false,
      historyDialogUrl: '',
      historyDialogLoading: false,
      historyDialogError: '',
    };

    this._refreshEventHandler = this.handleExternalRefresh.bind(this);
    this._runtimeConfigEventHandler = this.handleRuntimeConfig.bind(this);
    this._scrollArrowResizeHandler = this.refreshTableViewport.bind(this);
    this._scrollArrowScrollHandler = this.refreshTableViewport.bind(this);
  }

  public componentDidMount(): void {
    this.logDiagnostic('Component mounted. listName=' + String(this.props.listName || '(none)') + ', defaultViewId=' + String(this.props.defaultViewId || '(none)'));
    if (typeof window !== 'undefined' && window.addEventListener) {
      window.addEventListener(GRID_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
      window.addEventListener(GRID_CONTROL_RUNTIME_CONFIG_EVENT, this._runtimeConfigEventHandler);
      var requestDetail = {
        instanceId: String(this.props.context && this.props.context.instanceId || '').toLowerCase()
      };
      var requestEvent: any;
      if (typeof (window as any).CustomEvent === 'function') {
        requestEvent = new (window as any).CustomEvent(GRID_CONTROL_RUNTIME_CONFIG_REQUEST_EVENT, { detail: requestDetail });
      } else {
        requestEvent = document.createEvent('CustomEvent');
        requestEvent.initCustomEvent(GRID_CONTROL_RUNTIME_CONFIG_REQUEST_EVENT, false, false, requestDetail);
      }
      window.dispatchEvent(requestEvent);
    }
    this.loadAccessAndRows();
    if (typeof window !== 'undefined' && window.addEventListener) {
      window.addEventListener('resize', this._scrollArrowResizeHandler);
      window.addEventListener('scroll', this._scrollArrowScrollHandler, true);
    }
  }

  public componentWillUnmount(): void {
    this._loadRowsRequestId += 1;
    if (typeof window !== 'undefined' && window.removeEventListener) {
      window.removeEventListener(GRID_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
      window.removeEventListener(GRID_CONTROL_RUNTIME_CONFIG_EVENT, this._runtimeConfigEventHandler);
      window.removeEventListener('resize', this._scrollArrowResizeHandler);
      window.removeEventListener('scroll', this._scrollArrowScrollHandler, true);
    }
  }

  public componentDidUpdate(prevProps: IGridControlProps, prevState: IGridControlState): void {
    if (prevProps.readSecurityGroupId !== this.props.readSecurityGroupId
      || prevProps.editSecurityGroupId !== this.props.editSecurityGroupId) {
      this.loadAccessAndRows();
      return;
    }

    if (prevProps.listName !== this.props.listName || prevProps.defaultViewId !== this.props.defaultViewId
      || prevProps.gridSchemaJson !== this.props.gridSchemaJson) {
      this.logDiagnostic('Props changed; resetting selection and reloading rows. listName=' + String(this.props.listName || '(none)') + ', viewId=' + String(this.props.defaultViewId || '(none)'));
      if (prevProps.listName !== this.props.listName) {
        this._listId = '';
        this._listIdListName = '';
      }
      var nextSelectedViewId = this.props.defaultViewId || this.getInitialViewId(this.props.views);
      var selectedViewWillChange = nextSelectedViewId !== this.state.selectedViewId;
      this.setState({
        selectedViewId: nextSelectedViewId,
        selectedItemId: 0,
        selectedItemIds: [],
        selectedMode: 'view',
        deleteMessage: '',
        sortFieldName: '',
        sortDirection: '',
        activeFilterFieldName: '',
        draftFilterOperator: 'contains',
        draftFilterValue: '',
        columnFilters: {},
        currentPage: 0,
        collapsedGroupKeys: {},
        editingItemId: -1,
        editingValues: {},
        editingErrors: {},
        saving: false,
        historyDialogOpen: false,
        historyDialogUrl: '',
        historyDialogLoading: false,
        historyDialogError: '',
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
    if (prevProps.fetchBatchSize !== this.props.fetchBatchSize) {
      this.loadRows();
      return;
    }

    if ((prevState.sortFieldName !== this.state.sortFieldName
      || prevState.sortDirection !== this.state.sortDirection
      || prevState.columnFilters !== this.state.columnFilters
      || prevState.runtimeFilterJson !== this.state.runtimeFilterJson
      || prevProps.filterJson !== this.props.filterJson) && this.state.nextPageHref) {
      this.loadAllRemainingRows();
    }

    if (typeof window !== 'undefined') {
      this._stickyHeaderSourceHtml = '';
      window.setTimeout(() => this.refreshTableViewport(), 0);
    }
  }

  private _setTableWrapRef = (el: HTMLDivElement): void => {
    this._tableWrapEl = el;
    this.refreshTableViewport();
  }

  private _setTableHeadRef = (el: HTMLTableSectionElement): void => {
    this._tableHeadEl = el;
    this.updateStickyHeaderPosition();
  }

  private _setStickyHeaderViewportRef = (el: HTMLDivElement): void => {
    this._stickyHeaderViewportEl = el;
    if (el) {
      el.onclick = this.handleStickyHeaderClick;
    }
    this.updateStickyHeaderPosition();
  }

  private handleStickyHeaderClick = (event: MouseEvent): void => {
    if (!this._stickyHeaderViewportEl || !this._tableHeadEl) { return; }
    var target = event.target as HTMLElement;
    while (target && target !== this._stickyHeaderViewportEl && !target.getAttribute('data-lc-sticky-index')) {
      target = target.parentElement;
    }
    if (!target || target === this._stickyHeaderViewportEl) { return; }
    event.preventDefault();
    event.stopPropagation();
    var sourceControls = this.getStickyHeaderControls(this._tableHeadEl);
    var sourceIndex = parseInt(String(target.getAttribute('data-lc-sticky-index') || ''), 10);
    var sourceControl = sourceControls[sourceIndex] as HTMLElement;
    if (!sourceControl) { return; }
    var isFilterButton = (' ' + String(target.className || '') + ' ').indexOf(' lc-header-action ') >= 0;
    if (isFilterButton) {
      this._filterAnchorEl = target;
    }
    (sourceControl as any).click();
    if (isFilterButton && typeof window !== 'undefined') {
      this._filterAnchorEl = target;
      window.setTimeout(() => this.updateActiveFilterPosition(), 0);
    }
  }

  private getStickyHeaderControls(root: Element): Element[] {
    var controls = root.querySelectorAll('button,input');
    var result: Element[] = [];
    for (var index = 0; index < controls.length; index += 1) {
      var control = controls[index] as HTMLElement;
      var ancestor = control.parentElement;
      var insideFilter = false;
      while (ancestor && ancestor !== root) {
        if ((' ' + String(ancestor.className || '') + ' ').indexOf(' lc-filter-popover ') >= 0) {
          insideFilter = true;
          break;
        }
        ancestor = ancestor.parentElement;
      }
      if (!insideFilter) {
        result.push(control);
      }
    }
    return result;
  }

  private refreshTableViewport(): void {
    this.refreshScrollArrows();
    this.updateStickyHeaderPosition();
    this.updateActiveFilterPosition();
  }

  private refreshScrollArrows(): void {
    this.updateScrollArrowVisibility();
    this.updateScrollArrowPosition();
  }

  private updateScrollArrowVisibility(): void {
    if (!this._tableWrapEl) {
      return;
    }

    var overflowing = this._tableWrapEl.scrollWidth > this._tableWrapEl.clientWidth + 1;
    if (overflowing !== this.state.showScrollArrows) {
      this.setState({ showScrollArrows: overflowing });
    }
  }

  // Positions the arrows via fixed coordinates centered on the currently visible slice of the
  // table (intersection of its bounding rect with the viewport), so they stay reachable in the
  // middle of the view regardless of vertical scroll position instead of drifting to the bottom.
  private updateScrollArrowPosition(): void {
    if (!this._tableWrapEl || typeof window === 'undefined') {
      return;
    }

    var viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
    var viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
    var rect = this._tableWrapEl.getBoundingClientRect();
    var visibleTop = Math.max(rect.top, 0);
    var visibleBottom = Math.min(rect.bottom, viewportHeight);
    var centerY = Math.round((visibleTop + visibleBottom) / 2);
    var leftPx = Math.round(rect.left);
    var rightPx = Math.round(viewportWidth - rect.right);

    if (centerY !== this.state.scrollArrowTop || leftPx !== this.state.scrollArrowLeft || rightPx !== this.state.scrollArrowRight) {
      this.setState({ scrollArrowTop: centerY, scrollArrowLeft: leftPx, scrollArrowRight: rightPx });
    }
  }

  private updateStickyHeaderPosition(): void {
    if (!this._tableWrapEl || !this._tableHeadEl || !this._stickyHeaderViewportEl) { return; }
    var wrapRect = this._tableWrapEl.getBoundingClientRect();
    var headerHeight = this._tableHeadEl.getBoundingClientRect().height;
    var stickyTop = this.getStickyViewportTop(wrapRect);
    var shouldStick = wrapRect.top < stickyTop && wrapRect.bottom > stickyTop + headerHeight;
    if (shouldStick !== this._stickyHeaderVisible) {
      this._stickyHeaderVisible = shouldStick;
      this.logDiagnostic('Sticky header ' + (shouldStick ? 'shown' : 'hidden')
        + '. tableTop=' + String(Math.round(wrapRect.top))
        + ', tableBottom=' + String(Math.round(wrapRect.bottom))
        + ', stickyTop=' + String(Math.round(stickyTop))
        + ', headerHeight=' + String(Math.round(headerHeight)));
    }
    if (!shouldStick) {
      this._stickyHeaderViewportEl.style.display = 'none';
      return;
    }

    var sourceHtml = this._tableHeadEl.innerHTML;
    if (sourceHtml !== this._stickyHeaderSourceHtml || this._stickyHeaderViewportEl.children.length === 0) {
      var activeStickyIndex = this._filterAnchorEl && this._stickyHeaderViewportEl.contains(this._filterAnchorEl)
        ? parseInt(String(this._filterAnchorEl.getAttribute('data-lc-sticky-index') || ''), 10) : -1;
      this._stickyHeaderViewportEl.innerHTML = '';
      var stickyTable = document.createElement('table');
      stickyTable.className = 'lc-table lc-sticky-header-table';
      var stickyHead = this._tableHeadEl.cloneNode(true) as HTMLTableSectionElement;
      var filterPopovers = stickyHead.querySelectorAll('.lc-filter-popover');
      for (var popoverIndex = 0; popoverIndex < filterPopovers.length; popoverIndex += 1) {
        var popover = filterPopovers[popoverIndex];
        if (popover.parentElement) {
          popover.parentElement.removeChild(popover);
        }
      }
      var stickyControls = this.getStickyHeaderControls(stickyHead);
      for (var controlIndex = 0; controlIndex < stickyControls.length; controlIndex += 1) {
        stickyControls[controlIndex].setAttribute('data-lc-sticky-index', String(controlIndex));
      }
      if (this.state.activeFilterFieldName && activeStickyIndex >= 0 && activeStickyIndex < stickyControls.length) {
        this._filterAnchorEl = stickyControls[activeStickyIndex] as HTMLElement;
      }
      var stickyColGroup = document.createElement('colgroup');
      var stickyHeaderCells = stickyHead.children.length > 0 ? stickyHead.children[0].children : [];
      for (var colIndex = 0; colIndex < stickyHeaderCells.length; colIndex += 1) {
        stickyColGroup.appendChild(document.createElement('col'));
      }
      stickyTable.appendChild(stickyColGroup);
      stickyTable.appendChild(stickyHead);
      this._stickyHeaderViewportEl.appendChild(stickyTable);
      this._stickyHeaderSourceHtml = sourceHtml;
    }

    var tableElement = this._tableHeadEl.parentElement as HTMLElement;
    var clonedTable = this._stickyHeaderViewportEl.children[0] as HTMLElement;
    var sourceCells = this._tableHeadEl.children.length > 0 ? this._tableHeadEl.children[0].children : [];
    var clonedColGroup = clonedTable.children[0] as HTMLElement;
    var clonedHead = clonedTable.children[1] as HTMLElement;
    var clonedCells = clonedHead && clonedHead.children.length > 0 ? clonedHead.children[0].children : [];
    var clonedColumns = clonedColGroup ? clonedColGroup.children : [];
    var totalTableWidth = 0;
    for (var cellIndex = 0; cellIndex < sourceCells.length && cellIndex < clonedCells.length; cellIndex += 1) {
      var measuredCellWidth = (sourceCells[cellIndex] as HTMLElement).getBoundingClientRect().width;
      var cellWidth = measuredCellWidth.toFixed(2) + 'px';
      var clonedCell = clonedCells[cellIndex] as HTMLElement;
      clonedCell.style.width = '';
      clonedCell.style.minWidth = '';
      clonedCell.style.maxWidth = '';
      if (cellIndex < clonedColumns.length) {
        (clonedColumns[cellIndex] as HTMLElement).style.width = cellWidth;
      }
      totalTableWidth += measuredCellWidth;
    }
    clonedTable.style.width = totalTableWidth.toFixed(2) + 'px';
    clonedTable.style.left = Math.round(-this._tableWrapEl.scrollLeft) + 'px';
    this._stickyHeaderViewportEl.style.display = 'block';
    this._stickyHeaderViewportEl.style.top = Math.round(stickyTop) + 'px';
    this._stickyHeaderViewportEl.style.left = Math.round(wrapRect.left + this._tableWrapEl.clientLeft) + 'px';
    this._stickyHeaderViewportEl.style.width = Math.round(this._tableWrapEl.clientWidth) + 'px';
    this._stickyHeaderViewportEl.style.height = Math.round(headerHeight) + 'px';
  }

  private getStickyViewportTop(wrapRect: ClientRect): number {
    if (typeof document === 'undefined' || typeof window === 'undefined' || !document.elementFromPoint) { return 0; }
    var viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
    var probeX = Math.max(1, Math.min((window.innerWidth || document.documentElement.clientWidth || 0) - 1, Math.round((wrapRect.left + wrapRect.right) / 2)));
    var stickyTop = this.getScrollViewportTop(viewportHeight);
    for (var probeCount = 0; probeCount < 6 && stickyTop < Math.min(240, viewportHeight); probeCount += 1) {
      var probeY = Math.max(1, stickyTop + 1);
      var element: HTMLElement = document.elementFromPoint(probeX, probeY) as HTMLElement;
      var chromeElement: HTMLElement | undefined = undefined;
      while (element && element !== document.body) {
        var position = window.getComputedStyle(element).position;
        var elementRect = element.getBoundingClientRect();
        if (element !== this._stickyHeaderViewportEl && (position === 'fixed' || position === 'sticky') && elementRect.top <= probeY && elementRect.bottom > probeY) {
          chromeElement = element;
        }
        element = element.parentElement;
      }
      if (!chromeElement) { break; }
      var chromeBottom = Math.round(chromeElement.getBoundingClientRect().bottom);
      if (chromeBottom <= stickyTop) { break; }
      stickyTop = chromeBottom;
    }
    return stickyTop;
  }

  private getScrollViewportTop(viewportHeight: number): number {
    if (!this._tableWrapEl || typeof window === 'undefined') { return 0; }
    var stickyTop = 0;
    var ancestor = this._tableWrapEl.parentElement;
    while (ancestor && ancestor !== document.body && ancestor !== document.documentElement) {
      var computedStyle = window.getComputedStyle(ancestor);
      var overflowY = String(computedStyle.overflowY || '').toLowerCase();
      var canScrollVertically = (overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay')
        && ancestor.scrollHeight > ancestor.clientHeight + 1;
      if (canScrollVertically) {
        var ancestorRect = ancestor.getBoundingClientRect();
        if (ancestorRect.bottom > 0 && ancestorRect.top < viewportHeight) {
          stickyTop = Math.max(stickyTop, Math.max(0, Math.round(ancestorRect.top)));
        }
      }
      ancestor = ancestor.parentElement;
    }
    return stickyTop;
  }

  private scrollTableHorizontally(direction: number): void {
    if (!this._tableWrapEl) {
      return;
    }

    var amount = Math.max(120, Math.round(this._tableWrapEl.clientWidth * 0.6));
    var nextLeft = this._tableWrapEl.scrollLeft + (amount * direction);
    if (this._tableWrapEl.scrollTo) {
      this._tableWrapEl.scrollTo({ left: nextLeft, behavior: 'smooth' });
    } else {
      this._tableWrapEl.scrollLeft = nextLeft;
    }
  }

  private getInitialViewId(views: IGridControlViewOption[]): string {
    for (var i = 0; i < views.length; i += 1) {
      if (views[i].isDefault) {
        return String(views[i].key);
      }
    }
    return views.length > 0 ? String(views[0].key) : '';
  }

  private getSelectedViewLabel(): string {
    var views = this.props.views || [];
    for (var i = 0; i < views.length; i += 1) {
      if (String(views[i].key) === String(this.state.selectedViewId)) {
        return views[i].text;
      }
    }
    return '';
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
      this.cancelRuntimeEdit();
      this.setState({
        embeddedByDynamicForms: false,
        runtimeFilterJson: '',
        runtimeDefaultField: '',
        runtimeDefaultValue: '',
        runtimeReadOnly: false,
        runtimeConfigOwner: '',
        currentPage: 0
      });
      return;
    }

    var nextReadOnly = detail.readOnly === true;
    if (nextReadOnly) {
      this.cancelRuntimeEdit();
    }
    this.setState({
      embeddedByDynamicForms: true,
      runtimeFilterJson: String(detail.filterJson || ''),
      runtimeDefaultField: String(detail.defaultField || ''),
      runtimeDefaultValue: detail.defaultValue,
      runtimeReadOnly: nextReadOnly,
      runtimeConfigOwner: owner,
      currentPage: 0
    });
  }

  private cancelRuntimeEdit(): void {
    if (this.state.editingItemId < 0) {
      return;
    }
    this.setState({
      editingItemId: -1,
      editingValues: {},
      editingErrors: {},
      saving: false,
      selectedMode: 'view'
    });
    this.props.onSelectionChange(this.state.selectedItemId, 'view');
  }

  private getWebUrl(): string {
    return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
  }

  private async loadAccessAndRows(): Promise<void> {
    var readGroupId = String(this.props.readSecurityGroupId || '').trim();
    var editGroupId = String(this.props.editSecurityGroupId || '').trim();
    if (!readGroupId && !editGroupId) {
      this.setState({ accessLoading: false, hasReadAccess: true, hasEditAccess: true }, () => this.loadRows());
      return;
    }

    this.setState({ accessLoading: true, hasReadAccess: false, hasEditAccess: false, rows: [], fields: [] });
    try {
      var response = await this.getJsonWithFallback(this.getWebUrl() + '/_api/web/currentuser/groups?$select=Id');
      if (!response.ok) {
        throw new Error('Failed to load current user groups. HTTP ' + String(response.status) + ' ' + String(response.statusText || ''));
      }
      var data = await response.json();
      var groups = toArray(data && data.value);
      if (groups.length === 0) {
        groups = toArray(data && data.d && data.d.results);
      }
      var groupIds = groups.map(function(group: any) { return String(group.Id); });
      this.logDiagnostic('Evaluated SharePoint group access. currentUserGroupIds=' + groupIds.join(',') + ', readGroupId=' + readGroupId + ', editGroupId=' + editGroupId);
      var hasEditAccess = !editGroupId || groupIds.indexOf(editGroupId) >= 0;
      var hasReadAccess = hasEditAccess || !readGroupId || groupIds.indexOf(readGroupId) >= 0;
      this.setState({
        accessLoading: false,
        hasReadAccess: hasReadAccess,
        hasEditAccess: hasEditAccess,
        error: null
      }, () => {
        if (hasReadAccess) {
          this.loadRows();
        }
      });
    } catch (error) {
      this.logDiagnostic('Failed to evaluate SharePoint group access: ' + (error && error.message ? error.message : String(error)));
      this.setState({ accessLoading: false, hasReadAccess: false, hasEditAccess: false, rows: [], fields: [] });
    }
  }

  private isReadOnly(): boolean {
    return this.state.runtimeReadOnly || !this.state.hasEditAccess;
  }

  private getCommandButtonClass(className?: string): string {
    return joinClassNames([
      className || '',
      'gc-command-button',
      this.props.buttonDisplayMode === 'icon' ? 'gc-command-button-icon-only' : ''
    ]);
  }

  private renderCommandContent(iconName: string, label: string): React.ReactNode {
    if (this.props.buttonDisplayMode !== 'icon' && this.props.buttonDisplayMode !== 'iconText') {
      return label;
    }
    return (
      <span className="gc-command-content">
        <i className={'ms-Icon ms-Icon--' + iconName} aria-hidden="true"></i>
        <span className={this.props.buttonDisplayMode === 'icon' ? 'gc-visually-hidden' : ''}>{label}</span>
      </span>
    );
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

  private async logPostAttempt(label: string, url: string, response: any, payload: string): Promise<void> {
    if (this.props.enableDiagnostics === false) { return; }
    var responseText = '';
    try {
      responseText = await response.clone().text();
    } catch (_responseReadError) {
      responseText = '';
    }
    this.logDiagnostic('POST ' + label + ' status=' + String(response.status) + ' ' + String(response.statusText || '')
      + ', url=' + url + ', payload=' + payload.substring(0, 2000)
      + (responseText ? ', response=' + responseText.substring(0, 2000) : ''));
  }

  private async postJsonWithFallback(url: string, body: any, baseHeaders?: { [key: string]: string }, verboseBody?: any, allowFallback?: boolean): Promise<any> {
    var payload = JSON.stringify(body || {});
    var verbosePayload = JSON.stringify(verboseBody || body || {});
    var createHeaders = function(contentType: string, accept?: string): { [key: string]: string } {
      var headers: { [key: string]: string } = {};
      var headerName: string;
      if (baseHeaders) {
        for (headerName in baseHeaders) {
          if (Object.prototype.hasOwnProperty.call(baseHeaders, headerName)) {
            headers[headerName] = baseHeaders[headerName];
          }
        }
      }
      headers['Content-Type'] = contentType;
      if (accept) { headers.Accept = accept; }
      return headers;
    };
    var response: any;
    var preferredErrorResponse: any = null;
    if (verboseBody) {
      var verboseItemHeaders = createHeaders(
        'application/json;odata=verbose;charset=utf-8',
        'application/json;odata=verbose'
      );
      verboseItemHeaders['OData-Version'] = '3.0';
      response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: verboseItemHeaders,
        body: verbosePayload
      });
      await this.logPostAttempt('verbose-item-v3', url, response, verbosePayload);
      if (response.ok) { return response; }
      if (allowFallback === false) { return response; }
      if (response.status !== 406) { preferredErrorResponse = response; }
    }

    response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: createHeaders('application/json; charset=utf-8', 'application/json;odata=verbose'),
      body: payload
    });
    await this.logPostAttempt('generic-json', url, response, payload);
    if (response.ok) { return response; }
    if (response.status !== 406) { preferredErrorResponse = response; }

    response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: createHeaders('application/json;odata=verbose', 'application/json;odata=verbose'),
      body: verbosePayload
    });
    await this.logPostAttempt('verbose-json', url, response, verbosePayload);
    if (response.ok) { return response; }
    if (response.status !== 406) { preferredErrorResponse = response; }

    response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: createHeaders('application/json;odata=minimalmetadata', 'application/json;odata=minimalmetadata'),
      body: payload
    });
    await this.logPostAttempt('minimal-json', url, response, payload);
    if (response.ok) { return response; }
    if (response.status !== 406) { preferredErrorResponse = response; }

    response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: createHeaders('application/json;odata=nometadata', 'application/json;odata=nometadata'),
      body: payload
    });
    await this.logPostAttempt('nometadata-json', url, response, payload);
    if (response.ok) { return response; }

    return preferredErrorResponse || response;
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

  private addFieldsToViewXml(viewXml: string, fieldNames: string[]): string {
    if (!viewXml || fieldNames.length === 0) {
      return viewXml;
    }

    var xmlDocument = new DOMParser().parseFromString(viewXml, 'text/xml');
    if (xmlDocument.getElementsByTagName('parsererror').length > 0) {
      throw new Error('The selected SharePoint view definition is invalid.');
    }

    var viewElements = xmlDocument.getElementsByTagName('View');
    if (viewElements.length === 0) {
      throw new Error('The selected SharePoint view definition does not contain a View element.');
    }

    var viewElement = viewElements[0];
    var viewFieldsElements = viewElement.getElementsByTagName('ViewFields');
    var viewFieldsElement = viewFieldsElements.length > 0 ? viewFieldsElements[0] : xmlDocument.createElement('ViewFields');
    if (viewFieldsElements.length === 0) {
      viewElement.appendChild(viewFieldsElement);
    }

    var existingFields: { [fieldName: string]: boolean } = {};
    var fieldRefs = viewFieldsElement.getElementsByTagName('FieldRef');
    for (var i = 0; i < fieldRefs.length; i += 1) {
      existingFields[String(fieldRefs[i].getAttribute('Name') || '').toLowerCase()] = true;
    }

    for (var j = 0; j < fieldNames.length; j += 1) {
      var fieldName = String(fieldNames[j] || '');
      if (!fieldName || existingFields[fieldName.toLowerCase()]) {
        continue;
      }

      var fieldRef = xmlDocument.createElement('FieldRef');
      fieldRef.setAttribute('Name', fieldName);
      viewFieldsElement.appendChild(fieldRef);
      existingFields[fieldName.toLowerCase()] = true;
    }

    return new XMLSerializer().serializeToString(xmlDocument);
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
    var baseFields: IListFieldDefinition[] = this.state.fields;

    if (this.state.selectedViewId === this.props.defaultViewId && this.props.viewColumns && this.props.viewColumns.length > 0) {
      var configuredByName: { [fieldName: string]: IListFieldDefinition } = {};
      for (var i = 0; i < this.state.fields.length; i += 1) {
        var field = this.state.fields[i];
        var fieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
        var responseName = String(field.Name || '').toLowerCase();
        if (fieldName) { configuredByName[fieldName] = field; }
        if (responseName) { configuredByName[responseName] = field; }
      }

      var configured: IListFieldDefinition[] = [];
      var addedViewColumnNames: { [name: string]: boolean } = {};
      for (var j = 0; j < this.props.viewColumns.length; j += 1) {
        var column = this.props.viewColumns[j];
        var columnKey = String(column.fieldName || '').toLowerCase();
        // Guard against a duplicate saved column (e.g. a SharePoint view listing the same field
        // twice) rendering the same value twice with no visible separator between cells.
        if (!columnKey || addedViewColumnNames[columnKey]) { continue; }
        var configuredMatch = configuredByName[columnKey];
        if (configuredMatch) {
          addedViewColumnNames[columnKey] = true;
          configured.push(Object.assign({}, configuredMatch, {
            DisplayName: column.displayName || configuredMatch.DisplayName || configuredMatch.Name,
            ConfiguredWidth: column.width || ''
          }));
        }
      }
      baseFields = configured;
    }

    var schemaFields = this.getGridSchemaFields();
    if (schemaFields.length === 0) {
      return baseFields;
    }

    var byName: { [fieldName: string]: IListFieldDefinition } = {};
    for (var baseIndex = 0; baseIndex < this.state.fields.length; baseIndex += 1) {
      var baseField = this.state.fields[baseIndex];
      var baseFieldName = String(baseField.Name || '').toLowerCase();
      byName[baseFieldName] = baseField;
      byName[String(baseField.RealFieldName || '').toLowerCase()] = baseField;
      if (baseFieldName === 'linktitle' || baseFieldName === 'linktitlenomenu') {
        byName.title = baseField;
      }
    }
    var schemaDisplayFields: IListFieldDefinition[] = [];
    var addedSchemaFieldNames: { [name: string]: boolean } = {};
    for (var schemaIndex = 0; schemaIndex < schemaFields.length; schemaIndex += 1) {
      var schemaField = schemaFields[schemaIndex];
      if (schemaField.visible === false) {
        continue;
      }
      var schemaFieldKey = String(schemaField.fieldName || '').toLowerCase();
      // Guard against a duplicate schema entry rendering the same value twice with no visible
      // separator between cells.
      if (!schemaFieldKey || addedSchemaFieldNames[schemaFieldKey]) { continue; }
      var match = byName[schemaFieldKey];
      if (match) {
        addedSchemaFieldNames[schemaFieldKey] = true;
        schemaDisplayFields.push(Object.assign({}, match, {
          DisplayName: schemaField.label || match.DisplayName || match.Name,
          ConfiguredWidth: schemaField.gridWidth || ''
        }));
      }
    }
    return schemaDisplayFields;
  }

  private getGridSchemaFields(): IGridSchemaField[] {
    var schema = this.getGridSchema();
    try {
      var steps = schema && Array.isArray(schema.steps) ? schema.steps : [];
      var fields: IGridSchemaField[] = [];
      for (var stepIndex = 0; stepIndex < steps.length; stepIndex += 1) {
        var stepFields = steps[stepIndex] && Array.isArray(steps[stepIndex].fields) ? steps[stepIndex].fields : [];
        for (var fieldIndex = 0; fieldIndex < stepFields.length; fieldIndex += 1) {
          var field = stepFields[fieldIndex];
          if (field && field.fieldName && field.type !== 'newline' && field.type !== 'richtext') {
            fields.push(field as IGridSchemaField);
          }
        }
      }
      return fields;
    } catch (_error) {
      return [];
    }
  }

  private getGridSchema(): any {
    var json = String(this.props.gridSchemaJson || '').trim();
    if (!json) { return {}; }
    try { return JSON.parse(json); } catch (_error) { return {}; }
  }

  private getAdvancedValidationRules(): IGridAdvancedValidationRule[] {
    var schema = this.getGridSchema();
    var advancedValidation = schema && schema.advancedValidation;
    return advancedValidation && advancedValidation.enabled === true && Array.isArray(advancedValidation.rules)
      ? advancedValidation.rules : [];
  }

  private getGridGroupingConfig(): IGridGroupingConfig {
    var schema = this.getGridSchema();
    var grouping = schema && schema.grouping;
    var field1 = grouping ? String(grouping.field1 || '') : '';
    return {
      enabled: !!(grouping && grouping.enabled === true && field1),
      field1: field1,
      field2: grouping ? String(grouping.field2 || '') : '',
      collapsedByDefault: !!(grouping && grouping.collapsedByDefault === true),
      showCount: !grouping || grouping.showCount !== false
    };
  }

  private getGridSchemaField(fieldName: string): IGridSchemaField | undefined {
    var normalized = String(fieldName || '').toLowerCase();
    var fields = this.getGridSchemaFields();
    for (var i = 0; i < fields.length; i += 1) {
      if (String(fields[i].fieldName || '').toLowerCase() === normalized) {
        return fields[i];
      }
    }
    return undefined;
  }

  private getConfiguredColumnStyle(field: IListFieldDefinition): React.CSSProperties {
    var configuredWidth = String(field.ConfiguredWidth || '').trim();
    if (!configuredWidth) {
      return {};
    }
    var width = /^\d+(?:\.\d+)?$/.test(configuredWidth) ? configuredWidth + 'px' : configuredWidth;
    if (!/^\d+(?:\.\d+)?(?:px|%|rem|em|vw)$/.test(width)) {
      return {};
    }
    return { width: width, minWidth: width, maxWidth: width };
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

  private async loadGridFieldMetadata(): Promise<{ [fieldName: string]: IGridFieldMetadata }> {
    try {
      var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName)
        + "')/fields?$select=InternalName,Title,TypeAsString,Required,ReadOnlyField,Hidden,Description,Choices,DisplayFormat,LookupList,LookupField,AllowMultipleValues,RichText";
      var response = await this.getJsonWithFallback(endpoint);
      if (!response.ok) {
        return {};
      }

      var data = await response.json();
      var fields = toArray(data.value);
      if (fields.length === 0) {
        fields = toArray(data && data.d && data.d.results);
      }

      var metadataByName: { [fieldName: string]: IGridFieldMetadata } = {};
      for (var i = 0; i < fields.length; i += 1) {
        var source = fields[i] || {};
        var internalName = String(source.InternalName || '');
        if (!internalName) {
          continue;
        }
        metadataByName[internalName.toLowerCase()] = {
          internalName: internalName,
          title: String(source.Title || internalName),
          typeAsString: String(source.TypeAsString || 'Text'),
          required: source.Required === true,
          readOnly: source.ReadOnlyField === true,
          hidden: source.Hidden === true,
          description: String(source.Description || ''),
          choices: toArray(source.Choices).map(function(choice: any) { return String(choice); }),
          displayFormat: parseInt(String(source.DisplayFormat || '0'), 10) || 0,
          lookupList: String(source.LookupList || '').replace(/^\{|\}$/g, ''),
          lookupField: String(source.LookupField || 'Title'),
          allowMultiple: source.AllowMultipleValues === true || String(source.TypeAsString || '').toLowerCase().indexOf('multi') >= 0,
          richText: source.RichText === true
        };
      }
      return metadataByName;
    } catch (error) {
      this.logDiagnostic('loadGridFieldMetadata failed: ' + (error && error.message ? error.message : String(error)));
      return {};
    }
  }

  private async loadListItemEntityTypeName(): Promise<string> {
    var listName = String(this.props.listName || '');
    if (this._listItemEntityTypeListName === listName && this._listItemEntityTypeName) {
      return this._listItemEntityTypeName;
    }
    var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(listName)
      + "')?$select=ListItemEntityTypeFullName";
    var response = await this.getJsonWithFallback(endpoint);
    if (!response.ok) {
      throw new Error('Unable to resolve the SharePoint list item entity type. HTTP ' + String(response.status) + ' ' + String(response.statusText || ''));
    }
    var data = await response.json();
    var source = data && data.d ? data.d : data;
    var entityTypeName = String(source && source.ListItemEntityTypeFullName || '');
    if (!entityTypeName) {
      throw new Error('SharePoint did not return ListItemEntityTypeFullName for list ' + listName + '.');
    }
    this._listItemEntityTypeListName = listName;
    this._listItemEntityTypeName = entityTypeName;
    this.logDiagnostic('Resolved list item entity type. listName=' + listName + ', type=' + entityTypeName);
    return entityTypeName;
  }

  private async loadLookupOptions(metadata: IGridFieldMetadata): Promise<IGridLookupOption[]> {
    var typeName = String(metadata.typeAsString || '');
    var endpoint = '';
    if (typeName === 'User' || typeName === 'UserMulti') {
      endpoint = this.getWebUrl() + '/_api/web/siteusers?$select=Id,Title,Email,PrincipalType&$top=5000';
    } else if ((typeName === 'Lookup' || typeName === 'LookupMulti') && metadata.lookupList) {
      var lookupField = /^[A-Za-z0-9_]+$/.test(metadata.lookupField) ? metadata.lookupField : 'Title';
      // On-prem SharePoint can report LookupList as a plain list title instead of a GUID.
      var lookupListSegment = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(metadata.lookupList)
        ? "lists(guid'" + metadata.lookupList + "')"
        : "lists/getByTitle('" + escapeODataText(metadata.lookupList) + "')";
      endpoint = this.getWebUrl() + '/_api/web/' + lookupListSegment + '/items?$select='
        + encodeURIComponent('Id,' + lookupField) + '&$top=5000';
    }
    if (!endpoint) {
      return [];
    }

    try {
      var response = await this.getJsonWithFallback(endpoint);
      if (!response.ok) {
        this.logDiagnostic('Lookup options unavailable for field ' + metadata.internalName + '. Status=' + String(response.status));
        return [];
      }
      var data = await response.json();
      var items = toArray(data.value);
      if (items.length === 0) {
        items = toArray(data && data.d && data.d.results);
      }
      return items.filter(function(item: any) {
        if (!item || toPositiveInt(item.Id) <= 0) { return false; }
        if (typeName !== 'User' && typeName !== 'UserMulti') { return true; }
        return parseInt(String(item.PrincipalType || '0'), 10) > 0;
      }).map(function(item: any) {
        return {
          id: toPositiveInt(item.Id),
          text: String(item[metadata.lookupField] || item.Title || item.Email || item.Id)
        };
      });
    } catch (error) {
      this.logDiagnostic('Lookup options failed for field ' + metadata.internalName + ': ' + (error && error.message ? error.message : String(error)));
      return [];
    }
  }

  private async loadLookupOptionsByField(
    metadataByName: { [fieldName: string]: IGridFieldMetadata },
    visibleFields: IListFieldDefinition[]
  ): Promise<{ [fieldName: string]: IGridLookupOption[] }> {
    var optionsByField: { [fieldName: string]: IGridLookupOption[] } = {};
    var visibleFieldNames: { [fieldName: string]: boolean } = {};
    visibleFields.forEach(function(field: IListFieldDefinition) {
      var fieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
      if (fieldName) {
        visibleFieldNames[fieldName] = true;
      }
    });
    var fieldName: string;
    for (fieldName in metadataByName) {
      if (!Object.prototype.hasOwnProperty.call(metadataByName, fieldName)) { continue; }
      var metadata = metadataByName[fieldName];
      if (visibleFieldNames[fieldName] && !metadata.hidden && !metadata.readOnly
        && (metadata.typeAsString === 'Lookup' || metadata.typeAsString === 'LookupMulti'
        || metadata.typeAsString === 'User' || metadata.typeAsString === 'UserMulti')) {
        optionsByField[fieldName] = await this.loadLookupOptions(metadata);
      }
    }
    return optionsByField;
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

  private applyFieldTypes(fields: IListFieldDefinition[], metadataByName: { [fieldName: string]: IGridFieldMetadata }): IListFieldDefinition[] {
    return fields.map(function(field: IListFieldDefinition) {
      var internalName = String(field.RealFieldName || field.Name || '');
      var metadata = metadataByName[internalName.toLowerCase()];
      return {
        ...field,
        TypeAsString: metadata ? metadata.typeAsString : (field.TypeAsString || ''),
        DisplayFormat: metadata ? metadata.displayFormat : field.DisplayFormat
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

    if (typeof value === 'boolean') {
      return value ? strings.RuntimeBooleanYes : strings.RuntimeBooleanNo;
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

    if (typeof value === 'string') {
      // RenderListDataAsStream can return Person/Lookup fields as JSON-encoded strings.
      var trimmedValue = value.trim();
      if (trimmedValue.length > 0 && (trimmedValue.charAt(0) === '{' || trimmedValue.charAt(0) === '[')) {
        var parsedJsonValue = tryParseObject(trimmedValue);
        if (parsedJsonValue && typeof parsedJsonValue === 'object') {
          return this.stringifyCellValue(parsedJsonValue);
        }
      }
    }

    if (typeof value === 'object') {
      if (Array.isArray((value as any).results)) {
        return this.stringifyCellValue((value as any).results);
      }
      if ((value as any).lookupValue !== undefined && (value as any).lookupValue !== null) {
        return String((value as any).lookupValue);
      }
      if ((value as any).Title) {
        return String((value as any).Title);
      }
      if ((value as any).LookupValue) {
        return String((value as any).LookupValue);
      }
      if ((value as any).title) {
        return String((value as any).title);
      }
      if ((value as any).Name) {
        return String((value as any).Name);
      }
      if ((value as any).Email) {
        return String((value as any).Email);
      }
      if ((value as any).email) {
        return String((value as any).email);
      }
      if ((value as any).Url) {
        return String((value as any).Url);
      }
      if ((value as any).Id !== undefined && (value as any).Id !== null) {
        return String((value as any).Id);
      }
      if ((value as any).id !== undefined && (value as any).id !== null) {
        return String((value as any).id);
      }
      try {
        return JSON.stringify(value);
      } catch (_jsonError) {
        return String(value);
      }
    }

    return String(value);
  }

  private formatDateCellValue(value: any, field: IListFieldDefinition, row?: any): string {
    if (String(field.TypeAsString || '') === 'Attachments') {
      return String(this.getAttachmentCountForRow(row, value));
    }

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
    var monthPadded = padTwoDigits(Number(month));
    var dayPadded = padTwoDigits(Number(day));

    // Date-only fields build named tokens from the corrected local date to avoid a UTC day shift.
    var namedTokenDate = dateOnlyMatch ? new Date(Number(year), Number(month) - 1, Number(day)) : dateValue;

    var hours = dateValue.getHours();
    var twelveHour = hours % 12 || 12;
    var tokenValues: { [token: string]: string } = {
      YYYY: year,
      YY: year.length > 2 ? year.substring(year.length - 2) : year,
      MMMM: namedTokenDate.toLocaleString(undefined, { month: 'long' }),
      MMM: namedTokenDate.toLocaleString(undefined, { month: 'short' }),
      MM: monthPadded,
      M: String(Number(month)),
      DD: dayPadded,
      D: String(Number(day)),
      dddd: namedTokenDate.toLocaleString(undefined, { weekday: 'long' }),
      ddd: namedTokenDate.toLocaleString(undefined, { weekday: 'short' }),
      HH: padTwoDigits(hours),
      H: String(hours),
      hh: padTwoDigits(twelveHour),
      h: String(twelveHour),
      mm: padTwoDigits(dateValue.getMinutes()),
      ss: padTwoDigits(dateValue.getSeconds()),
      tt: hours >= 12 ? 'PM' : 'AM'
    };

    var dateText: string;
    if (this.props.dateDisplayFormat === 'custom') {
      dateText = applyTextCase(
        applyDateFormatPattern(String(this.props.dateCustomFormat || '').trim() || 'MM/DD/YYYY', tokenValues),
        String(this.props.dateCustomFormatCase || '')
      );
    } else if (this.props.dateDisplayFormat === 'dmy') {
      dateText = dayPadded + '/' + monthPadded + '/' + year;
    } else if (this.props.dateDisplayFormat === 'ymd') {
      dateText = year + '-' + monthPadded + '-' + dayPadded;
    } else {
      dateText = monthPadded + '/' + dayPadded + '/' + year;
    }

    if (field.DisplayFormat === 0) {
      return dateText;
    }

    var minutes = padTwoDigits(dateValue.getMinutes());
    if (this.props.timeDisplayFormat === 'custom') {
      // Custom time format fully controls the combined output so date/time tokens can be freely interleaved.
      return applyTextCase(
        applyDateFormatPattern(String(this.props.timeCustomFormat || '').trim() || 'MM/DD/YYYY HH:mm', tokenValues),
        String(this.props.timeCustomFormatCase || '')
      );
    }
    if (this.props.timeDisplayFormat === '12hour') {
      var period = hours >= 12 ? 'PM' : 'AM';
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

  private async loadRowsFromItemsEndpoint(viewFieldNames: string[], itemIds?: number[]): Promise<{ rows: any[]; fields: IListFieldDefinition[] }> {
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
    if (itemIds && itemIds.length > 0) {
      endpoint += '&$filter=' + encodeURIComponent(itemIds.map(function(itemId: number) {
        return 'ID eq ' + String(itemId);
      }).join(' or '));
    }
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

    if (itemIds && itemIds.length > 0 && rows.length > 1) {
      rows.sort(function(left: any, right: any) {
        return itemIds.indexOf(toPositiveInt(left.ID || left.Id || left.id))
          - itemIds.indexOf(toPositiveInt(right.ID || right.Id || right.id));
      });
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

  private async loadAttachmentCounts(itemIds: number[]): Promise<{ [itemId: number]: number }> {
    var counts: { [itemId: number]: number } = {};
    var uniqueItemIds = itemIds.filter(function(itemId: number, index: number) {
      return itemId > 0 && itemIds.indexOf(itemId) === index;
    });
    if (uniqueItemIds.length === 0) {
      return counts;
    }

    try {
      var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items"
        + '?$top=200&$select=ID,AttachmentFiles&$expand=AttachmentFiles'
        + '&$filter=' + encodeURIComponent(uniqueItemIds.map(function(itemId: number) {
          return 'ID eq ' + String(itemId);
        }).join(' or '));
      var response = await this.getJsonWithFallback(endpoint);
      if (!response.ok) {
        return counts;
      }

      var data = await response.json();
      var items = toArray(data.value);
      if (items.length === 0) {
        items = toArray(data && data.d && data.d.results);
      }

      for (var i = 0; i < items.length; i += 1) {
        var item = items[i];
        var itemId = toPositiveInt(item.ID || item.Id || item.id);
        var attachmentFiles = item.AttachmentFiles && item.AttachmentFiles.results ? item.AttachmentFiles.results : item.AttachmentFiles;
        counts[itemId] = Array.isArray(attachmentFiles) ? attachmentFiles.length : 0;
      }
    } catch (attachmentCountError) {
      this.logDiagnostic('loadAttachmentCounts failed: ' + (attachmentCountError && attachmentCountError.message ? attachmentCountError.message : String(attachmentCountError)));
    }

    return counts;
  }

  private async loadRows(): Promise<void> {
    if (this.state.accessLoading || !this.state.hasReadAccess) {
      this.setState({ loading: false, error: null, fields: [], rows: [] });
      return;
    }

    if (!this.props.listName) {
      this.setState({ loading: false, error: null, fields: [], rows: [] });
      return;
    }

    var requestId = ++this._loadRowsRequestId;
    this.logDiagnostic('Starting loadRows. requestId=' + String(requestId) + ', listName=' + String(this.props.listName) + ', selectedViewId=' + String(this.state.selectedViewId || '(none)'));
    this.setState({ loading: true, loadingMore: false, nextPageHref: '', error: null });

    try {
      var baseEndpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/RenderListDataAsStream";
      var selectedViewId = this.state.selectedViewId;
      var schemaFields = this.getGridSchemaFields();
      var schemaFieldNames = schemaFields.filter(function(field: IGridSchemaField) {
        return field.visible !== false && !!field.fieldName;
      }).map(function(field: IGridSchemaField) {
        return String(field.fieldName);
      });
      var groupingConfig = this.getGridGroupingConfig();
      // Ensure grouping fields are fetched even if the column itself is hidden in the schema.
      [groupingConfig.field1, groupingConfig.field2].forEach(function(groupFieldName: string) {
        if (groupFieldName && schemaFieldNames.indexOf(groupFieldName) < 0) {
          schemaFieldNames.push(groupFieldName);
        }
      });
      var viewFieldNames = await this.loadSelectedViewFieldNames(selectedViewId);
      var selectedViewXml = await this.loadSelectedViewXml(selectedViewId, viewFieldNames);

      var body: any = {
        parameters: {
          RenderOptions: 7
        }
      };
      body.parameters.ViewXml = this.applyFetchBatchSize(selectedViewXml || '<View></View>');
      this._pagingEndpoint = baseEndpoint;
      this._pagingRequestBody = body;
      this._pagingSchemaFieldNames = schemaFieldNames.slice(0);

      var selectedRows: any[] = [];
      var selectedFields: IListFieldDefinition[] = [];
      var nextPageHref = '';
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
        nextPageHref = extracted.nextHref;

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

      if (schemaFields.length > 0) {
        viewFieldNames = schemaFieldNames;
        if (selectedViewId && rows.length > 0) {
          var filteredItemIds = rows.map((row: any) => this.getRowItemId(row)).filter(function(itemId: number, index: number, values: number[]) {
            return itemId > 0 && values.indexOf(itemId) === index;
          });
          if (filteredItemIds.length > 0) {
            var filteredSchemaItems = await this.loadRowsFromItemsEndpoint(schemaFieldNames, filteredItemIds);
            if (filteredSchemaItems.rows.length > 0) {
              rows = filteredSchemaItems.rows;
              fields = filteredSchemaItems.fields;
              this.logDiagnostic('Hydrated selected view rows with grid schema fields. rows=' + String(rows.length));
            }
          }
        } else if (!selectedViewId && rows.length === 0) {
          var schemaItems = await this.loadRowsFromItemsEndpoint(schemaFieldNames);
          rows = schemaItems.rows;
          fields = schemaItems.fields;
        }
      } else if (rows.length === 0 && !selectedViewId) {
        var itemsFallback = await this.loadRowsFromItemsEndpoint(viewFieldNames);
        rows = itemsFallback.rows;

        if (fields.length === 0) {
          fields = itemsFallback.fields;
        }
      }

      var visibleFields = this.getFieldsForConsumption(fields, viewFieldNames);
      var fieldTitleMap = await this.loadListFieldTitleMap();
      var fieldMetadataByName = await this.loadGridFieldMetadata();
      var lookupOptionsByField = await this.loadLookupOptionsByField(fieldMetadataByName, visibleFields);
      visibleFields = this.applyFieldDisplayNames(visibleFields, fieldTitleMap);
      visibleFields = this.applyFieldTypes(visibleFields, fieldMetadataByName);
      var renderableRows = this.filterRenderableRows(rows, visibleFields);

      if (requestId !== this._loadRowsRequestId) {
        this.logDiagnostic('Ignoring stale loadRows result. requestId=' + String(requestId) + ', latestRequestId=' + String(this._loadRowsRequestId));
        return;
      }

      var renderableItemIds = renderableRows.map((row: any) => this.getRowItemId(row));
      var selectedItemIds = this.state.selectedItemIds.filter(function(itemId: number) {
        return renderableItemIds.indexOf(itemId) >= 0;
      });

      var hasAttachmentsField = visibleFields.some(function(field: IListFieldDefinition) {
        return String(field.TypeAsString || '') === 'Attachments';
      });
      var attachmentCountsByItemId = hasAttachmentsField
        ? await this.loadAttachmentCounts(renderableItemIds)
        : {};

      if (requestId !== this._loadRowsRequestId) {
        this.logDiagnostic('Ignoring stale loadRows result after attachment count fetch. requestId=' + String(requestId) + ', latestRequestId=' + String(this._loadRowsRequestId));
        return;
      }

      this.setState({
        fields: visibleFields,
        fieldMetadataByName: fieldMetadataByName,
        lookupOptionsByField: lookupOptionsByField,
        rows: renderableRows,
        loadingMore: false,
        nextPageHref: nextPageHref,
        selectedItemIds: selectedItemIds,
        attachmentCountsByItemId: attachmentCountsByItemId,
        loading: false,
        error: null
      }, () => {
        if ((this.parsePresetFilterConditions().length > 0 || this.getGridGroupingConfig().enabled) && this.state.nextPageHref) {
          this.loadAllRemainingRows();
        }
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
        loadingMore: false,
        nextPageHref: '',
        error: loadError && loadError.message ? loadError.message : 'Failed to load data.',
        fields: [],
        fieldMetadataByName: {},
        lookupOptionsByField: {},
        rows: []
      });
      this.logDiagnostic('loadRows failed: ' + (loadError && loadError.message ? loadError.message : String(loadError)));
    }
  }

  private applyFetchBatchSize(viewXml: string): string {
    var parser = new DOMParser();
    var xmlDocument = parser.parseFromString(String(viewXml || '<View></View>'), 'text/xml');
    var viewElement = xmlDocument.documentElement;
    var rowLimits = viewElement.getElementsByTagName('RowLimit');
    var rowLimit = rowLimits.length > 0 ? rowLimits[0] : xmlDocument.createElement('RowLimit');
    rowLimit.setAttribute('Paged', 'TRUE');
    while (rowLimit.firstChild) {
      rowLimit.removeChild(rowLimit.firstChild);
    }
    rowLimit.appendChild(xmlDocument.createTextNode(String(this.props.fetchBatchSize)));
    if (rowLimits.length === 0) {
      viewElement.appendChild(rowLimit);
    }
    return new XMLSerializer().serializeToString(xmlDocument);
  }

  private async loadNextBatch(): Promise<boolean> {
    if (this.state.loadingMore || !this.state.nextPageHref || !this._pagingEndpoint || !this._pagingRequestBody) {
      return false;
    }
    var requestId = this._loadRowsRequestId;
    this.setState({ loadingMore: true });
    try {
      var parameters: any = {};
      var sourceParameters = this._pagingRequestBody.parameters || {};
      for (var parameterName in sourceParameters) {
        if (Object.prototype.hasOwnProperty.call(sourceParameters, parameterName)) {
          parameters[parameterName] = sourceParameters[parameterName];
        }
      }
      parameters.Paging = String(this.state.nextPageHref).replace(/^\?/, '');
      var response = await this.postJsonWithFallback(this._pagingEndpoint, { parameters: parameters });
      if (!response.ok) {
        this.setState({ loadingMore: false });
        return false;
      }
      var extracted = extractRenderRowsAndFields(await response.json());
      if (requestId !== this._loadRowsRequestId) {
        return false;
      }
      var nextRows = extracted.rows;
      if (this._pagingSchemaFieldNames.length > 0 && nextRows.length > 0) {
        var itemIds = nextRows.map((row: any) => this.getRowItemId(row)).filter(function(itemId: number) { return itemId > 0; });
        var hydrated = await this.loadRowsFromItemsEndpoint(this._pagingSchemaFieldNames, itemIds);
        if (hydrated.rows.length > 0) {
          nextRows = hydrated.rows;
        }
      }
      if (requestId !== this._loadRowsRequestId) {
        return false;
      }
      nextRows = this.filterRenderableRows(nextRows, this.state.fields);
      var existingIds: { [itemId: number]: boolean } = {};
      this.state.rows.forEach((row: any) => { existingIds[this.getRowItemId(row)] = true; });
      var uniqueRows = nextRows.filter((row: any) => !existingIds[this.getRowItemId(row)]);
      var nextAttachmentCounts: { [itemId: number]: number } = {};
      var hasAttachments = this.state.fields.some(function(field: IListFieldDefinition) { return String(field.TypeAsString || '') === 'Attachments'; });
      if (hasAttachments) {
        nextAttachmentCounts = await this.loadAttachmentCounts(uniqueRows.map((row: any) => this.getRowItemId(row)));
      }
      var nextPageHref = extracted.nextHref === this.state.nextPageHref && uniqueRows.length === 0 ? '' : extracted.nextHref;
      await new Promise<boolean>((resolve) => this.setState({
        rows: this.state.rows.concat(uniqueRows),
        attachmentCountsByItemId: { ...this.state.attachmentCountsByItemId, ...nextAttachmentCounts },
        nextPageHref: nextPageHref,
        loadingMore: false
      }, () => resolve(true)));
      return uniqueRows.length > 0 || !!nextPageHref;
    } catch (error) {
      this.logDiagnostic('Loading next batch failed: ' + (error && error.message ? error.message : String(error)));
      this.setState({ loadingMore: false });
      return false;
    }
  }

  private loadAllRemainingRows(): void {
    if (this.state.loadingMore || !this.state.nextPageHref) { return; }
    this.loadNextBatch().then((loaded) => {
      if (loaded && this.state.nextPageHref) {
        this.loadAllRemainingRows();
      }
    });
  }

  private goToNextPage(currentPage: number, pageCount: number): void {
    if (currentPage < pageCount - 1) {
      var nextPage = currentPage + 1;
      this.setState({ currentPage: nextPage }, () => {
        if (this.state.nextPageHref && nextPage >= pageCount - 2) {
          this.loadNextBatch();
        }
      });
      return;
    }
    if (this.state.nextPageHref && !this.state.loadingMore) {
      this.loadNextBatch().then((loaded) => {
        if (loaded) {
          this.setState({ currentPage: currentPage + 1 });
        }
      });
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

  private isItemChecked(itemId: number): boolean {
    return this.state.selectedItemIds.indexOf(itemId) >= 0;
  }

  private toggleItemChecked(itemId: number): void {
    if (itemId <= 0 || this.state.deleting || this.isReadOnly()) {
      return;
    }
    var selectedItemIds = this.state.selectedItemIds.slice(0);
    var selectedIndex = selectedItemIds.indexOf(itemId);
    if (selectedIndex >= 0) {
      selectedItemIds.splice(selectedIndex, 1);
    } else {
      selectedItemIds.push(itemId);
    }
    this.setState({ selectedItemIds: selectedItemIds });
  }

  private toggleVisibleItemsChecked(rows: any[]): void {
    if (this.isReadOnly()) {
      return;
    }
    var visibleItemIds = rows.map((row: any) => this.getRowItemId(row)).filter(function(itemId: number) {
      return itemId > 0;
    });
    var allVisibleSelected = visibleItemIds.length > 0 && visibleItemIds.every((itemId: number) => this.isItemChecked(itemId));
    var selectedItemIds = this.state.selectedItemIds.filter(function(itemId: number) {
      return allVisibleSelected ? visibleItemIds.indexOf(itemId) < 0 : true;
    });
    if (!allVisibleSelected) {
      for (var itemIndex = 0; itemIndex < visibleItemIds.length; itemIndex += 1) {
        if (selectedItemIds.indexOf(visibleItemIds[itemIndex]) < 0) {
          selectedItemIds.push(visibleItemIds[itemIndex]);
        }
      }
    }
    this.setState({ selectedItemIds: selectedItemIds });
  }

  private getGridFieldMetadata(field: IListFieldDefinition): IGridFieldMetadata | undefined {
    var fieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
    if (fieldName === 'linktitle' || fieldName === 'linktitlenomenu') {
      fieldName = 'title';
    }
    var metadata = this.state.fieldMetadataByName[fieldName];
    if (!metadata) {
      metadata = this.state.fieldMetadataByName[String(field.Name || '').toLowerCase()];
    }
    if (!metadata) {
      return undefined;
    }

    var schemaField = this.getGridSchemaField(metadata.internalName);
    if (!schemaField) {
      return metadata;
    }
    var typeMap: { [fieldType: string]: string } = {
      text: 'Text',
      multiline: 'Note',
      number: 'Number',
      boolean: 'Boolean',
      dropdown: 'Choice',
      multiselect: 'MultiChoice',
      datetime: 'DateTime',
      url: 'URL',
      lookup: 'Lookup',
      person: 'User'
    };
    var config = schemaField.config || {};
    var displayFormat = config.displayFormat === 'dateOnly' ? 0 : config.displayFormat === 'timeOnly' ? 2 : 1;
    var schemaControlType = String(schemaField.type || '').toLowerCase();
    var effectiveType = isCompatibleGridControlType(metadata.typeAsString, schemaControlType)
      ? typeMap[schemaControlType] : metadata.typeAsString;
    return Object.assign({}, metadata, {
      title: schemaField.label || metadata.title,
      typeAsString: effectiveType,
      required: metadata.required || schemaField.required === true,
      readOnly: schemaField.readOnly === true || schemaField.disabled === true || metadata.readOnly,
      description: metadata.description || schemaField.description || config.helpText,
      choices: Array.isArray(config.choices) ? config.choices.map(function(choice: any) { return String(choice); }) : metadata.choices,
      displayFormat: schemaField.type === 'datetime' ? displayFormat : metadata.displayFormat
    });
  }

  private isEditableGridField(field: IListFieldDefinition): boolean {
    var metadata = this.getGridFieldMetadata(field);
    if (!metadata || metadata.readOnly || metadata.hidden) {
      return false;
    }
    var supportedTypes = ['Text', 'Note', 'Number', 'Currency', 'Integer', 'Boolean', 'Choice', 'MultiChoice', 'DateTime', 'URL', 'Lookup', 'LookupMulti', 'User', 'UserMulti'];
    return supportedTypes.indexOf(metadata.typeAsString) >= 0;
  }

  private getLookupIds(value: any): string[] {
    var ids: string[] = [];
    var collect = function(candidate: any): void {
      if (candidate === undefined || candidate === null || candidate === '') { return; }
      if (Array.isArray(candidate)) {
        for (var arrayIndex = 0; arrayIndex < candidate.length; arrayIndex += 1) { collect(candidate[arrayIndex]); }
        return;
      }
      if (typeof candidate === 'object') {
        if (candidate.results !== undefined) { collect(candidate.results); return; }
        var objectId = candidate.LookupId !== undefined ? candidate.LookupId
          : candidate.Id !== undefined ? candidate.Id : candidate.ID !== undefined ? candidate.ID : candidate.id;
        collect(objectId);
        return;
      }
      var text = String(candidate);
      var lookupMatches = text.match(/(?:^|;#)(\d+)(?=;#|$)/g);
      if (lookupMatches && lookupMatches.length > 0) {
        for (var matchIndex = 0; matchIndex < lookupMatches.length; matchIndex += 1) {
          var matchedId = lookupMatches[matchIndex].replace(';#', '');
          if (toPositiveInt(matchedId) > 0 && ids.indexOf(matchedId) < 0) { ids.push(matchedId); }
        }
        return;
      }
      if (toPositiveInt(text) > 0 && ids.indexOf(String(toPositiveInt(text))) < 0) { ids.push(String(toPositiveInt(text))); }
    };
    collect(value);
    return ids;
  }

  private normalizeEditingValue(value: any, metadata: IGridFieldMetadata): any {
    if (metadata.typeAsString === 'Boolean') {
      return value === true || value === 1 || String(value).toLowerCase() === 'yes' || String(value).toLowerCase() === 'true';
    }
    if (metadata.typeAsString === 'MultiChoice') {
      return toArray(value).map(function(entry: any) { return String(entry); });
    }
    if (metadata.typeAsString === 'Lookup' || metadata.typeAsString === 'LookupMulti'
      || metadata.typeAsString === 'User' || metadata.typeAsString === 'UserMulti') {
      var lookupIds = this.getLookupIds(value);
      return metadata.allowMultiple ? lookupIds : (lookupIds.length > 0 ? lookupIds[0] : '');
    }
    if (metadata.typeAsString === 'DateTime') {
      if (!value) {
        return '';
      }
      var parsedDate = new Date(String(value));
      if (isNaN(parsedDate.getTime())) {
        return '';
      }
      var localDate = new Date(parsedDate.getTime() - (parsedDate.getTimezoneOffset() * 60000));
      var isoValue = localDate.toISOString();
      return metadata.displayFormat === 0 ? isoValue.substring(0, 10)
        : metadata.displayFormat === 2 ? isoValue.substring(11, 16) : isoValue.substring(0, 16);
    }
    if (metadata.typeAsString === 'URL' && value && typeof value === 'object') {
      return String(value.Url || '');
    }
    var normalizedValue = value === undefined || value === null ? '' : this.stringifyCellValue(value);
    return metadata.typeAsString === 'Note' && metadata.richText
      ? decodeRichTextHtml(normalizedValue) : normalizedValue;
  }

  private beginRowEdit(row: any): void {
    var itemId = this.getRowItemId(row);
    if (itemId <= 0 || this.state.saving || this.isReadOnly()) {
      return;
    }
    var values: { [fieldName: string]: any } = {};
    var fields = this.getDisplayFields();
    for (var i = 0; i < fields.length; i += 1) {
      var metadata = this.getGridFieldMetadata(fields[i]);
      if (metadata && this.isEditableGridField(fields[i])) {
        var rawValue = this.getRowFieldValue(row, fields[i]);
        if (metadata.typeAsString === 'Lookup' || metadata.typeAsString === 'LookupMulti'
          || metadata.typeAsString === 'User' || metadata.typeAsString === 'UserMulti') {
          var lookupCompanion = row[metadata.internalName + 'Id'];
          if (lookupCompanion === undefined) { lookupCompanion = row[String(fields[i].Name || '') + '.lookupId']; }
          if (lookupCompanion !== undefined) { rawValue = lookupCompanion; }
        }
        values[metadata.internalName] = this.normalizeEditingValue(rawValue, metadata);
      }
    }
    this.setState({
      selectedItemId: itemId,
      selectedMode: 'edit',
      editingItemId: itemId,
      editingValues: values,
      editingErrors: {},
      error: null
    });
    this.props.onSelectionChange(itemId, 'edit');
  }

  private beginNewRow(): void {
    if (this.state.saving || this.isReadOnly()) {
      return;
    }
    var values: { [fieldName: string]: any } = {};
    var fields = this.getDisplayFields();
    for (var i = 0; i < fields.length; i += 1) {
      var metadata = this.getGridFieldMetadata(fields[i]);
      if (metadata && this.isEditableGridField(fields[i])) {
        var schemaField = this.getGridSchemaField(metadata.internalName);
        values[metadata.internalName] = schemaField && schemaField.defaultValue !== undefined
          ? schemaField.defaultValue
          : metadata.typeAsString === 'Boolean' ? false
            : metadata.typeAsString === 'MultiChoice' ? [] : '';
        if (this.state.runtimeDefaultField && metadata.internalName.toLowerCase() === this.state.runtimeDefaultField.toLowerCase()) {
          values[metadata.internalName] = this.normalizeEditingValue(this.state.runtimeDefaultValue, metadata);
        }
      }
    }
    this.setState({
      selectedItemId: 0,
      selectedMode: 'new',
      editingItemId: 0,
      editingValues: values,
      editingErrors: {},
      error: null
    });
    this.props.onSelectionChange(0, 'new');
  }

  private cancelRowEdit(): void {
    this.setState({
      editingItemId: -1,
      editingValues: {},
      editingErrors: {},
      saving: false,
      selectedMode: 'view'
    });
    this.props.onSelectionChange(this.state.selectedItemId, 'view');
  }

  private updateEditingValue(fieldName: string, value: any): void {
    var values = Object.assign({}, this.state.editingValues);
    var errors = Object.assign({}, this.state.editingErrors);
    values[fieldName] = value;
    delete errors[fieldName];
    this.setState({ editingValues: values, editingErrors: errors });
  }

  private validateEditingValues(): { [fieldName: string]: string } {
    var errors: { [fieldName: string]: string } = {};
    var fields = this.getDisplayFields();
    var schemaFields = this.getGridSchemaFields();
    var validationFields: IGridValidationField[] = [];
    var validationValues: { [fieldId: string]: any } = {};
    for (var schemaIndex = 0; schemaIndex < schemaFields.length; schemaIndex += 1) {
      var validationField = schemaFields[schemaIndex];
      var validationFieldId = validationField.id || validationField.fieldName;
      validationFields.push({ id: validationFieldId, fieldName: validationField.fieldName, label: validationField.label });
      validationValues[validationFieldId] = this.state.editingValues[validationField.fieldName];
    }
    for (var i = 0; i < fields.length; i += 1) {
      var metadata = this.getGridFieldMetadata(fields[i]);
      if (!metadata || !this.isEditableGridField(fields[i])) {
        continue;
      }
      var value = this.state.editingValues[metadata.internalName];
      var isEmpty = value === undefined || value === null || value === ''
        || (Array.isArray(value) && value.length === 0);
      var schemaField = this.getGridSchemaField(metadata.internalName);
      if (metadata.required && isEmpty) {
        errors[metadata.internalName] = schemaField && schemaField.requiredMessage
          ? schemaField.requiredMessage : strings.RuntimeFieldRequired;
        continue;
      }
      if (!schemaField || !Array.isArray(schemaField.validation)) {
        continue;
      }
      for (var ruleIndex = 0; ruleIndex < schemaField.validation.length; ruleIndex += 1) {
        var rule = schemaField.validation[ruleIndex] || {};
        if (rule.applyWhen) {
          try {
            if (!evaluateGridValidationExpression(String(rule.applyWhen), validationFields, validationValues)) { continue; }
          } catch (applyWhenError) {
            errors[metadata.internalName] = applyWhenError && applyWhenError.message
              ? 'Validation condition error: ' + applyWhenError.message : strings.RuntimeFieldInvalid;
            break;
        }
        }
        if (isEmpty && rule.type !== 'required') {
          continue;
        }
        var textValue = String(value);
        var numericValue = Number(value);
        var ruleValue = Number(rule.value);
        var failed = rule.type === 'required' ? isEmpty
          : rule.type === 'minLength' ? textValue.length < ruleValue
            : rule.type === 'maxLength' ? textValue.length > ruleValue
              : rule.type === 'min' ? numericValue < ruleValue
                : rule.type === 'max' ? numericValue > ruleValue : false;
        if (rule.type === 'pattern') {
          try {
            failed = !(new RegExp(String(rule.value || ''))).test(textValue);
          } catch (patternError) {
            errors[metadata.internalName] = patternError && patternError.message
              ? 'Invalid validation pattern: ' + patternError.message : strings.RuntimeFieldInvalid;
            break;
          }
        }
        if (rule.type === 'custom') {
          try {
            failed = !evaluateGridValidationExpression(String(rule.value || ''), validationFields, validationValues);
          } catch (customError) {
            errors[metadata.internalName] = customError && customError.message
              ? 'Validation expression error: ' + customError.message : strings.RuntimeFieldInvalid;
            break;
          }
        }
        if (failed) {
          errors[metadata.internalName] = String(rule.message || strings.RuntimeFieldInvalid);
          break;
        }
      }
    }
    var advancedRules = this.getAdvancedValidationRules();
    for (var advancedIndex = 0; advancedIndex < advancedRules.length; advancedIndex += 1) {
      var advancedRule = advancedRules[advancedIndex];
      if (!advancedRule || !advancedRule.expression || !advancedRule.message) { continue; }
      try {
        if (evaluateGridValidationExpression(advancedRule.expression, validationFields, validationValues)) { continue; }
        var targetFieldName = String(advancedRule.targetField || '').toLowerCase();
        var resolvedTarget = '';
        for (var targetIndex = 0; targetIndex < validationFields.length; targetIndex += 1) {
          var target = validationFields[targetIndex];
          if (String(target.id || '').toLowerCase() === targetFieldName
            || String(target.fieldName || '').toLowerCase() === targetFieldName
            || String(target.label || '').toLowerCase() === targetFieldName) {
            resolvedTarget = target.fieldName;
            break;
          }
        }
        errors[resolvedTarget || '__form'] = advancedRule.message;
      } catch (advancedError) {
        errors['__form'] = advancedError && advancedError.message
          ? 'Validation rule error: ' + advancedError.message : strings.RuntimeFieldInvalid;
      }
    }
    return errors;
  }

  private buildEditingPayload(): any {
    var payload: any = {};
    var fields = this.getDisplayFields();
    for (var i = 0; i < fields.length; i += 1) {
      var metadata = this.getGridFieldMetadata(fields[i]);
      if (!metadata || !this.isEditableGridField(fields[i])) {
        continue;
      }
      var value = this.state.editingValues[metadata.internalName];
      var storageMetadata = this.state.fieldMetadataByName[String(metadata.internalName || '').toLowerCase()] || metadata;
      var storageType = storageMetadata.typeAsString;
      if (storageType === 'Number' || storageType === 'Currency' || storageType === 'Integer') {
        payload[metadata.internalName] = value === '' ? null : Number(value);
      } else if (storageType === 'Boolean') {
        payload[metadata.internalName] = value === true;
      } else if (storageType === 'MultiChoice') {
        payload[metadata.internalName] = Array.isArray(value) ? value : [];
      } else if (storageType === 'Lookup' || storageType === 'LookupMulti' || storageType === 'User' || storageType === 'UserMulti') {
        var isMultipleLookup = storageMetadata.allowMultiple || storageType === 'LookupMulti' || storageType === 'UserMulti';
        payload[metadata.internalName + 'Id'] = isMultipleLookup
          ? (Array.isArray(value) ? value.map(function(entry: any) { return toPositiveInt(entry); }).filter(function(entry: number) { return entry > 0; }) : [])
          : (toPositiveInt(value) || null);
      } else if (storageType === 'DateTime') {
        var dateValue = String(value || '');
        payload[metadata.internalName] = !dateValue ? null
          : metadata.displayFormat === 0 ? new Date(dateValue + 'T00:00:00Z').toISOString()
            : metadata.displayFormat === 2 ? new Date('2000-01-01T' + dateValue + ':00Z').toISOString()
              : new Date(dateValue + 'Z').toISOString();
      } else if (storageType === 'URL' || storageType === 'Hyperlink') {
        payload[metadata.internalName] = value ? { Url: String(value), Description: String(value) } : null;
      } else if (storageType === 'Text' || storageType === 'Note') {
        payload[metadata.internalName] = value === undefined || value === null ? '' : String(value);
      } else {
        payload[metadata.internalName] = value === undefined || value === null ? '' : value;
      }
    }
    return payload;
  }

  private buildVerboseEditingPayload(payload: any, entityTypeName: string): any {
    var verbosePayload: any = {
      __metadata: { type: entityTypeName }
    };
    var payloadKey: string;
    for (payloadKey in payload) {
      if (Object.prototype.hasOwnProperty.call(payload, payloadKey)) {
        verbosePayload[payloadKey] = payload[payloadKey];
      }
    }
    var fields = this.getDisplayFields();
    for (var fieldIndex = 0; fieldIndex < fields.length; fieldIndex += 1) {
      var metadata = this.getGridFieldMetadata(fields[fieldIndex]);
      if (!metadata) { continue; }
      var storageMetadata = this.state.fieldMetadataByName[String(metadata.internalName || '').toLowerCase()] || metadata;
      var storageType = storageMetadata.typeAsString;
      if (storageType === 'MultiChoice' && Array.isArray(payload[metadata.internalName])) {
        verbosePayload[metadata.internalName] = {
          __metadata: { type: 'Collection(Edm.String)' },
          results: payload[metadata.internalName]
        };
      } else if (storageType === 'LookupMulti' || storageType === 'UserMulti' || storageMetadata.allowMultiple) {
        var idProperty = metadata.internalName + 'Id';
        if (Array.isArray(payload[idProperty])) {
          verbosePayload[idProperty] = {
            __metadata: { type: 'Collection(Edm.Int32)' },
            results: payload[idProperty]
          };
        }
      }
    }
    return verbosePayload;
  }

  private async saveEditingRow(): Promise<void> {
    if (this.isReadOnly()) {
      this.cancelRuntimeEdit();
      return;
    }
    var validationErrors = this.validateEditingValues();
    if (Object.keys(validationErrors).length > 0) {
      this.setState({ editingErrors: validationErrors });
      return;
    }

    this.setState({ saving: true, error: null });
    try {
      var listUrl = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items";
      var payload = this.buildEditingPayload();
      var entityTypeName = await this.loadListItemEntityTypeName();
      var verbosePayload = this.buildVerboseEditingPayload(payload, entityTypeName);
      this.logDiagnostic('Saving row. mode=' + (this.state.editingItemId > 0 ? 'edit' : 'new')
        + ', itemId=' + String(this.state.editingItemId) + ', fields=' + Object.keys(payload).join(','));
      var response: any;
      if (this.state.editingItemId > 0) {
        response = await this.postJsonWithFallback(
          listUrl + '(' + this.state.editingItemId + ')',
          payload,
          {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
            'Prefer': 'return-no-content'
          },
          verbosePayload,
          false
        );
      } else {
        response = await this.postJsonWithFallback(
          listUrl,
          payload,
          { 'Prefer': 'return-no-content' },
          verbosePayload,
          false
        );
      }
      if (!response.ok) {
        var responseText = await response.text();
        throw new Error(responseText || strings.RuntimeSaveFailed);
      }
      this.setState({
        editingItemId: -1,
        editingValues: {},
        editingErrors: {},
        saving: false,
        selectedMode: 'view'
      });
      await this.loadRows();
    } catch (error) {
      var saveError: any = error as any;
      this.setState({
        saving: false,
        error: saveError && saveError.message ? saveError.message : strings.RuntimeSaveFailed
      });
    }
  }

  private renderEditingControl(field: IListFieldDefinition): JSX.Element | null {
    var metadata = this.getGridFieldMetadata(field);
    if (!metadata || !this.isEditableGridField(field)) {
      return null;
    }
    var value = this.state.editingValues[metadata.internalName];
    var error = this.state.editingErrors[metadata.internalName];
    var schemaField = this.getGridSchemaField(metadata.internalName);
    var config = schemaField && schemaField.config ? schemaField.config : {};
    var control: JSX.Element;
    if (metadata.typeAsString === 'Boolean') {
      control = <input type="checkbox" checked={value === true} title={metadata.description} onChange={(ev) => this.updateEditingValue(metadata.internalName, ev.currentTarget.checked)} />;
    } else if (metadata.typeAsString === 'Note') {
      control = metadata.richText ? (
        <GridRichTextEditor
          value={String(value || '')}
          title={metadata.description}
          placeholder={String(config.placeholder || '')}
          onChange={(html) => this.updateEditingValue(metadata.internalName, html)}
        />
      ) : (
        <textarea value={String(value || '')} title={metadata.description} placeholder={String(config.placeholder || '')} maxLength={config.maxLength} onChange={(ev) => this.updateEditingValue(metadata.internalName, ev.currentTarget.value)} />
      );
    } else if (metadata.typeAsString === 'Choice') {
      control = (
        <select value={String(value || '')} title={metadata.description} onChange={(ev) => this.updateEditingValue(metadata.internalName, ev.currentTarget.value)}>
          <option value=""></option>
          {metadata.choices.map(function(choice: string) { return <option key={choice} value={choice}>{choice}</option>; })}
        </select>
      );
    } else if (metadata.typeAsString === 'MultiChoice') {
      var selectedChoices = Array.isArray(value) ? value : [];
      control = (
        <div title={metadata.description}>
          {metadata.choices.map((choice: string) => (
            <label key={choice} style={{ display: 'block' }}>
              <input type="checkbox" checked={selectedChoices.indexOf(choice) >= 0} onChange={(ev) => {
                var nextChoices = selectedChoices.slice(0);
                var choiceIndex = nextChoices.indexOf(choice);
                if (ev.currentTarget.checked && choiceIndex < 0) { nextChoices.push(choice); }
                else if (!ev.currentTarget.checked && choiceIndex >= 0) { nextChoices.splice(choiceIndex, 1); }
                this.updateEditingValue(metadata.internalName, nextChoices);
              }} /> {choice}
            </label>
          ))}
        </div>
      );
    } else if (metadata.typeAsString === 'Lookup' || metadata.typeAsString === 'LookupMulti'
      || metadata.typeAsString === 'User' || metadata.typeAsString === 'UserMulti') {
      var lookupOptions = this.state.lookupOptionsByField[metadata.internalName.toLowerCase()] || [];
      if (metadata.allowMultiple) {
        var selectedLookupIds = Array.isArray(value) ? value.map(function(entry: any) { return String(entry); }) : [];
        control = (
          <div title={metadata.description}>
            {lookupOptions.map((option: IGridLookupOption) => (
              <label key={String(option.id)} style={{ display: 'block' }}>
                <input type="checkbox" checked={selectedLookupIds.indexOf(String(option.id)) >= 0} onChange={(ev) => {
                  var nextIds = selectedLookupIds.slice(0);
                  var optionId = String(option.id);
                  var optionIndex = nextIds.indexOf(optionId);
                  if (ev.currentTarget.checked && optionIndex < 0) { nextIds.push(optionId); }
                  else if (!ev.currentTarget.checked && optionIndex >= 0) { nextIds.splice(optionIndex, 1); }
                  this.updateEditingValue(metadata.internalName, nextIds);
                }} /> {option.text}
              </label>
            ))}
          </div>
        );
      } else {
        control = (
          <select value={String(value || '')} title={metadata.description} onChange={(ev) => this.updateEditingValue(metadata.internalName, ev.currentTarget.value)}>
            <option value=""></option>
            {lookupOptions.map(function(option: IGridLookupOption) { return <option key={String(option.id)} value={String(option.id)}>{option.text}</option>; })}
          </select>
        );
      }
    } else {
      var inputType = metadata.typeAsString === 'DateTime' ? (metadata.displayFormat === 0 ? 'date' : metadata.displayFormat === 2 ? 'time' : 'datetime-local')
        : (metadata.typeAsString === 'Number' || metadata.typeAsString === 'Currency' || metadata.typeAsString === 'Integer') ? 'number'
          : metadata.typeAsString === 'URL' ? 'url' : 'text';
      var numberStep = config.decimals !== undefined && Number(config.decimals) > 0
        ? String(1 / Math.pow(10, Number(config.decimals))) : metadata.typeAsString === 'Integer' ? '1' : 'any';
      control = <input type={inputType} value={String(value || '')} title={metadata.description} placeholder={String(config.placeholder || '')} maxLength={config.maxLength} min={config.min} max={config.max} step={inputType === 'number' ? numberStep : undefined} onChange={(ev) => this.updateEditingValue(metadata.internalName, ev.currentTarget.value)} />;
    }
    return (
      <div className={joinClassNames(['gc-cell-editor', error ? 'gc-cell-editor-error' : ''])}>
        {control}
        {error && <div className="gc-field-error">{error}</div>}
      </div>
    );
  }

  private renderEditingRow(
    displayFields: IListFieldDefinition[],
    sourceRow: any,
    fieldsByKey: { [key: string]: IListFieldDefinition },
    conditionalRules: IConditionalStyleRule[]
  ): JSX.Element {
    var isNew = this.state.editingItemId === 0;
    var editingRow = Object.assign({}, sourceRow || {});
    for (var fieldIndex = 0; fieldIndex < displayFields.length; fieldIndex += 1) {
      var editingField = displayFields[fieldIndex];
      var editingMetadata = this.getGridFieldMetadata(editingField);
      if (!editingMetadata || !Object.prototype.hasOwnProperty.call(this.state.editingValues, editingMetadata.internalName)) {
        continue;
      }
      var editingValue = this.state.editingValues[editingMetadata.internalName];
      editingRow[editingMetadata.internalName] = editingValue;
      if (editingField.Name) { editingRow[editingField.Name] = editingValue; }
      if (editingField.RealFieldName) { editingRow[editingField.RealFieldName] = editingValue; }
    }
    var conditionalStyle = this.getConditionalStyleForRow(editingRow, fieldsByKey, conditionalRules);
    var actions = (
      <td className="gc-row-actions">
        <button type="button" className={this.getCommandButtonClass()} disabled={this.state.saving} title={this.state.saving ? strings.RuntimeSaving : strings.RuntimeSave} aria-label={this.state.saving ? strings.RuntimeSaving : strings.RuntimeSave} onClick={() => this.saveEditingRow()}>{this.renderCommandContent('Save', this.state.saving ? strings.RuntimeSaving : strings.RuntimeSave)}</button>
        <button type="button" className={this.getCommandButtonClass()} disabled={this.state.saving} title={strings.RuntimeCancel} aria-label={strings.RuntimeCancel} onClick={() => this.cancelRowEdit()}>{this.renderCommandContent('Cancel', strings.RuntimeCancel)}</button>
        {this.state.editingErrors['__form'] && <div className="gc-field-error">{this.state.editingErrors['__form']}</div>}
      </td>
    );
    return (
      <tr className="lc-row gc-row-editing" onClick={(ev) => ev.stopPropagation()}>
        {this.props.showDelete && !this.isReadOnly() && <td className="gc-selection-cell"></td>}
        {this.props.actionButtonsPosition === 'beginning' && actions}
        {displayFields.map((field) => {
          var editor = this.renderEditingControl(field);
          var fieldKey = this.getFieldKey(field);
          var columnStyle = conditionalStyle.columnStylesByFieldKey[fieldKey] || {};
          var cellStyle = mergeStyleObjects(
            mergeStyleObjects(conditionalStyle.rowStyle, columnStyle),
            this.getConfiguredColumnStyle(field)
          );
          return <td key={field.Name} style={cellStyle}>{editor || (isNew ? null : strings.RuntimeReadOnlyCell)}</td>;
        })}
        {this.props.actionButtonsPosition !== 'beginning' && actions}
      </tr>
    );
  }

  private async openVersionHistory(row: any): Promise<void> {
    var itemId = this.getRowItemId(row);
    if (itemId <= 0 || !this.props.historyAvailable) { return; }

    this.setState({
      historyDialogOpen: true,
      historyDialogUrl: '',
      historyDialogLoading: true,
      historyDialogError: ''
    });

    try {
      var listName = String(this.props.listName || '');
      if (!this._listId || this._listIdListName !== listName) {
        var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')?$select=Id";
        var response = await this.props.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) {
          throw new Error(strings.RuntimeHistoryLoadFailed);
        }
        var data = await response.json();
        var listId = String(data.Id || (data.d && data.d.Id) || '').replace(/[{}]/g, '');
        if (!listId) {
          throw new Error(strings.RuntimeHistoryLoadFailed);
        }
        this._listId = listId;
        this._listIdListName = listName;
      }

      var historyUrl = this.getWebUrl() + '/_layouts/15/Versions.aspx?list='
        + encodeURIComponent('{' + this._listId + '}') + '&ID=' + encodeURIComponent(String(itemId)) + '&IsDlg=1';
      this.setState({ historyDialogUrl: historyUrl, historyDialogLoading: false });
    } catch (error) {
      var historyError: any = error as any;
      this.setState({
        historyDialogLoading: false,
        historyDialogError: historyError && historyError.message ? historyError.message : strings.RuntimeHistoryLoadFailed
      });
    }
  }

  private closeVersionHistory(): void {
    this.setState({
      historyDialogOpen: false,
      historyDialogUrl: '',
      historyDialogLoading: false,
      historyDialogError: ''
    });
  }

  private renderRowActionCell(row: any, isEditingRow: boolean, isReadOnly: boolean): JSX.Element {
    return (
      <td className="gc-row-actions">
        {!isReadOnly && (
          <button
            type="button"
            className={this.getCommandButtonClass()}
            disabled={isEditingRow || this.state.deleting}
            title={strings.RuntimeEdit}
            aria-label={strings.RuntimeEdit}
            onClick={(ev) => {
              ev.stopPropagation();
              this.beginRowEdit(row);
            }}
          >
            {this.renderCommandContent('Edit', strings.RuntimeEdit)}
          </button>
        )}
        {this.props.showHistory !== false && this.props.historyAvailable && (
          <button
            type="button"
            className={this.getCommandButtonClass()}
            disabled={isEditingRow || this.state.deleting}
            title={strings.RuntimeHistory}
            aria-label={strings.RuntimeHistory}
            onClick={(ev) => {
              ev.stopPropagation();
              this.openVersionHistory(row);
            }}
          >
            {this.renderCommandContent('History', strings.RuntimeHistory)}
          </button>
        )}
      </td>
    );
  }

  private async deleteSelected(): Promise<void> {
    if (this.isReadOnly()) {
      return;
    }
    var itemIds = this.state.selectedItemIds.slice(0);
    if (itemIds.length === 0) {
      return;
    }

    if (typeof window !== 'undefined' && !window.confirm(formatString(strings.RuntimeDeleteConfirm, itemIds.length))) {
      return;
    }

    this.logDiagnostic('Deleting items. itemIds=' + itemIds.join(','));
    this.setState({ deleting: true, deleteMessage: '', error: null });

    try {
      var failedItemIds: number[] = [];
      for (var itemIndex = 0; itemIndex < itemIds.length; itemIndex += 1) {
        var itemId = itemIds[itemIndex];
        var endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items(" + itemId + ")";
        try {
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
            failedItemIds.push(itemId);
          }
        } catch (_itemDeleteError) {
          failedItemIds.push(itemId);
        }
      }

      var remainingSelectedItemId = failedItemIds.indexOf(this.state.selectedItemId) >= 0 ? this.state.selectedItemId : 0;
      this.setState({
        selectedItemId: remainingSelectedItemId,
        selectedItemIds: failedItemIds,
        selectedMode: 'view',
        deleting: false
      });
      this.props.onSelectionChange(remainingSelectedItemId, 'view');
      this.logDiagnostic('Bulk delete completed. deleted=' + String(itemIds.length - failedItemIds.length) + ', failed=' + String(failedItemIds.length));
      await this.loadRows();
      if (failedItemIds.length > 0) {
        this.setState({ deleteMessage: formatString(strings.RuntimeDeleteFailedCount, failedItemIds.length) });
      }
    } catch (error) {
      var deleteError: any = error as any;
      this.setState({
        deleting: false,
        deleteMessage: deleteError && deleteError.message ? deleteError.message : strings.RuntimeDeleteFailed,
        error: deleteError && deleteError.message ? deleteError.message : strings.RuntimeDeleteFailed
      });
      this.logDiagnostic('Delete failed: ' + (deleteError && deleteError.message ? deleteError.message : String(deleteError)));
    }
  }

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }

    console.log('[GridControl] ' + message);
  }

  private getCellMarkup(row: any, field: IListFieldDefinition): { __html: string } | null {
    var value = this.getRowFieldValue(row, field);

    if (value === undefined || value === null || value === '') {
      return null;
    }

    var text = this.formatDateCellValue(value, field, row);
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

    if (String(field.TypeAsString || '') === 'Attachments') {
      return String(this.getAttachmentCountForRow(row, value));
    }

    return this.stringifyCellValue(value).replace(/<[^>]*>/g, '').trim();
  }

  private getAttachmentCountForRow(row: any, rawAttachmentsValue: any): number {
    var itemId = row ? this.getRowItemId(row) : 0;
    var knownCount = this.state.attachmentCountsByItemId[itemId];
    if (knownCount !== undefined) {
      return knownCount;
    }
    // Fall back to the boolean-only presence indicator if the exact count hasn't loaded yet.
    return rawAttachmentsValue === true ? 1 : 0;
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

    this._filterAnchorEl = anchorElement;
    var popoverStyle = this.getFilterPopoverStyle(anchorElement);
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

  private getFilterPopoverStyle(anchorElement: HTMLElement): any {
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

    return popoverStyle;
  }

  private updateActiveFilterPosition(): void {
    if (!this.state.activeFilterFieldName || !this._filterAnchorEl || !document.body.contains(this._filterAnchorEl)) { return; }
    var nextStyle = this.getFilterPopoverStyle(this._filterAnchorEl);
    var currentStyle = this.state.filterPopoverStyle || {};
    if (nextStyle.left !== currentStyle.left || nextStyle.top !== currentStyle.top
      || nextStyle.bottom !== currentStyle.bottom || nextStyle.maxHeight !== currentStyle.maxHeight) {
      this.setState({ filterPopoverStyle: nextStyle });
    }
  }

  private closeFilter(): void {
    this._filterAnchorEl = undefined;
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
          <button type="button" className="gc-compact-icon-button" title="Previous month" aria-label="Previous month" onClick={() => this.changeDatePickerMonth(-1)}><i className="ms-Icon ms-Icon--ChevronLeft" aria-hidden="true"></i></button>
          <strong>{new Date(year, monthIndex, 1).toLocaleString(undefined, { month: 'long', year: 'numeric' })}</strong>
          <button type="button" className="gc-compact-icon-button" title="Next month" aria-label="Next month" onClick={() => this.changeDatePickerMonth(1)}><i className="ms-Icon ms-Icon--ChevronRight" aria-hidden="true"></i></button>
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
    var lookupIds: string[] = [];
    var collectLookupIds = function(value: any): void {
      if (value === undefined || value === null || value === '') { return; }
      if (Array.isArray(value)) {
        for (var valueIndex = 0; valueIndex < value.length; valueIndex += 1) { collectLookupIds(value[valueIndex]); }
        return;
      }
      if (typeof value === 'string') {
        var trimmedLookupValue = value.trim();
        if (trimmedLookupValue.length > 0 && (trimmedLookupValue.charAt(0) === '{' || trimmedLookupValue.charAt(0) === '[')) {
          var parsedLookupValue = tryParseObject(trimmedLookupValue);
          if (parsedLookupValue && typeof parsedLookupValue === 'object') {
            collectLookupIds(parsedLookupValue);
            return;
          }
        }
      }
      if (typeof value === 'object') {
        var lookupId = value.LookupId !== undefined ? value.LookupId : value.lookupId !== undefined ? value.lookupId : value.Id !== undefined ? value.Id : value.ID !== undefined ? value.ID : value.id;
        if (lookupId !== undefined && lookupId !== null && String(lookupId).trim()) { lookupIds.push(String(lookupId).trim()); }
      }
    };
    collectLookupIds(this.getRowFieldValue(row, field));
    var fieldNames = [String(field.RealFieldName || ''), String(field.Name || '')];
    for (var fieldIndex = 0; fieldIndex < fieldNames.length; fieldIndex += 1) {
      var fieldName = fieldNames[fieldIndex];
      if (!fieldName) { continue; }
      var companions = [row[fieldName + 'Id'], row[fieldName + '.lookupId']];
      for (var companionIndex = 0; companionIndex < companions.length; companionIndex += 1) {
        var companion = companions[companionIndex];
        if (Array.isArray(companion)) {
          for (var lookupIndex = 0; lookupIndex < companion.length; lookupIndex += 1) { lookupIds.push(String(companion[lookupIndex]).trim()); }
        } else if (companion !== undefined && companion !== null && String(companion).trim()) {
          lookupIds.push(String(companion).trim());
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
      if (!source) { continue; }
      try {
        var parsed = JSON.parse(source);
        if (!Array.isArray(parsed)) { continue; }
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

  private resolveFieldDefinitionByName(fieldName: string): IListFieldDefinition {
    var normalized = String(fieldName || '').toLowerCase();
    for (var i = 0; i < this.state.fields.length; i += 1) {
      var field = this.state.fields[i];
      if (String(field.RealFieldName || field.Name || '').toLowerCase() === normalized
        || String(field.Name || '').toLowerCase() === normalized) {
        return field;
      }
    }
    return { Name: fieldName, DisplayName: fieldName };
  }

  private getGroupValueText(row: any, fieldName: string): string {
    var field = this.resolveFieldDefinitionByName(fieldName);
    var text = this.getCellPlainText(row, field).trim();
    return text || strings.RuntimeGroupNoneValue;
  }

  private buildRowGroups(rows: any[], grouping: IGridGroupingConfig): IGridRowGroup[] {
    var groupsLevel1: IGridRowGroup[] = [];
    var indexByKey1: { [key: string]: number } = {};

    for (var i = 0; i < rows.length; i += 1) {
      var row = rows[i];
      var value1 = this.getGroupValueText(row, grouping.field1);
      var key1 = value1.toLowerCase();
      var groupIndex1 = indexByKey1[key1];
      if (groupIndex1 === undefined) {
        groupIndex1 = groupsLevel1.length;
        indexByKey1[key1] = groupIndex1;
        groupsLevel1.push({ key: key1, label: value1, rows: [], subGroups: [] });
      }
      var group1 = groupsLevel1[groupIndex1];

      if (!grouping.field2) {
        group1.rows.push(row);
        continue;
      }

      var value2 = this.getGroupValueText(row, grouping.field2);
      var key2 = value2.toLowerCase();
      // Must explicitly reset to undefined each iteration: a bare "var subGroup;" without an
      // initializer is a no-op on repeat visits and would keep the previous row's match/creation.
      var subGroup: IGridRowGroup | undefined = undefined;
      for (var s = 0; s < group1.subGroups.length; s += 1) {
        if (group1.subGroups[s].key === key2) { subGroup = group1.subGroups[s]; break; }
      }
      if (!subGroup) {
        subGroup = { key: key2, label: value2, rows: [], subGroups: [] };
        group1.subGroups.push(subGroup);
      }
      subGroup.rows.push(row);
    }

    return groupsLevel1;
  }

  private countGroupRows(group: IGridRowGroup): number {
    if (group.subGroups.length === 0) {
      return group.rows.length;
    }
    var total = 0;
    for (var i = 0; i < group.subGroups.length; i += 1) {
      total += group.subGroups[i].rows.length;
    }
    return total;
  }

  private isGroupCollapsed(groupKey: string, collapsedByDefault: boolean): boolean {
    var override = this.state.collapsedGroupKeys[groupKey];
    return override !== undefined ? override : collapsedByDefault;
  }

  private toggleGroupCollapsed(groupKey: string, collapsedByDefault: boolean): void {
    var next = Object.assign({}, this.state.collapsedGroupKeys);
    next[groupKey] = !this.isGroupCollapsed(groupKey, collapsedByDefault);
    this.setState({ collapsedGroupKeys: next });
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
          borderColor: toCssString(item.borderColor || (item.style && item.style.borderColor) || ''),
          borderStyle: toCssString(item.borderStyle || (item.style && item.style.borderStyle) || ''),
          borderWidth: toCssString(item.borderWidth !== undefined ? item.borderWidth : (item.style && item.style.borderWidth)),
          borderRadius: toCssString(item.borderRadius !== undefined ? item.borderRadius : (item.style && item.style.borderRadius)),
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
    var borderColor = toCssString(styleDefinition.borderColor);
    var borderStyle = toCssString(styleDefinition.borderStyle);
    var borderWidth = toCssString(styleDefinition.borderWidth);
    var borderRadius = toCssString(styleDefinition.borderRadius);
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
    if (borderColor) {
      style.borderColor = borderColor;
    }
    if (borderStyle) {
      style.borderStyle = borderStyle as any;
    }
    if (borderWidth) {
      style.borderWidth = borderWidth + (isNaN(Number(borderWidth)) ? '' : 'px');
    }
    if (borderRadius) {
      style.borderRadius = borderRadius + (isNaN(Number(borderRadius)) ? '' : 'px');
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

  private renderDataRow(
    row: any,
    keySuffix: string,
    isEditingRow: boolean,
    isReadOnly: boolean,
    displayFields: IListFieldDefinition[],
    fieldsByKey: { [key: string]: IListFieldDefinition },
    conditionalRules: IConditionalStyleRule[]
  ): JSX.Element {
    var rowItemId = this.getRowItemId(row);
    var isSelected = rowItemId > 0 && rowItemId === this.state.selectedItemId;
    var showRowActions = !isReadOnly || (this.props.showHistory !== false && this.props.historyAvailable);
    var conditionalStyle = this.getConditionalStyleForRow(row, fieldsByKey, conditionalRules);
    if (this.state.editingItemId === rowItemId) {
      return React.cloneElement(this.renderEditingRow(displayFields, row, fieldsByKey, conditionalRules), { key: String(rowItemId) });
    }
    return (
      <tr
        key={rowItemId > 0 ? String(rowItemId) : keySuffix}
        onClick={() => this.selectRow(row)}
        className={joinClassNames(['lc-row', isSelected ? 'lc-row-selected' : '', this.isItemChecked(rowItemId) ? 'gc-row-delete-selected' : ''])}
      >
        {this.props.showDelete && !isReadOnly && (
          <td className="gc-selection-cell">
            <input
              type="checkbox"
              checked={this.isItemChecked(rowItemId)}
              disabled={this.state.deleting || isEditingRow}
              title={strings.RuntimeSelectRowForDelete}
              aria-label={strings.RuntimeSelectRowForDelete}
              onClick={(ev) => ev.stopPropagation()}
              onChange={() => this.toggleItemChecked(rowItemId)}
            />
          </td>
        )}
        {showRowActions && this.props.actionButtonsPosition === 'beginning' && this.renderRowActionCell(row, isEditingRow, isReadOnly)}
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
            <td key={field.Name} style={mergedCellStyle} title={(this.getGridFieldMetadata(field) || {} as IGridFieldMetadata).description || ''}>
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
        {showRowActions && this.props.actionButtonsPosition !== 'beginning' && this.renderRowActionCell(row, isEditingRow, isReadOnly)}
      </tr>
    );
  }

  private renderGroupHeaderRow(group: IGridRowGroup, keyPath: string, level: number, totalColumnCount: number, groupingConfig: IGridGroupingConfig): JSX.Element {
    var collapsed = this.isGroupCollapsed(keyPath, groupingConfig.collapsedByDefault);
    var count = this.countGroupRows(group);
    var toggleLabel = collapsed ? strings.RuntimeGroupExpand : strings.RuntimeGroupCollapse;
    return (
      <tr key={'group-' + keyPath} className={joinClassNames(['gc-group-row', 'gc-group-row-level-' + String(level)])}>
        <td colSpan={totalColumnCount} style={{ paddingLeft: (10 + level * 20) + 'px' }}>
          <button
            type="button"
            className={this.getCommandButtonClass('gc-group-toggle')}
            title={toggleLabel}
            aria-label={toggleLabel}
            onClick={() => this.toggleGroupCollapsed(keyPath, groupingConfig.collapsedByDefault)}
          >
            {this.renderCommandContent(collapsed ? 'ChevronRight' : 'ChevronDown', toggleLabel)}
          </button>
          <span className="gc-group-label">{group.label}</span>
          {groupingConfig.showCount && <span className="gc-group-count">{formatString(strings.RuntimeGroupItemCount, count)}</span>}
        </td>
      </tr>
    );
  }

  private renderRowGroupRows(
    group: IGridRowGroup,
    keyPath: string,
    level: number,
    groupingConfig: IGridGroupingConfig,
    totalColumnCount: number,
    isEditingRow: boolean,
    isReadOnly: boolean,
    displayFields: IListFieldDefinition[],
    fieldsByKey: { [key: string]: IListFieldDefinition },
    conditionalRules: IConditionalStyleRule[]
  ): JSX.Element[] {
    var rowsOutput: JSX.Element[] = [this.renderGroupHeaderRow(group, keyPath, level, totalColumnCount, groupingConfig)];
    if (this.isGroupCollapsed(keyPath, groupingConfig.collapsedByDefault)) {
      return rowsOutput;
    }
    if (group.subGroups.length > 0) {
      for (var i = 0; i < group.subGroups.length; i += 1) {
        var subGroup = group.subGroups[i];
        var subKeyPath = keyPath + '/' + subGroup.key;
        rowsOutput = rowsOutput.concat(this.renderRowGroupRows(
          subGroup, subKeyPath, level + 1, groupingConfig, totalColumnCount, isEditingRow, isReadOnly, displayFields, fieldsByKey, conditionalRules
        ));
      }
    } else {
      for (var r = 0; r < group.rows.length; r += 1) {
        rowsOutput.push(this.renderDataRow(group.rows[r], keyPath + '-' + String(r), isEditingRow, isReadOnly, displayFields, fieldsByKey, conditionalRules));
      }
    }
    return rowsOutput;
  }

  public render(): JSX.Element {
    if (!this.props.listName) {
      return <div>{strings.RuntimeNoListSelected}</div>;
    }

    if (this.state.accessLoading) {
      return <div>{strings.RuntimeLoading}</div>;
    }

    if (!this.state.hasReadAccess) {
      return <div>{this.props.readAccessDeniedMessage || strings.RuntimeReadAccessDenied}</div>;
    }

    var isEditingRow = this.state.editingItemId >= 0;
    var isReadOnly = this.isReadOnly();
    var processedRows = this.getProcessedRows();
    var groupingConfig = this.getGridGroupingConfig();
    var pageSize = this.props.pageSize > 0 ? this.props.pageSize : 0;
    var pageCount = pageSize > 0 ? Math.max(1, Math.ceil(processedRows.length / pageSize)) : 1;
    var currentPage = Math.min(this.state.currentPage, pageCount - 1);
    var visibleRows = pageSize > 0
      ? processedRows.slice(currentPage * pageSize, (currentPage + 1) * pageSize)
      : processedRows;
    var rowGroups = groupingConfig.enabled ? this.buildRowGroups(visibleRows, groupingConfig) : [];
    var allVisibleItemsChecked = visibleRows.length > 0 && visibleRows.every((row: any) => {
      return this.isItemChecked(this.getRowItemId(row));
    });
    var conditionalRules = this.parseConditionalStyleRules();
    var showRowActions = !isReadOnly || (this.props.showHistory !== false && this.props.historyAvailable);
    var hasActiveFilters = Object.keys(this.state.columnFilters || {}).length > 0;
    var displayFields = this.getDisplayFields();
    var totalColumnCount = displayFields.length
      + (this.props.showDelete && !isReadOnly ? 1 : 0)
      + (showRowActions ? 1 : 0);
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
      '--lc-header-bg-color': this.props.headerBackgroundColor || '#ffffff',
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
      <div className="lc-root gc-root" style={containerStyle}>
        <div className="lc-toolbar">
          {this.props.showViewSelector && (
            <label>
              <span style={{ marginRight: '6px' }}>{strings.RuntimeViewLabel}</span>
              {this.props.showViewAsDropdown !== false ? (
                <select
                  value={this.state.selectedViewId}
                  onChange={(ev) => this.setState({ selectedViewId: ev.currentTarget.value })}
                >
                  {this.props.views.map((view) => {
                    return <option key={view.key} value={view.key}>{view.text}</option>;
                  })}
                </select>
              ) : (
                <span className="gc-view-static-label">{this.getSelectedViewLabel()}</span>
              )}
            </label>
          )}
          {this.props.showRefresh && <button type="button" className={this.getCommandButtonClass()} title={strings.RuntimeRefresh} aria-label={strings.RuntimeRefresh} onClick={() => this.loadRows()}>{this.renderCommandContent('Refresh', strings.RuntimeRefresh)}</button>}
          {this.props.showAdd && !isReadOnly && (
            <button
              type="button"
              className={this.getCommandButtonClass('gc-add-row')}
              disabled={isEditingRow}
              title={strings.RuntimeAddRow}
              aria-label={strings.RuntimeAddRow}
              onClick={() => this.beginNewRow()}
            >
              {this.renderCommandContent('Add', strings.RuntimeAddRow)}
            </button>
          )}
          {this.props.showDelete && !isReadOnly && (
            <button
              type="button"
              className={this.getCommandButtonClass()}
              disabled={this.state.selectedItemIds.length === 0 || this.state.deleting || isEditingRow}
              title={this.state.deleting ? strings.RuntimeDeleting : formatString(strings.RuntimeDeleteSelected, this.state.selectedItemIds.length)}
              aria-label={this.state.deleting ? strings.RuntimeDeleting : formatString(strings.RuntimeDeleteSelected, this.state.selectedItemIds.length)}
              onClick={() => this.deleteSelected()}
            >
              {this.renderCommandContent('Delete', this.state.deleting
                ? strings.RuntimeDeleting
                : formatString(strings.RuntimeDeleteSelected, this.state.selectedItemIds.length))}
            </button>
          )}
        </div>

        {!this.state.hasEditAccess && <div className="lc-status">{this.props.editAccessDeniedMessage || strings.RuntimeEditAccessDenied}</div>}

        {this.props.isEditMode && (
          <div className="lc-status">
            <div>{formatString(strings.RuntimeSelectedItem, this.state.selectedItemId || 0)}</div>
            <div>{formatString(strings.RuntimeSelectedMode, this.state.selectedMode)}</div>
            <div>{formatString(strings.RuntimeConditionalStyleMatches, selectedMatchedRuleNames.length > 0 ? selectedMatchedRuleNames.join(', ') : strings.RuntimeConditionalStyleMatchesNone)}</div>
            <div>{formatString(strings.RuntimeDeleteSelectionCount, this.state.selectedItemIds.length)}</div>
          </div>
        )}

        {this.state.loading && <div>{strings.RuntimeLoading}</div>}
        {!!this.state.deleteMessage && <div className="gc-delete-message">{this.state.deleteMessage}</div>}
        {!!this.state.error && <div>{this.state.error}</div>}
        {!this.state.loading && !this.state.error && processedRows.length === 0 && !isEditingRow && (
          <div>{hasActiveFilters ? strings.RuntimeNoItemsAfterFilter : strings.RuntimeNoItems}</div>
        )}

        {!this.state.loading && !this.state.error && (processedRows.length > 0 || isEditingRow) && (
          <div className="lc-table-container">
            <div className="lc-sticky-header-viewport" ref={this._setStickyHeaderViewportRef}></div>
            <div className="lc-table-wrap" ref={this._setTableWrapRef}>
            {this.state.showScrollArrows && (
              <button
                type="button"
                className="lc-scroll-arrow lc-scroll-arrow-left"
                style={{ top: this.state.scrollArrowTop + 'px', left: (this.state.scrollArrowLeft + 6) + 'px' }}
                title={strings.RuntimeScrollLeft}
                aria-label={strings.RuntimeScrollLeft}
                onClick={() => this.scrollTableHorizontally(-1)}
              >&#8249;</button>
            )}
            {this.state.showScrollArrows && (
              <button
                type="button"
                className="lc-scroll-arrow lc-scroll-arrow-right"
                style={{ top: this.state.scrollArrowTop + 'px', right: (this.state.scrollArrowRight + 6) + 'px' }}
                title={strings.RuntimeScrollRight}
                aria-label={strings.RuntimeScrollRight}
                onClick={() => this.scrollTableHorizontally(1)}
              >&#8250;</button>
            )}
            <table className="lc-table">
              <thead ref={this._setTableHeadRef}>
                <tr>
                  {this.props.showDelete && !isReadOnly && (
                    <th className="gc-selection-header">
                      <input
                        type="checkbox"
                        checked={allVisibleItemsChecked}
                        disabled={visibleRows.length === 0 || this.state.deleting || isEditingRow}
                        title={strings.RuntimeSelectVisibleRows}
                        aria-label={strings.RuntimeSelectVisibleRows}
                        onChange={() => this.toggleVisibleItemsChecked(visibleRows)}
                      />
                    </th>
                  )}
                  {showRowActions && this.props.actionButtonsPosition === 'beginning' && <th className="gc-actions-header">{strings.RuntimeActions}</th>}
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
                      <th key={field.Name} style={this.getConfiguredColumnStyle(field)} title={(this.getGridFieldMetadata(field) || {} as IGridFieldMetadata).description || ''}>
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
                              className={this.getCommandButtonClass(joinClassNames(['lc-header-action', filterActive ? 'lc-header-action-active' : '']))}
                              title={strings.RuntimeFilterTitle}
                              aria-label={strings.RuntimeFilterTitle}
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
                              {this.renderCommandContent('Filter', strings.RuntimeFilterIcon)}
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
                                    <button type="button" className="lc-date-picker-button gc-compact-icon-button" title="Choose from date" aria-label="Choose from date" onClick={() => this.toggleDatePicker('start')}><i className="ms-Icon ms-Icon--Calendar" aria-hidden="true"></i></button>
                                  </div>
                                  <label className="lc-filter-label">{strings.RuntimeFilterEndDateLabel}</label>
                                  <div className="lc-date-input-row">
                                    <input className="lc-filter-input" type="text" placeholder="YYYY-MM-DD" value={this.state.draftFilterEndValue} onChange={(ev) => this.setState({ draftFilterEndValue: ev.currentTarget.value })} />
                                    <button type="button" className="lc-date-picker-button gc-compact-icon-button" title="Choose to date" aria-label="Choose to date" onClick={() => this.toggleDatePicker('end')}><i className="ms-Icon ms-Icon--Calendar" aria-hidden="true"></i></button>
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
                                <button type="button" className={this.getCommandButtonClass()} title={strings.RuntimeFilterApply} aria-label={strings.RuntimeFilterApply} onClick={() => this.applyActiveFilter()}>{this.renderCommandContent('CheckMark', strings.RuntimeFilterApply)}</button>
                                <button type="button" className={this.getCommandButtonClass()} title={strings.RuntimeFilterClear} aria-label={strings.RuntimeFilterClear} onClick={() => this.clearActiveFilter()}>{this.renderCommandContent('ClearFilter', strings.RuntimeFilterClear)}</button>
                                <button type="button" className={this.getCommandButtonClass()} title={strings.RuntimeFilterClose} aria-label={strings.RuntimeFilterClose} onClick={() => this.closeFilter()}>{this.renderCommandContent('Cancel', strings.RuntimeFilterClose)}</button>
                              </div>
                            </div>
                          )}
                        </div>
                      </th>
                    );
                  })}
                  {showRowActions && this.props.actionButtonsPosition !== 'beginning' && <th className="gc-actions-header">{strings.RuntimeActions}</th>}
                </tr>
              </thead>
              <tbody>
                {this.state.editingItemId === 0 && this.renderEditingRow(displayFields, {}, fieldsByKey, conditionalRules)}
                {groupingConfig.enabled
                  ? rowGroups.reduce((accumulated: JSX.Element[], group) => accumulated.concat(this.renderRowGroupRows(
                      group, group.key, 0, groupingConfig, totalColumnCount, isEditingRow, isReadOnly, displayFields, fieldsByKey, conditionalRules
                    )), [] as JSX.Element[])
                  : visibleRows.map((row, index) => this.renderDataRow(row, String(index), isEditingRow, isReadOnly, displayFields, fieldsByKey, conditionalRules))}
              </tbody>
            </table>
            </div>
          </div>
        )}
        {!this.state.loading && !this.state.error && (
          <div className="lc-footer">
            <div className="lc-total-items">{this.state.nextPageHref
              ? strings.RuntimeLoadedItemsLabel.replace('{0}', String(processedRows.length))
              : formatString(strings.RuntimeTotalItemsLabel, processedRows.length)}</div>
            {pageSize > 0 && (pageCount > 1 || !!this.state.nextPageHref) && (
              <div className="lc-pagination">
                <button
                  type="button"
                  className={this.getCommandButtonClass()}
                  disabled={currentPage === 0}
                  title={strings.RuntimePreviousPage}
                  aria-label={strings.RuntimePreviousPage}
                  onClick={() => this.setState({ currentPage: Math.max(0, currentPage - 1) })}
                >
                  {this.renderCommandContent('ChevronLeft', strings.RuntimePreviousPage)}
                </button>
                <span>{(this.state.nextPageHref ? strings.RuntimePageStatusMore : strings.RuntimePageStatus).replace('{0}', String(currentPage + 1)).replace('{1}', String(pageCount)).replace('{2}', String(pageSize))}</span>
                <button
                  type="button"
                  className={this.getCommandButtonClass()}
                  disabled={this.state.loadingMore || (currentPage >= pageCount - 1 && !this.state.nextPageHref)}
                  title={strings.RuntimeNextPage}
                  aria-label={strings.RuntimeNextPage}
                  onClick={() => this.goToNextPage(currentPage, pageCount)}
                >
                  {this.renderCommandContent('ChevronRight', strings.RuntimeNextPage)}
                </button>
                {this.state.loadingMore && <span>{strings.RuntimeLoadingMore}</span>}
              </div>
            )}
          </div>
        )}
        {this.state.historyDialogOpen && (
          <div
            className="gc-history-overlay"
            role="presentation"
            onClick={() => this.closeVersionHistory()}
            onKeyDown={(ev) => {
              if (ev.key === 'Escape') { this.closeVersionHistory(); }
            }}
          >
            <div className="gc-history-dialog" role="dialog" aria-modal="true" aria-label={strings.RuntimeHistoryTitle} onClick={(ev) => ev.stopPropagation()}>
              <div className="gc-history-header">
                <div className="gc-history-title">{strings.RuntimeHistoryTitle}</div>
                <button type="button" className="gc-history-close" title={strings.RuntimeClose} aria-label={strings.RuntimeClose} onClick={() => this.closeVersionHistory()}>
                  <i className="ms-Icon ms-Icon--Cancel" aria-hidden="true"></i>
                </button>
              </div>
              <div className="gc-history-content">
                {this.state.historyDialogLoading && <div className="gc-history-status">{strings.RuntimeLoading}</div>}
                {!!this.state.historyDialogError && <div className="gc-history-status gc-history-error">{this.state.historyDialogError}</div>}
                {!!this.state.historyDialogUrl && <iframe className="gc-history-frame" src={this.state.historyDialogUrl} title={strings.RuntimeHistoryTitle}></iframe>}
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }
}