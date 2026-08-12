import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { PropertyPaneCustomField } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldCustomList, CustomListFieldType } from 'sp-client-custom-fields/lib/PropertyFieldCustomList';

import * as strings from 'ListControlWebPartStrings';
import { ListControl, IListControlColumnConfiguration, IListControlViewOption } from './components/ListControl';

var packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface IDynamicDataPropertyDefinitionCompat {
  id: string;
  title: string;
  description?: string;
}

export interface IDynamicDataSourceMetadataCompat {
  title: string;
  description?: string;
  alias?: string;
  componentId?: string;
  instanceId?: string;
}

export interface IListControlWebPartProps {
  description: string;
  instanceName?: string;
  listName: string;
  viewId: string;
  viewColumns?: IListControlColumnConfiguration[];
  pageSize?: number | string;
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
  filterDesignerField?: string;
  filterDesignerOperator?: 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le' | 'contains' | 'startswith' | 'endswith' | 'notcontains';
  filterDesignerLogical?: 'and' | 'or';
  filterDesignerValueType?: 'static' | 'expression';
  filterDesignerValue?: string;
  filterDesignerSelectedIndex?: string;
  filterDesignerLookupPick?: string;
  conditionalStyleJson?: string;
  conditionalStyleRuleName?: string;
  conditionalStyleScope?: 'row' | 'column';
  conditionalStyleApplyField?: string;
  conditionalStyleConditionField?: string;
  conditionalStyleConditionOperator?: 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le' | 'contains' | 'startswith' | 'endswith' | 'notcontains';
  conditionalStyleConditionLogical?: 'and' | 'or';
  conditionalStyleConditionValueType?: 'static' | 'expression';
  conditionalStyleConditionValue?: string;
  conditionalStyleEnabled?: boolean;
  conditionalStylePriority?: string;
  conditionalStyleBackgroundColor?: string;
  conditionalStyleForegroundColor?: string;
  conditionalStyleFontFamily?: string;
  conditionalStyleFontSize?: string;
  conditionalStyleFontStyle?: string;
  conditionalStyleFontWeight?: string;
  conditionalStyleTextAlign?: string;
  conditionalStyleSelectedIndex?: string;
  conditionalStyleLookupPick?: string;
}

export interface IDropdownOption {
  key: string | number;
  text: string;
}

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

function normalizeQueryParamName(value: string | undefined, fallback: string): string {
  var normalized = String(value || '').trim();
  if (!normalized || normalized.toLowerCase() === 'id') {
    return fallback;
  }

  return normalized;
}

function normalizePageSize(value: number | string | undefined): number {
  var parsed = parseInt(String(value === undefined || value === null ? '' : value), 10);
  return !isNaN(parsed) && parsed > 0 ? parsed : 0;
}

function normalizeStringCollection(value: any): string[] {
  var source = value && value.results ? value.results : value;
  if (!Array.isArray(source)) {
    return [];
  }

  return source.map(function(item: any) {
    return String(item && item.Name !== undefined ? item.Name : item || '');
  }).filter(function(item: string) {
    return !!item;
  });
}

export default class ListControlWebPart extends BaseClientSideWebPart<IListControlWebPartProps> {
  private _lists: IDropdownOption[] = [];
  private _sitePages: IDropdownOption[] = [];
  private _views: IListControlViewOption[] = [];
  private _listFields: IDropdownOption[] = [];
  private _selectedItemId: number = 0;
  private _selectedMode: string = 'view';
  private _dynamicDataSourceManager: any;
  private _filterJsonValidationMessage: string = '';
  private _filterDesignerMessage: string = '';
  private _showFilterExpressionHelp: boolean = false;
  private _conditionalStyleJsonValidationMessage: string = '';
  private _conditionalStyleDesignerMessage: string = '';
  private _showConditionalStyleExpressionHelp: boolean = false;
  private _isEditingConditionalStyle: boolean = false;
  private _conditionalStyleDesignerRevision: number = 0;
  private _fieldTypeByInternalName: { [internalName: string]: string } = {};
  private _fieldLookupListByInternalName: { [internalName: string]: string } = {};
  private _filterLookupItemOptions: IDropdownOption[] = [];
  private _conditionalStyleLookupItemOptions: IDropdownOption[] = [];
  private _filterLookupMessage: string = '';
  private _conditionalStyleLookupMessage: string = '';

  public constructor() {
    super();
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }

  public get id(): string {
    return this.context.instanceId;
  }

  public get metadata(): IDynamicDataSourceMetadataCompat {
    var instanceName = String(this.properties.instanceName || '').trim();
    return {
      title: instanceName || strings.DynamicSourceTitle,
      description: strings.DynamicSourceDescription,
      alias: this.context.manifest.alias,
      componentId: this.context.manifest.id,
      instanceId: this.context.instanceId,
    };
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinitionCompat> {
    return [
      {
        id: 'instanceName',
        title: strings.PropInstanceNameLabel,
        description: strings.PropInstanceNameDescription,
      },
      {
        id: 'selectedItemId',
        title: strings.DynamicPropertySelectedItemIdTitle,
        description: strings.DynamicPropertySelectedItemIdDescription,
      },
      {
        id: 'selectedMode',
        title: strings.DynamicPropertySelectedModeTitle,
        description: strings.DynamicPropertySelectedModeDescription,
      }
    ];
  }

  public getPropertyValue(propertyId: string): any {
    if (propertyId === 'instanceName') {
      return String(this.properties.instanceName || '').trim();
    }

    if (propertyId === 'selectedItemId') {
      return this._selectedItemId;
    }

    if (propertyId === 'selectedMode') {
      return this._selectedMode;
    }

    throw new Error('Bad property id');
  }

  public render(): void {
    const element: React.ReactElement<any> = React.createElement(ListControl, {
      context: this.context,
      listName: this.properties.listName || '',
      defaultViewId: this.properties.viewId || '',
      views: this._views,
      viewColumns: this.properties.viewColumns || [],
      pageSize: normalizePageSize(this.properties.pageSize),
      isEditMode: this.displayMode === DisplayMode.Edit,
      showViewSelector: this.properties.showViewSelector !== false,
      showRefresh: this.properties.showRefresh !== false,
      showAdd: this.properties.showAdd !== false,
      showEdit: this.properties.showEdit !== false,
      showView: this.properties.showView !== false,
      showDelete: this.properties.showDelete !== false,
      showLinkToItem: this.properties.showLinkToItem === true,
      linkTargetPageUrl: this.properties.linkTargetPageUrl === '__defaultForm__'
        ? '' : (this.properties.linkTargetPageUrl || ''),
      linkTargetIdParam: this.properties.linkTargetIdParam || 'itemid',
      includeReturnUrlParam: this.properties.includeReturnUrlParam === true,
      enableDiagnostics: this.properties.enableDiagnostics !== false,
      bodyTextColor: this.properties.bodyTextColor || '',
      bodyFontFamily: this.properties.bodyFontFamily || '',
      bodyFontSize: this.properties.bodyFontSize || '14px',
      bodyFontStyle: this.properties.bodyFontStyle || 'normal',
      bodyFontBold: this.properties.bodyFontBold === true,
      bodyTextAlign: this.properties.bodyTextAlign || 'left',
      selectedTextColor: this.properties.selectedTextColor || '',
      selectedBackgroundColor: this.properties.selectedBackgroundColor || '',
      selectedFontStyle: this.properties.selectedFontStyle || 'normal',
      selectedFontBold: this.properties.selectedFontBold === true,
      headerTextColor: this.properties.headerTextColor || '',
      headerBackgroundColor: this.properties.headerBackgroundColor || '',
      headerFontFamily: this.properties.headerFontFamily || '',
      headerFontSize: this.properties.headerFontSize || '14px',
      headerFontStyle: this.properties.headerFontStyle || 'normal',
      headerFontBold: this.properties.headerFontBold === true,
      headerTextAlign: this.properties.headerTextAlign || 'left',
      tableBackgroundColor: this.properties.tableBackgroundColor || '#ffffff',
      tableBorderColor: this.properties.tableBorderColor || '#cccccc',
      tableBorderWidth: this.properties.tableBorderWidth || '1px',
      tableCornerStyle: this.properties.tableCornerStyle || 'square',
      tableCornerRadius: typeof this.properties.tableCornerRadius === 'number' ? this.properties.tableCornerRadius : 0,
      tableRowLineWidth: typeof this.properties.tableRowLineWidth === 'number' ? this.properties.tableRowLineWidth : 1,
      cornerStyle: this.properties.cornerStyle || 'square',
      cornerRadius: typeof this.properties.cornerRadius === 'number' ? this.properties.cornerRadius : 12,
      alternateRowShading: this.properties.alternateRowShading === true,
      alternateRowShadingColor: this.properties.alternateRowShadingColor || '#f9f9f9',
      buttonTextColor: this.properties.buttonTextColor || '#000000',
      buttonBackgroundColor: this.properties.buttonBackgroundColor || '#f0f0f0',
      buttonFontFamily: this.properties.buttonFontFamily || 'inherit',
      buttonFontSize: this.properties.buttonFontSize || '14px',
      buttonFontStyle: this.properties.buttonFontStyle || 'normal',
      buttonFontBold: this.properties.buttonFontBold === true,
      buttonCornerStyle: this.properties.buttonCornerStyle || 'square',
      buttonCornerRadius: typeof this.properties.buttonCornerRadius === 'number' ? this.properties.buttonCornerRadius : 4,
      webpartBackgroundColor: this.properties.webpartBackgroundColor || 'transparent',
      webpartBorderColor: this.properties.webpartBorderColor || '#ccc',
      webpartBorderWidth: this.properties.webpartBorderWidth || 0,
      filterJson: this.properties.filterJson || '',
      conditionalStyleJson: this.properties.conditionalStyleJson || '',
      onSelectionChange: (itemId: number, mode: string) => this.handleSelectionChange(itemId, mode)
    });

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.logDiagnostic('onInit started. listName=' + String(this.properties.listName || '(none)'));
    this.properties.linkTargetIdParam = normalizeQueryParamName(this.properties.linkTargetIdParam, 'itemid');
    this.initializeDynamicDataSource();

    return Promise.all([this.loadLists(), this.loadSitePages()]).then(() => {
      this.logDiagnostic('List metadata loaded. Count=' + String(this._lists.length));
      if (this.properties.listName) {
        this.logDiagnostic('Loading views for configured list: ' + String(this.properties.listName));
        return this.loadViews(this.properties.listName).then(() => {
          return this.loadListFields(this.properties.listName).then(() => {
            if ((!this.properties.viewColumns || this.properties.viewColumns.length === 0) && this.properties.viewId) {
              return this.loadViewColumns(this.properties.listName, this.properties.viewId);
            }
            return Promise.resolve();
          });
        });
      }
      return Promise.resolve();
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.logDiagnostic('Property pane opened. listName=' + String(this.properties.listName || '(none)'));
    if (this._lists.length === 0) {
      this.loadLists();
    }
    if (this._sitePages.length === 0) {
      this.loadSitePages();
    }
    if (this.properties.listName && this._views.length === 0) {
      this.loadViews(this.properties.listName);
    }
    if (this.properties.listName && this._listFields.length === 0) {
      this.loadListFields(this.properties.listName);
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
    if (propertyPath === 'instanceName' && oldValue !== newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.notifyDynamicData('instanceName');
      this.notifyDynamicSourceChanged();
      return;
    }

    if (propertyPath === 'listName' && oldValue !== newValue) {
      this.properties.viewId = '';
      this.properties.viewColumns = [];
      this.properties.filterDesignerField = '';
      this.properties.conditionalStyleApplyField = '';
      this.properties.conditionalStyleConditionField = '';
      this._views = [];
      this._listFields = [];
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      if (newValue) {
        this.loadViews(String(newValue));
        this.loadListFields(String(newValue));
      }
      return;
    }

    if (propertyPath === 'viewId' && oldValue !== newValue) {
      this.properties.viewColumns = [];
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
      this.render();
      if (this.properties.listName && newValue) {
        this.loadViewColumns(this.properties.listName, String(newValue));
      }
      return;
    }

    if (
      propertyPath === 'filterDesignerField'
      || propertyPath === 'filterDesignerOperator'
      || propertyPath === 'filterDesignerLogical'
      || propertyPath === 'filterDesignerValueType'
      || propertyPath === 'filterDesignerSelectedIndex'
      || propertyPath === 'filterDesignerLookupPick'
      || propertyPath === 'conditionalStyleScope'
      || propertyPath === 'conditionalStyleApplyField'
      || propertyPath === 'conditionalStyleConditionField'
      || propertyPath === 'conditionalStyleConditionOperator'
      || propertyPath === 'conditionalStyleConditionLogical'
      || propertyPath === 'conditionalStyleConditionValueType'
      || propertyPath === 'conditionalStyleEnabled'
      || propertyPath === 'conditionalStylePriority'
      || propertyPath === 'conditionalStyleSelectedIndex'
      || propertyPath === 'conditionalStyleLookupPick'
    ) {
      if (propertyPath === 'filterDesignerField') {
        this._filterLookupItemOptions = [];
        this.properties.filterDesignerLookupPick = '';
        this._filterLookupMessage = '';
        if (this.isLookupTypeField(String(newValue || ''))) {
          this.handleLoadFilterLookupItems();
        }
      }
      if (propertyPath === 'conditionalStyleConditionField') {
        this._conditionalStyleLookupItemOptions = [];
        this.properties.conditionalStyleLookupPick = '';
        this._conditionalStyleLookupMessage = '';
        if (this.isLookupTypeField(String(newValue || ''))) {
          this.handleLoadConditionalStyleLookupItems();
        }
      }
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
      return;
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private async loadListFields(listName: string): Promise<void> {
    if (!listName) {
      this._listFields = [];
      this._fieldTypeByInternalName = {};
      this._fieldLookupListByInternalName = {};
      this.context.propertyPane.refresh();
      return;
    }

    try {
      var endpoint = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '')
        + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/fields?$select=InternalName,Title,TypeAsString,LookupList,Hidden,ReadOnlyField,Sealed&$filter=Hidden eq false and ((ReadOnlyField eq false and Sealed eq false) or InternalName eq 'ID')";
      var data = await this.getJsonWithAcceptFallback(endpoint);
      var fields = data && data.value ? data.value : [];
      if (!fields || fields.length === 0) {
        fields = data && data.d && data.d.results ? data.d.results : [];
      }

      this._fieldTypeByInternalName = {};
      this._fieldLookupListByInternalName = {};
      fields.forEach((field: any) => {
        var fieldInternalName = String(field.InternalName || '');
        if (!fieldInternalName) { return; }
        this._fieldTypeByInternalName[fieldInternalName] = String(field.TypeAsString || '').toLowerCase();
        var lookupListId = field.LookupList ? String(field.LookupList).replace(/[{}]/g, '') : '';
        if (lookupListId) {
          this._fieldLookupListByInternalName[fieldInternalName] = lookupListId;
        }
      });

      this._listFields = fields.map(function(field: any) {
        var internalName = String(field.InternalName || '');
        var title = String(field.Title || internalName);
        return {
          key: internalName,
          text: title + (title !== internalName ? ' (' + internalName + ')' : '')
        };
      }).sort(function(a: IDropdownOption, b: IDropdownOption) {
        return String(a.text).localeCompare(String(b.text));
      });

      this.context.propertyPane.refresh();
    } catch (error) {
      this._listFields = [];
      this._fieldTypeByInternalName = {};
      this._fieldLookupListByInternalName = {};
      this.context.propertyPane.refresh();
      this.logDiagnostic('Failed to load fields for list "' + listName + '": ' + (error && error.message ? error.message : String(error)));
    }
  }

  private isLookupTypeField(internalName: string): boolean {
    var typeName = this._fieldTypeByInternalName[internalName || ''];
    return typeName === 'lookup' || typeName === 'lookupmulti';
  }

  private async loadLookupListItems(listId: string): Promise<IDropdownOption[]> {
    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      var url = webUrl + "/_api/web/lists(guid'" + listId + "')/items?$select=Id,Title&$top=500&$orderby=Title";
      var data = await this.getJsonWithAcceptFallback(url);
      var items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      return items.map(function(item: any) {
        var title = item.Title ? String(item.Title) : '(no title)';
        return { key: title, text: title + ' (ID: ' + item.Id + ')' };
      });
    } catch (error) {
      this.logDiagnostic('Failed to load lookup list items for listId=' + listId + ': ' + (error && error.message ? error.message : String(error)));
      return [];
    }
  }

  private async handleLoadFilterLookupItems(): Promise<void> {
    var fieldName = String(this.properties.filterDesignerField || '');
    var listId = this._fieldLookupListByInternalName[fieldName];
    if (!listId) {
      this._filterLookupMessage = 'Selected field is not a lookup column, or its target list could not be resolved.';
      this._filterLookupItemOptions = [];
      this.context.propertyPane.refresh();
      return;
    }

    this._filterLookupItemOptions = await this.loadLookupListItems(listId);
    this._filterLookupMessage = this._filterLookupItemOptions.length > 0
      ? 'Loaded ' + this._filterLookupItemOptions.length + ' item(s). Pick one and click Use selected value.'
      : 'The target list has no items.';
    this.context.propertyPane.refresh();
  }

  private handleApplyFilterLookupPick(): void {
    var picked = String(this.properties.filterDesignerLookupPick || '');
    if (!picked) {
      this._filterLookupMessage = 'Select an item from the dropdown before applying.';
      this.context.propertyPane.refresh();
      return;
    }

    this.properties.filterDesignerValue = picked;
    this._filterLookupMessage = 'Applied "' + picked + '" to the filter value.';
    this.context.propertyPane.refresh();
  }

  private async handleLoadConditionalStyleLookupItems(): Promise<void> {
    var fieldName = String(this.properties.conditionalStyleConditionField || '');
    var listId = this._fieldLookupListByInternalName[fieldName];
    if (!listId) {
      this._conditionalStyleLookupMessage = 'Selected field is not a lookup column, or its target list could not be resolved.';
      this._conditionalStyleLookupItemOptions = [];
      this.context.propertyPane.refresh();
      return;
    }

    this._conditionalStyleLookupItemOptions = await this.loadLookupListItems(listId);
    this._conditionalStyleLookupMessage = this._conditionalStyleLookupItemOptions.length > 0
      ? 'Loaded ' + this._conditionalStyleLookupItemOptions.length + ' item(s). Pick one and click Use selected value.'
      : 'The target list has no items.';
    this.context.propertyPane.refresh();
  }

  private handleApplyConditionalStyleLookupPick(): void {
    var picked = String(this.properties.conditionalStyleLookupPick || '');
    if (!picked) {
      this._conditionalStyleLookupMessage = 'Select an item from the results before applying.';
      this.context.propertyPane.refresh();
      return;
    }

    this.properties.conditionalStyleConditionValue = picked;
    this._conditionalStyleLookupMessage = 'Applied "' + picked + '" to the condition value.';
    this.context.propertyPane.refresh();
  }

  private createPropertyPanePageTitle(key: string, title: string): any {
    return PropertyPaneCustomField({
      key: key,
      onRender: function(domElement: HTMLElement): void {
        domElement.textContent = title;
        domElement.style.fontSize = '20px';
        domElement.style.fontWeight = '600';
        domElement.style.lineHeight = '28px';
        domElement.style.marginBottom = '12px';
      }
    });
  }

  private createFilterExpressionHelpField(): any {
    var isExpanded = this._showFilterExpressionHelp;
    return PropertyPaneCustomField({
      key: 'filterExpressionHelp-' + String(isExpanded),
      onRender: (domElement: HTMLElement): void => {
        domElement.innerHTML = '';

        var button = document.createElement('button');
        button.type = 'button';
        button.title = isExpanded ? strings.PropFilterDesignerExpressionHelpHide : strings.PropFilterDesignerExpressionHelpShow;
        button.setAttribute('aria-label', button.title);
        button.setAttribute('aria-expanded', String(isExpanded));
        button.style.background = 'transparent';
        button.style.border = '0';
        button.style.color = '#0078d4';
        button.style.cursor = 'pointer';
        button.style.padding = '2px';

        var icon = document.createElement('span');
        icon.className = 'ms-Icon ms-Icon--Info';
        icon.setAttribute('aria-hidden', 'true');
        button.appendChild(icon);
        button.onclick = () => {
          this._showFilterExpressionHelp = !this._showFilterExpressionHelp;
          this.context.propertyPane.refresh();
        };
        domElement.appendChild(button);

        if (isExpanded) {
          var help = document.createElement('div');
          help.textContent = strings.PropFilterDesignerExpressionHelp;
          help.style.fontSize = '12px';
          help.style.lineHeight = '18px';
          help.style.marginTop = '4px';
          domElement.appendChild(help);
        }
      }
    });
  }

  private createConditionalStyleExpressionHelpField(): any {
    var isExpanded = this._showConditionalStyleExpressionHelp;
    return PropertyPaneCustomField({
      key: 'conditionalStyleExpressionHelp-' + String(isExpanded),
      onRender: (domElement: HTMLElement): void => {
        domElement.innerHTML = '';
        var button = document.createElement('button');
        button.type = 'button';
        button.title = isExpanded ? strings.PropFilterDesignerExpressionHelpHide : strings.PropFilterDesignerExpressionHelpShow;
        button.setAttribute('aria-label', button.title);
        button.setAttribute('aria-expanded', String(isExpanded));
        button.style.background = 'transparent';
        button.style.border = '0';
        button.style.color = '#0078d4';
        button.style.cursor = 'pointer';
        button.style.padding = '2px';
        var icon = document.createElement('span');
        icon.className = 'ms-Icon ms-Icon--Info';
        icon.setAttribute('aria-hidden', 'true');
        button.appendChild(icon);
        button.onclick = () => {
          this._showConditionalStyleExpressionHelp = !this._showConditionalStyleExpressionHelp;
          this.context.propertyPane.refresh();
        };
        domElement.appendChild(button);
        if (isExpanded) {
          var help = document.createElement('div');
          help.textContent = strings.PropFilterDesignerExpressionHelp;
          help.style.fontSize = '12px';
          help.style.lineHeight = '18px';
          help.style.marginTop = '4px';
          domElement.appendChild(help);
        }
      }
    });
  }

  private handleColorPropertyChange(propertyPath: string, oldValue: any, newValue: any): void {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.context.propertyPane.refresh();
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    var fontFamilyOptions = this.getFontFamilyOptions();
    var fontStyleOptions = this.getFontStyleOptions();
    var fullVersionLabel = 'Version: ' + this.getWebPartVersion();
    var filterFieldOptions = this._listFields.length > 0 ? this._listFields : [{ key: '', text: strings.PropFilterDesignerFieldNone }];
    var filterConditionOptions: IDropdownOption[] = [{ key: '', text: strings.PropFilterDesignerExistingNone }];
    var conditionalStyleFieldOptions = this._listFields.length > 0 ? this._listFields : [{ key: '', text: strings.PropConditionalStyleFieldNone }];
    var conditionalStyleConditionOptions: IDropdownOption[] = [{ key: '', text: strings.PropConditionalStyleExistingNone }];
    var conditionalStyleRuleSummaryText = strings.PropConditionalStyleSummaryNone;

    try {
      var existingConditions = this.parseFilterJsonArray(this.properties.filterJson);
      if (existingConditions.length > 0) {
        filterConditionOptions = existingConditions.map(function(condition: any, index: number) {
          var field = String((condition && condition.field) || '(field)');
          var operator = String((condition && condition.operator) || 'contains').toLowerCase();
          var logical = String((condition && condition.logical) || 'and').toUpperCase();
          var valueType = String((condition && condition.valueType) || 'static').toLowerCase();
          var value = condition && condition.value !== undefined ? String(condition.value) : '';
          var logicalPrefix = index === 0 ? '' : logical + ' ';
          return {
            key: String(index),
            text: '#' + String(index + 1) + ': ' + logicalPrefix + field + ' ' + operator + ' ' + (valueType === 'expression' ? '[' + value + ']' : value)
          };
        });
      }
    } catch (_parseFilterError) {
      filterConditionOptions = [{ key: '', text: strings.PropFilterDesignerExistingNone }];
    }

    try {
      var existingConditionalStyles = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
      if (existingConditionalStyles.length > 0) {
        conditionalStyleConditionOptions = existingConditionalStyles.map(function(condition: any, index: number) {
          var rule = String((condition && condition.rule) || '(rule)');
          var enabled = condition && condition.enabled !== false;
          var priority = parseInt(String(condition && condition.priority !== undefined ? condition.priority : '100'), 10);
          var scope = String((condition && condition.scope) || 'row').toLowerCase();
          var applyField = String((condition && condition.applyField) || '');
          var field = String((condition && condition.conditionField) || '(field)');
          var operator = String((condition && condition.operator) || 'contains').toLowerCase();
          var logical = String((condition && condition.logical) || 'and').toUpperCase();
          var valueType = String((condition && condition.valueType) || 'static').toLowerCase();
          var value = condition && condition.value !== undefined ? String(condition.value) : '';
          var scopeInfo = scope === 'column' ? ' -> ' + (applyField || '(column)') : ' -> row';
          var statusInfo = enabled ? 'enabled' : 'disabled';
          var priorityInfo = isNaN(priority) ? 'P100' : ('P' + String(priority));
          var logicalPrefix = index === 0 ? '' : logical + ' ';
          return {
            key: String(index),
            text: '#' + String(index + 1) + ': [' + priorityInfo + '|' + statusInfo + '] ' + rule + scopeInfo + ' | ' + logicalPrefix + field + ' ' + operator + ' ' + (valueType === 'expression' ? '[' + value + ']' : value)
          };
        });

        var groupedSummaries: { [ruleName: string]: { name: string; enabled: boolean; priority: number; scope: string; applyField: string; count: number; firstIndex: number; } } = {};
        var groupedOrder: string[] = [];
        for (var summaryIndex = 0; summaryIndex < existingConditionalStyles.length; summaryIndex += 1) {
          var summaryCondition = existingConditionalStyles[summaryIndex] || {};
          var summaryRuleName = String(summaryCondition.rule || '(rule)').trim() || '(rule)';
          if (!groupedSummaries[summaryRuleName]) {
            groupedSummaries[summaryRuleName] = {
              name: summaryRuleName,
              enabled: summaryCondition.enabled !== false,
              priority: parseInt(String(summaryCondition.priority !== undefined ? summaryCondition.priority : '100'), 10),
              scope: String(summaryCondition.scope || 'row').toLowerCase(),
              applyField: String(summaryCondition.applyField || ''),
              count: 0,
              firstIndex: summaryIndex
            };
            groupedOrder.push(summaryRuleName);
          }

          groupedSummaries[summaryRuleName].count += 1;
        }

        var groupedItems = groupedOrder.map(function(ruleName: string) { return groupedSummaries[ruleName]; });
        groupedItems.sort(function(left: any, right: any) {
          var leftPriority = isNaN(left.priority) ? 100 : left.priority;
          var rightPriority = isNaN(right.priority) ? 100 : right.priority;
          if (leftPriority !== rightPriority) {
            return leftPriority - rightPriority;
          }

          return left.firstIndex - right.firstIndex;
        });

        conditionalStyleRuleSummaryText = groupedItems.map(function(item: any) {
          var enabledText = item.enabled ? strings.PropConditionalStyleSummaryEnabled : strings.PropConditionalStyleSummaryDisabled;
          var priorityText = strings.PropConditionalStyleSummaryPriorityPrefix + String(isNaN(item.priority) ? 100 : item.priority);
          var scopeText = item.scope === 'column'
            ? strings.PropConditionalStyleSummaryScopeColumn + ' ' + (item.applyField || '(column)')
            : strings.PropConditionalStyleSummaryScopeRow;
          var conditionText = strings.PropConditionalStyleSummaryConditionCountPrefix + String(item.count);
          return item.name + ' | ' + enabledText + ' | ' + priorityText + ' | ' + scopeText + ' | ' + conditionText;
        }).join('\n');
      }
    } catch (_parseConditionalStyleError) {
      conditionalStyleConditionOptions = [{ key: '', text: strings.PropConditionalStyleExistingNone }];
      conditionalStyleRuleSummaryText = strings.PropConditionalStyleSummaryInvalid;
    }

    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: fullVersionLabel,
              groupFields: [
                PropertyPaneLabel('propertyPaneVersionInfo', {
                  text: ' '
                })
              ]
            },
            {
              groupName: strings.PropertyPanePageGuideTitle,
              groupFields: [
                PropertyPaneLabel('propertyPanePageGuideFilters', {
                  text: strings.PropertyPanePageGuideFilters
                }),
                PropertyPaneLabel('propertyPanePageGuideStyles', {
                  text: strings.PropertyPanePageGuideStyles
                })
              ]
            },
            {
              groupName: strings.PropertyGroupData,
              groupFields: [
                PropertyPaneTextField('instanceName', {
                  label: strings.PropInstanceNameLabel,
                  description: strings.PropInstanceNameDescription,
                  placeholder: strings.DynamicSourceTitle,
                  value: this.properties.instanceName || ''
                }),
                PropertyPaneDropdown('listName', {
                  label: strings.PropListLabel,
                  options: this._lists,
                  selectedKey: this.properties.listName
                }),
                PropertyPaneDropdown('viewId', {
                  label: strings.PropViewLabel,
                  options: this._views.length > 0 ? this._views : [{ key: '', text: strings.PropNoViews }],
                  selectedKey: this.properties.viewId || ''
                }),
                PropertyFieldCustomList('viewColumns', {
                  key: 'listControlViewColumns-' + String(this.properties.viewId || 'none'),
                  label: strings.PropViewColumnsLabel,
                  headerText: strings.PropViewColumnsHeader,
                  value: this.properties.viewColumns || [],
                  context: this.context,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  fields: [
                    { id: 'fieldName', title: strings.PropViewColumnField, required: true, type: CustomListFieldType.string },
                    { id: 'displayName', title: strings.PropViewColumnDisplayName, required: true, type: CustomListFieldType.string },
                    { id: 'width', title: strings.PropViewColumnWidth, required: false, type: CustomListFieldType.string }
                  ]
                }),
                PropertyPaneTextField('pageSize', {
                  label: strings.PropPageSizeLabel,
                  description: strings.PropPageSizeDescription,
                  value: String(this.properties.pageSize || '')
                }),
                PropertyPaneCheckbox('showViewSelector', {
                  text: strings.PropShowViewSelectorLabel,
                  checked: this.properties.showViewSelector !== false
                }),
                PropertyPaneCheckbox('showLinkToItem', {
                  text: strings.PropShowLinkToItemLabel,
                  checked: this.properties.showLinkToItem === true
                }),
                ...(this.properties.showLinkToItem === true ? [
                  PropertyPaneDropdown('linkTargetPageUrl', {
                    label: strings.PropLinkTargetPageUrlLabel,
                    options: this.getTargetPageOptions(),
                    selectedKey: this.properties.linkTargetPageUrl || '__defaultForm__'
                  }),
                  PropertyPaneTextField('linkTargetIdParam', {
                    label: strings.PropLinkTargetIdParamLabel,
                    placeholder: strings.PropLinkTargetIdParamPlaceholder,
                    value: this.properties.linkTargetIdParam || 'itemid'
                  }),
                  PropertyPaneCheckbox('includeReturnUrlParam', {
                    text: strings.PropIncludeReturnUrlParamLabel,
                    checked: this.properties.includeReturnUrlParam === true
                  })
                ] : [])
              ]
            },
            {
              groupName: strings.PropertyGroupButtons,
              groupFields: [
                PropertyPaneCheckbox('showRefresh', {
                  text: strings.PropShowRefreshLabel,
                  checked: this.properties.showRefresh !== false
                }),
                PropertyPaneCheckbox('showAdd', {
                  text: strings.PropShowAddLabel,
                  checked: this.properties.showAdd !== false
                }),
                PropertyPaneCheckbox('showEdit', {
                  text: strings.PropShowEditLabel,
                  checked: this.properties.showEdit !== false
                }),
                PropertyPaneCheckbox('showView', {
                  text: strings.PropShowViewLabel,
                  checked: this.properties.showView !== false
                }),
                PropertyPaneCheckbox('showDelete', {
                  text: strings.PropShowDeleteLabel,
                  checked: this.properties.showDelete !== false
                })
              ]
            },
            {
              groupName: strings.PropertyGroupBodyStyle,
              groupFields: [
                PropertyFieldColorPicker('bodyTextColor', {
                  label: strings.PropBodyTextColorLabel,
                  selectedColor: this.properties.bodyTextColor || '',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'bodyTextColorField'
                }),
                PropertyPaneDropdown('bodyFontFamily', {
                  label: strings.PropBodyFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.bodyFontFamily || ''
                }),
                PropertyPaneDropdown('bodyFontSize', {
                  label: strings.PropBodyFontSizeLabel,
                  options: [
                    { key: '12px', text: '12px' },
                    { key: '14px', text: '14px' },
                    { key: '16px', text: '16px' },
                    { key: '18px', text: '18px' },
                    { key: '20px', text: '20px' }
                  ],
                  selectedKey: this.properties.bodyFontSize || '14px'
                }),
                PropertyPaneDropdown('bodyFontStyle', {
                  label: strings.PropBodyFontStyleLabel,
                  options: fontStyleOptions,
                  selectedKey: this.properties.bodyFontStyle || 'normal'
                }),
                PropertyPaneDropdown('bodyTextAlign', {
                  label: strings.PropBodyTextAlignLabel,
                  options: this.getTextAlignOptions(),
                  selectedKey: this.properties.bodyTextAlign || 'left'
                }),
                PropertyPaneCheckbox('bodyFontBold', {
                  text: strings.PropBodyBoldLabel,
                  checked: this.properties.bodyFontBold === true
                })
              ]
            },
            {
              groupName: strings.PropertyGroupSelectedStyle,
              groupFields: [
                PropertyFieldColorPicker('selectedTextColor', {
                  label: strings.PropSelectedTextColorLabel,
                  selectedColor: this.properties.selectedTextColor || '',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'selectedTextColorField'
                }),
                PropertyFieldColorPicker('selectedBackgroundColor', {
                  label: strings.PropSelectedBackgroundColorLabel,
                  selectedColor: this.properties.selectedBackgroundColor || '#eef6ff',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'selectedBackgroundColorField'
                }),
                PropertyPaneDropdown('selectedFontStyle', {
                  label: strings.PropSelectedFontStyleLabel,
                  options: fontStyleOptions,
                  selectedKey: this.properties.selectedFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('selectedFontBold', {
                  text: strings.PropSelectedBoldLabel,
                  checked: this.properties.selectedFontBold === true
                })
              ]
            },
            {
              groupName: strings.PropertyGroupHeaderStyle,
              groupFields: [
                PropertyFieldColorPicker('headerTextColor', {
                  label: strings.PropHeaderTextColorLabel,
                  selectedColor: this.properties.headerTextColor || '',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'headerTextColorField'
                }),
                PropertyFieldColorPicker('headerBackgroundColor', {
                  label: strings.PropHeaderBackgroundColorLabel,
                  selectedColor: this.properties.headerBackgroundColor || '',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'headerBackgroundColorField'
                }),
                PropertyPaneDropdown('headerFontFamily', {
                  label: strings.PropHeaderFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.headerFontFamily || ''
                }),
                PropertyPaneDropdown('headerFontSize', {
                  label: strings.PropHeaderFontSizeLabel,
                  options: [
                    { key: '12px', text: '12px' },
                    { key: '14px', text: '14px' },
                    { key: '16px', text: '16px' },
                    { key: '18px', text: '18px' },
                    { key: '20px', text: '20px' }
                  ],
                  selectedKey: this.properties.headerFontSize || '14px'
                }),
                PropertyPaneDropdown('headerFontStyle', {
                  label: strings.PropHeaderFontStyleLabel,
                  options: fontStyleOptions,
                  selectedKey: this.properties.headerFontStyle || 'normal'
                }),
                PropertyPaneDropdown('headerTextAlign', {
                  label: strings.PropHeaderTextAlignLabel,
                  options: this.getTextAlignOptions(),
                  selectedKey: this.properties.headerTextAlign || 'left'
                }),
                PropertyPaneCheckbox('headerFontBold', {
                  text: strings.PropHeaderBoldLabel,
                  checked: this.properties.headerFontBold === true
                })
              ]
            },
            {
              groupName: strings.PropertyGroupTableStyle,
              groupFields: [
                PropertyFieldColorPicker('tableBackgroundColor', {
                  label: strings.PropTableBackgroundColorLabel,
                  selectedColor: this.properties.tableBackgroundColor || '#ffffff',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'tableBackgroundColorField'
                }),
                PropertyFieldColorPicker('tableBorderColor', {
                  label: strings.PropTableBorderColorLabel,
                  selectedColor: this.properties.tableBorderColor || '#cccccc',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'tableBorderColorField'
                }),
                PropertyPaneDropdown('tableBorderWidth', {
                  label: strings.PropTableBorderWidthLabel,
                  options: [
                    { key: '0px', text: 'None' },
                    { key: '1px', text: '1px' },
                    { key: '2px', text: '2px' },
                    { key: '3px', text: '3px' },
                    { key: '4px', text: '4px' },
                    { key: '5px', text: '5px' }
                  ],
                  selectedKey: this.properties.tableBorderWidth || '1px'
                }),
                PropertyPaneSlider('tableRowLineWidth', {
                  label: strings.PropTableRowLineWidthLabel,
                  min: 0,
                  max: 10,
                  step: 1,
                  value: typeof this.properties.tableRowLineWidth === 'number' ? this.properties.tableRowLineWidth : 1
                }),
                PropertyPaneDropdown('tableCornerStyle', {
                  label: strings.PropTableCornerStyleLabel,
                  options: [
                    { key: 'square', text: strings.PropCornerStyleSquare },
                    { key: 'rounded', text: strings.PropCornerStyleRounded }
                  ],
                  selectedKey: this.properties.tableCornerStyle || 'square'
                }),
                ...((this.properties.tableCornerStyle || 'square') === 'rounded' ? [
                  PropertyPaneSlider('tableCornerRadius', {
                    label: strings.PropTableCornerRadiusLabel,
                    min: 0,
                    max: 40,
                    step: 1,
                    value: typeof this.properties.tableCornerRadius === 'number' ? this.properties.tableCornerRadius : 8
                  })
                ] : []),
                PropertyPaneCheckbox('alternateRowShading', {
                  text: strings.PropAlternateRowShadingLabel,
                  checked: this.properties.alternateRowShading === true
                }),
                PropertyFieldColorPicker('alternateRowShadingColor', {
                  label: strings.PropAlternateRowShadingColorLabel,
                  selectedColor: this.properties.alternateRowShadingColor || '#f9f9f9',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'alternateRowShadingColorField'
                })
              ]
            },
            {              groupName: strings.PropertyGroupWebpartStyle,
              groupFields: [
                PropertyFieldColorPicker('webpartBackgroundColor', {
                  label: strings.PropWebpartBackgroundColorLabel,
                  selectedColor: this.properties.webpartBackgroundColor || '#ffffff',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'webpartBackgroundColorField'
                }),
                PropertyFieldColorPicker('webpartBorderColor', {
                  label: strings.PropWebpartBorderColorLabel,
                  selectedColor: this.properties.webpartBorderColor || '#ccc',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'webpartBorderColorField'
                }),
                PropertyPaneSlider('webpartBorderWidth', {
                  label: strings.PropWebpartBorderWidthLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: this.properties.webpartBorderWidth || 0
                }),
                PropertyPaneDropdown('cornerStyle', {
                  label: strings.PropCornerStyleLabel,
                  options: [
                    { key: 'square', text: strings.PropCornerStyleSquare },
                    { key: 'rounded', text: strings.PropCornerStyleRounded }
                  ],
                  selectedKey: this.properties.cornerStyle || 'square'
                }),
                ...((this.properties.cornerStyle || 'square') === 'rounded' ? [
                  PropertyPaneSlider('cornerRadius', {
                    label: strings.PropCornerRadiusLabel,
                    min: 0,
                    max: 40,
                    step: 1,
                    value: typeof this.properties.cornerRadius === 'number' ? this.properties.cornerRadius : 12
                  })
                ] : [])
              ]
            },
            {              groupName: strings.PropertyGroupButtonStyle,
              groupFields: [
                PropertyFieldColorPicker('buttonTextColor', {
                  label: strings.PropButtonTextColorLabel,
                  selectedColor: this.properties.buttonTextColor || '#000000',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'buttonTextColorField'
                }),
                PropertyFieldColorPicker('buttonBackgroundColor', {
                  label: strings.PropButtonBackgroundColorLabel,
                  selectedColor: this.properties.buttonBackgroundColor || '#f0f0f0',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'buttonBackgroundColorField'
                }),
                PropertyPaneDropdown('buttonFontFamily', {
                  label: strings.PropButtonFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.buttonFontFamily || ''
                }),
                PropertyPaneDropdown('buttonFontSize', {
                  label: strings.PropButtonFontSizeLabel,
                  options: [
                    { key: '12px', text: '12px' },
                    { key: '14px', text: '14px' },
                    { key: '16px', text: '16px' },
                    { key: '18px', text: '18px' },
                    { key: '20px', text: '20px' }
                  ],
                  selectedKey: this.properties.buttonFontSize || '14px'
                }),
                PropertyPaneDropdown('buttonFontStyle', {
                  label: strings.PropButtonFontStyleLabel,
                  options: fontStyleOptions,
                  selectedKey: this.properties.buttonFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('buttonFontBold', {
                  text: strings.PropButtonBoldLabel,
                  checked: this.properties.buttonFontBold === true
                }),
                PropertyPaneDropdown('buttonCornerStyle', {
                  label: strings.PropButtonCornerStyleLabel,
                  options: [
                    { key: 'square', text: strings.PropCornerStyleSquare },
                    { key: 'rounded', text: strings.PropCornerStyleRounded }
                  ],
                  selectedKey: this.properties.buttonCornerStyle || 'square'
                }),
                ...((this.properties.buttonCornerStyle || 'square') === 'rounded' ? [
                  PropertyPaneSlider('buttonCornerRadius', {
                    label: strings.PropButtonCornerRadiusLabel,
                    min: 0,
                    max: 40,
                    step: 1,
                    value: typeof this.properties.buttonCornerRadius === 'number' ? this.properties.buttonCornerRadius : 4
                  })
                ] : [])
              ]
            },
            {
              groupName: 'Diagnostics',
              groupFields: [
                PropertyPaneCheckbox('enableDiagnostics', {
                  text: strings.PropEnableDiagnosticsLabel,
                  checked: this.properties.enableDiagnostics !== false
                })
              ]
            },
            {
              groupName: '',
              groupFields: [
                PropertyPaneLabel('propertyPaneBottomSpacer', {
                  text: ' '
                })
              ]
            }
          ]
        },
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: '',
              groupFields: [
                this.createPropertyPanePageTitle('presetFilterPageTitle', strings.PresetFilterPageTitle)
              ]
            },
            {
              groupName: strings.PresetFilterDesignerGroupName,
              groupFields: [
                PropertyPaneDropdown('filterDesignerField', {
                  label: strings.PropFilterDesignerFieldLabel,
                  options: filterFieldOptions,
                  selectedKey: this.properties.filterDesignerField || ''
                }),
                PropertyPaneDropdown('filterDesignerOperator', {
                  label: strings.PropFilterDesignerOperatorLabel,
                  options: [
                    { key: 'eq', text: strings.RuntimeFilterOperatorEquals },
                    { key: 'ne', text: strings.RuntimeFilterOperatorNotEquals },
                    { key: 'contains', text: strings.RuntimeFilterOperatorContains },
                    { key: 'notcontains', text: strings.RuntimeFilterOperatorNotContains },
                    { key: 'startswith', text: strings.RuntimeFilterOperatorStartsWith },
                    { key: 'endswith', text: strings.RuntimeFilterOperatorEndsWith },
                    { key: 'gt', text: strings.RuntimeFilterOperatorGreaterThan },
                    { key: 'ge', text: strings.RuntimeFilterOperatorGreaterThanOrEqual },
                    { key: 'lt', text: strings.RuntimeFilterOperatorLessThan },
                    { key: 'le', text: strings.RuntimeFilterOperatorLessThanOrEqual }
                  ],
                  selectedKey: this.properties.filterDesignerOperator || 'contains'
                }),
                PropertyPaneDropdown('filterDesignerLogical', {
                  label: strings.PropFilterDesignerLogicalLabel,
                  options: [
                    { key: 'and', text: strings.PropFilterDesignerLogicalAnd },
                    { key: 'or', text: strings.PropFilterDesignerLogicalOr }
                  ],
                  selectedKey: this.properties.filterDesignerLogical || 'and'
                }),
                PropertyPaneDropdown('filterDesignerValueType', {
                  label: strings.PropFilterDesignerValueTypeLabel,
                  options: [
                    { key: 'static', text: strings.PropFilterDesignerValueTypeStatic },
                    { key: 'expression', text: strings.PropFilterDesignerValueTypeExpression }
                  ],
                  selectedKey: this.properties.filterDesignerValueType || 'static'
                }),
                PropertyPaneTextField('filterDesignerValue', {
                  label: this.properties.filterDesignerValueType === 'expression' ? strings.PropFilterDesignerExpressionLabel : strings.PropFilterDesignerValueLabel,
                  placeholder: this.properties.filterDesignerValueType === 'expression' ? strings.PropFilterDesignerExpressionPlaceholder : strings.PropFilterDesignerValuePlaceholder,
                  value: this.properties.filterDesignerValue || ''
                }),
                ...(this.properties.filterDesignerValueType === 'expression' ? [
                  this.createFilterExpressionHelpField()
                ] : []),
                ...(this.properties.filterDesignerValueType !== 'expression' && this.isLookupTypeField(this.properties.filterDesignerField || '') ? [
                  PropertyPaneLabel('filterLookupHelperTitle', {
                    text: 'Lookup item picker: choose an item from the target list to fill in its value.'
                  }),
                  PropertyPaneButton('loadFilterLookupItems', {
                    text: 'Refresh lookup items',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleLoadFilterLookupItems.bind(this)
                  }),
                  PropertyPaneDropdown('filterDesignerLookupPick', {
                    label: 'Select item',
                    options: this._filterLookupItemOptions.length > 0 ? this._filterLookupItemOptions : [{ key: '', text: 'No items loaded yet' }],
                    selectedKey: this.properties.filterDesignerLookupPick || ''
                  }),
                  PropertyPaneButton('applyFilterLookupPick', {
                    text: 'Use selected value',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleApplyFilterLookupPick.bind(this)
                  }),
                  PropertyPaneLabel('filterLookupHelperResult', {
                    text: this._filterLookupMessage || 'Lookup items load automatically when you select the column.'
                  }),
                ] : []),
                PropertyPaneButton('addFilterDesignerCondition', {
                  text: strings.PropFilterDesignerAddButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.handleAddFilterDesignerCondition.bind(this)
                }),
                PropertyPaneDropdown('filterDesignerSelectedIndex', {
                  label: strings.PropFilterDesignerExistingLabel,
                  options: filterConditionOptions,
                  selectedKey: this.properties.filterDesignerSelectedIndex || ''
                }),
                PropertyPaneButton('loadFilterDesignerCondition', {
                  text: strings.PropFilterDesignerLoadButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleLoadSelectedFilterCondition.bind(this)
                }),
                PropertyPaneButton('updateFilterDesignerCondition', {
                  text: strings.PropFilterDesignerUpdateButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleUpdateSelectedFilterCondition.bind(this)
                }),
                PropertyPaneButton('removeFilterDesignerCondition', {
                  text: strings.PropFilterDesignerRemoveButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleRemoveSelectedFilterCondition.bind(this)
                }),
                PropertyPaneButton('resetFilterDesignerJson', {
                  text: strings.PropFilterDesignerResetButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleResetFilterJson.bind(this)
                }),
                PropertyPaneLabel('filterDesignerResult', {
                  text: this._filterDesignerMessage || strings.PropFilterDesignerStatusPlaceholder
                })
              ]
            },
            {
              groupName: strings.PresetFilterJsonGroupName,
              groupFields: [
                PropertyPaneTextField('filterJson', {
                  label: strings.PropFilterJsonLabel,
                  placeholder: strings.PropFilterJsonPlaceholder,
                  value: this.properties.filterJson || '',
                  multiline: true,
                  resizable: true,
                  rows: 10,
                  description: strings.PropFilterJsonHelp
                }),
                PropertyPaneButton('validateFilterJson', {
                  text: strings.PropValidateFilterJsonButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleValidateFilterJson.bind(this)
                }),
                PropertyPaneLabel('validateFilterJsonResult', {
                  text: this._filterJsonValidationMessage || strings.JsonValidationResultPlaceholder
                })
              ]
            }
          ]
        },
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: '',
              groupFields: [
                this.createPropertyPanePageTitle('conditionalStylePageTitle', strings.ConditionalStylePageTitle)
              ]
            },
            {
              groupName: strings.ConditionalStyleSummaryGroupName,
              groupFields: [
                PropertyPaneLabel('conditionalStyleRuleSummaryLabel', {
                  text: conditionalStyleRuleSummaryText
                })
              ]
            },
            {
              groupName: strings.ConditionalStyleDesignerGroupName,
              groupFields: [
                PropertyPaneTextField('conditionalStyleRuleName', {
                  label: strings.PropConditionalStyleRuleNameLabel,
                  placeholder: strings.PropConditionalStyleRuleNamePlaceholder,
                  value: this.properties.conditionalStyleRuleName || ''
                }),
                PropertyPaneDropdown('conditionalStyleScope', {
                  label: strings.PropConditionalStyleScopeLabel,
                  options: [
                    { key: 'row', text: strings.PropConditionalStyleScopeRow },
                    { key: 'column', text: strings.PropConditionalStyleScopeColumn }
                  ],
                  selectedKey: this.properties.conditionalStyleScope || 'row'
                }),
                ...(this.properties.conditionalStyleScope === 'column' ? [
                  PropertyPaneDropdown('conditionalStyleApplyField', {
                    label: strings.PropConditionalStyleApplyFieldLabel,
                    options: conditionalStyleFieldOptions,
                    selectedKey: this.properties.conditionalStyleApplyField || ''
                  })
                ] : []),
                PropertyPaneDropdown('conditionalStyleConditionField', {
                  label: strings.PropConditionalStyleConditionFieldLabel,
                  options: conditionalStyleFieldOptions,
                  selectedKey: this.properties.conditionalStyleConditionField || ''
                }),
                PropertyPaneDropdown('conditionalStyleConditionOperator', {
                  label: strings.PropConditionalStyleConditionOperatorLabel,
                  options: [
                    { key: 'eq', text: strings.RuntimeFilterOperatorEquals },
                    { key: 'ne', text: strings.RuntimeFilterOperatorNotEquals },
                    { key: 'contains', text: strings.RuntimeFilterOperatorContains },
                    { key: 'notcontains', text: strings.RuntimeFilterOperatorNotContains },
                    { key: 'startswith', text: strings.RuntimeFilterOperatorStartsWith },
                    { key: 'endswith', text: strings.RuntimeFilterOperatorEndsWith },
                    { key: 'gt', text: strings.RuntimeFilterOperatorGreaterThan },
                    { key: 'ge', text: strings.RuntimeFilterOperatorGreaterThanOrEqual },
                    { key: 'lt', text: strings.RuntimeFilterOperatorLessThan },
                    { key: 'le', text: strings.RuntimeFilterOperatorLessThanOrEqual }
                  ],
                  selectedKey: this.properties.conditionalStyleConditionOperator || 'eq'
                }),
                PropertyPaneDropdown('conditionalStyleConditionLogical', {
                  label: strings.PropConditionalStyleConditionLogicalLabel,
                  options: [
                    { key: 'and', text: strings.PropFilterDesignerLogicalAnd },
                    { key: 'or', text: strings.PropFilterDesignerLogicalOr }
                  ],
                  selectedKey: this.properties.conditionalStyleConditionLogical || 'and'
                }),
                PropertyPaneCheckbox('conditionalStyleEnabled', {
                  text: strings.PropConditionalStyleEnabledLabel,
                  checked: this.properties.conditionalStyleEnabled !== false
                }),
                PropertyPaneTextField('conditionalStylePriority', {
                  label: strings.PropConditionalStylePriorityLabel,
                  placeholder: strings.PropConditionalStylePriorityPlaceholder,
                  value: this.properties.conditionalStylePriority || '100'
                }),
                PropertyPaneDropdown('conditionalStyleConditionValueType', {
                  label: strings.PropFilterDesignerValueTypeLabel,
                  options: [
                    { key: 'static', text: strings.PropFilterDesignerValueTypeStatic },
                    { key: 'expression', text: strings.PropFilterDesignerValueTypeExpression }
                  ],
                  selectedKey: this.properties.conditionalStyleConditionValueType || 'static'
                }),
                PropertyPaneTextField('conditionalStyleConditionValue', {
                  label: this.properties.conditionalStyleConditionValueType === 'expression' ? strings.PropFilterDesignerExpressionLabel : strings.PropConditionalStyleConditionValueLabel,
                  placeholder: this.properties.conditionalStyleConditionValueType === 'expression' ? strings.PropFilterDesignerExpressionPlaceholder : strings.PropConditionalStyleConditionValuePlaceholder,
                  value: this.properties.conditionalStyleConditionValue || ''
                }),
                ...(this.properties.conditionalStyleConditionValueType === 'expression' ? [
                  this.createConditionalStyleExpressionHelpField()
                ] : []),
                ...(this.properties.conditionalStyleConditionValueType !== 'expression' && this.isLookupTypeField(this.properties.conditionalStyleConditionField || '') ? [
                  PropertyPaneLabel('conditionalStyleLookupHelperTitle', {
                    text: 'Lookup item picker: choose an item from the target list to fill in its value.'
                  }),
                  PropertyPaneButton('loadConditionalStyleLookupItems', {
                    text: 'Refresh lookup items',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleLoadConditionalStyleLookupItems.bind(this)
                  }),
                  PropertyPaneDropdown('conditionalStyleLookupPick', {
                    label: 'Select item',
                    options: this._conditionalStyleLookupItemOptions.length > 0 ? this._conditionalStyleLookupItemOptions : [{ key: '', text: 'No items loaded yet' }],
                    selectedKey: this.properties.conditionalStyleLookupPick || ''
                  }),
                  PropertyPaneButton('applyConditionalStyleLookupPick', {
                    text: 'Use selected value',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleApplyConditionalStyleLookupPick.bind(this)
                  }),
                  PropertyPaneLabel('conditionalStyleLookupHelperResult', {
                    text: this._conditionalStyleLookupMessage || 'Lookup items load automatically when you select the column.'
                  }),
                ] : []),
                PropertyFieldColorPicker('conditionalStyleBackgroundColor', {
                  label: strings.PropConditionalStyleBackgroundColorLabel,
                  selectedColor: this.properties.conditionalStyleBackgroundColor || '',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'conditionalStyleBackgroundColorField-' + String(this._conditionalStyleDesignerRevision)
                }),
                PropertyFieldColorPicker('conditionalStyleForegroundColor', {
                  label: strings.PropConditionalStyleForegroundColorLabel,
                  selectedColor: this.properties.conditionalStyleForegroundColor || '',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'conditionalStyleForegroundColorField-' + String(this._conditionalStyleDesignerRevision)
                }),
                PropertyPaneDropdown('conditionalStyleFontFamily', {
                  label: strings.PropConditionalStyleFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.conditionalStyleFontFamily || ''
                }),
                PropertyPaneDropdown('conditionalStyleFontSize', {
                  label: strings.PropConditionalStyleFontSizeLabel,
                  options: [
                    { key: '', text: 'Default' },
                    { key: '12px', text: '12px' },
                    { key: '14px', text: '14px' },
                    { key: '16px', text: '16px' },
                    { key: '18px', text: '18px' },
                    { key: '20px', text: '20px' }
                  ],
                  selectedKey: this.properties.conditionalStyleFontSize || ''
                }),
                PropertyPaneDropdown('conditionalStyleFontStyle', {
                  label: strings.PropConditionalStyleFontStyleLabel,
                  options: [
                    { key: '', text: 'Default' },
                    { key: 'normal', text: strings.PropFontStyleNormal },
                    { key: 'italic', text: strings.PropFontStyleItalic },
                    { key: 'oblique', text: strings.PropFontStyleOblique }
                  ],
                  selectedKey: this.properties.conditionalStyleFontStyle || ''
                }),
                PropertyPaneDropdown('conditionalStyleFontWeight', {
                  label: strings.PropConditionalStyleFontWeightLabel,
                  options: [
                    { key: '', text: 'Default' },
                    { key: 'normal', text: 'Normal' },
                    { key: 'bold', text: 'Bold' }
                  ],
                  selectedKey: this.properties.conditionalStyleFontWeight || ''
                }),
                PropertyPaneDropdown('conditionalStyleTextAlign', {
                  label: strings.PropConditionalStyleTextAlignLabel,
                  options: [
                    { key: '', text: 'Default' },
                    { key: 'left', text: strings.PropTextAlignLeft },
                    { key: 'center', text: strings.PropTextAlignCenter },
                    { key: 'right', text: strings.PropTextAlignRight }
                  ],
                  selectedKey: this.properties.conditionalStyleTextAlign || ''
                }),
                PropertyPaneButton('addConditionalStyleCondition', {
                  text: strings.PropConditionalStyleAddButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  disabled: this._isEditingConditionalStyle,
                  onClick: this.handleAddConditionalStyleCondition.bind(this)
                }),
                ...(this._isEditingConditionalStyle ? [
                  PropertyPaneButton('newConditionalStyle', {
                    text: strings.PropConditionalStyleNewButtonLabel,
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleNewConditionalStyle.bind(this)
                  })
                ] : []),
                PropertyPaneDropdown('conditionalStyleSelectedIndex', {
                  label: strings.PropConditionalStyleExistingLabel,
                  options: conditionalStyleConditionOptions,
                  selectedKey: this.properties.conditionalStyleSelectedIndex || ''
                }),
                PropertyPaneButton('loadConditionalStyleCondition', {
                  text: strings.PropConditionalStyleLoadButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleLoadSelectedConditionalStyleCondition.bind(this)
                }),
                PropertyPaneButton('updateConditionalStyleCondition', {
                  text: strings.PropConditionalStyleUpdateButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  disabled: !this._isEditingConditionalStyle,
                  onClick: this.handleUpdateSelectedConditionalStyleCondition.bind(this)
                }),
                PropertyPaneButton('removeConditionalStyleCondition', {
                  text: strings.PropConditionalStyleRemoveButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleRemoveSelectedConditionalStyleCondition.bind(this)
                }),
                PropertyPaneButton('moveConditionalStyleConditionUp', {
                  text: strings.PropConditionalStyleMoveUpButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleMoveConditionalStyleConditionUp.bind(this)
                }),
                PropertyPaneButton('moveConditionalStyleConditionDown', {
                  text: strings.PropConditionalStyleMoveDownButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleMoveConditionalStyleConditionDown.bind(this)
                }),
                PropertyPaneButton('resetConditionalStyleJson', {
                  text: strings.PropConditionalStyleResetButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  disabled: this._isEditingConditionalStyle,
                  onClick: this.handleResetConditionalStyleJson.bind(this)
                }),
                PropertyPaneLabel('conditionalStyleDesignerResult', {
                  text: this._conditionalStyleDesignerMessage || strings.PropConditionalStyleStatusPlaceholder
                })
              ]
            },
            {
              groupName: strings.ConditionalStyleJsonGroupName,
              groupFields: [
                PropertyPaneTextField('conditionalStyleJson', {
                  label: strings.PropConditionalStyleJsonLabel,
                  placeholder: strings.PropConditionalStyleJsonPlaceholder,
                  value: this.properties.conditionalStyleJson || '',
                  multiline: true,
                  resizable: true,
                  rows: 12,
                  description: strings.PropConditionalStyleJsonHelp
                }),
                PropertyPaneButton('validateConditionalStyleJson', {
                  text: strings.PropValidateConditionalStyleJsonButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleValidateConditionalStyleJson.bind(this)
                }),
                PropertyPaneLabel('validateConditionalStyleJsonResult', {
                  text: this._conditionalStyleJsonValidationMessage || strings.ConditionalStyleJsonValidationResultPlaceholder
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private getWebPartVersion(): string {
    var solutionVersion = packageSolutionConfig && packageSolutionConfig.solution
      ? String(packageSolutionConfig.solution.version || '')
      : '';
    if (solutionVersion) {
      return solutionVersion;
    }

    var manifestVersion = this.context && this.context.manifest ? String(this.context.manifest.version || '') : '';
    if (manifestVersion && manifestVersion !== '*') {
      return manifestVersion;
    }

    return 'Unknown';
  }

  private getFontFamilyOptions(): IDropdownOption[] {
    return [
      { key: '', text: 'Default' },
      { key: 'Segoe UI', text: 'Segoe UI' },
      { key: 'Calibri', text: 'Calibri' },
      { key: 'Arial', text: 'Arial' },
      { key: 'Tahoma', text: 'Tahoma' },
      { key: 'Verdana', text: 'Verdana' },
      { key: 'Georgia', text: 'Georgia' },
      { key: '"Times New Roman"', text: 'Times New Roman' },
      { key: '"Courier New"', text: 'Courier New' }
    ];
  }

  private getFontStyleOptions(): IDropdownOption[] {
    return [
      { key: 'normal', text: strings.PropFontStyleNormal },
      { key: 'italic', text: strings.PropFontStyleItalic },
      { key: 'oblique', text: strings.PropFontStyleOblique },
      { key: 'initial', text: strings.PropFontStyleInitial },
      { key: 'inherit', text: strings.PropFontStyleInherit },
      { key: 'unset', text: strings.PropFontStyleUnset }
    ];
  }

  private getTextAlignOptions(): IDropdownOption[] {
    return [
      { key: 'left', text: strings.PropTextAlignLeft },
      { key: 'center', text: strings.PropTextAlignCenter },
      { key: 'right', text: strings.PropTextAlignRight }
    ];
  }

  private validateJsonText(raw: string | undefined, expectedShape: 'array' | 'object'): { valid: boolean; message: string } {
    var source = String(raw || '').trim();
    if (!source) {
      return {
        valid: true,
        message: strings.JsonValidationEmpty
      };
    }

    try {
      var parsed = JSON.parse(source);
      if (expectedShape === 'array' && !Array.isArray(parsed)) {
        return {
          valid: false,
          message: strings.JsonValidationExpectedArray
        };
      }

      if (expectedShape === 'object' && (!parsed || typeof parsed !== 'object' || Array.isArray(parsed))) {
        return {
          valid: false,
          message: strings.JsonValidationExpectedObject
        };
      }

      return {
        valid: true,
        message: strings.JsonValidationSuccess
      };
    } catch (error) {
      return {
        valid: false,
        message: strings.JsonValidationFailedPrefix + ': ' + (error && error.message ? error.message : '')
      };
    }
  }

  private parseFilterJsonArray(raw: string | undefined): any[] {
    var source = String(raw || '').trim();
    if (!source) {
      return [];
    }

    var parsed = JSON.parse(source);
    if (!Array.isArray(parsed)) {
      throw new Error(strings.JsonValidationExpectedArray);
    }

    return parsed;
  }

  private handleValidateFilterJson(): void {
    var result = this.validateJsonText(this.properties.filterJson, 'array');
    this._filterJsonValidationMessage = result.message;
    this.context.propertyPane.refresh();
  }

  private handleAddFilterDesignerCondition(): void {
    try {
      var field = String(this.properties.filterDesignerField || '').trim();
      if (!field) {
        this._filterDesignerMessage = strings.PropFilterDesignerFieldRequired;
        this.context.propertyPane.refresh();
        return;
      }

      var operator = String(this.properties.filterDesignerOperator || 'contains').trim().toLowerCase();
      var logical = String(this.properties.filterDesignerLogical || 'and').trim().toLowerCase();
      var valueType = this.properties.filterDesignerValueType === 'expression' ? 'expression' : 'static';
      var value = String(this.properties.filterDesignerValue || '').trim();
      if (!value) {
        this._filterDesignerMessage = strings.PropFilterDesignerValueRequired;
        this.context.propertyPane.refresh();
        return;
      }
      if (valueType === 'expression' && !this.isSupportedFilterExpression(value)) {
        this._filterDesignerMessage = strings.PropFilterDesignerExpressionInvalid;
        this.context.propertyPane.refresh();
        return;
      }

      var existing = this.parseFilterJsonArray(this.properties.filterJson);
      existing.push({
        field: field,
        operator: operator,
        logical: logical,
        valueType: valueType,
        value: value
      });

      this.properties.filterJson = JSON.stringify(existing, null, 2);
      this._filterDesignerMessage = strings.PropFilterDesignerAddSuccess;
      this._filterJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && error.message ? error.message : strings.PropFilterDesignerActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleLoadSelectedFilterCondition(): void {
    try {
      var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
      var existing = this.parseFilterJsonArray(this.properties.filterJson);
      if (isNaN(index) || index < 0 || index >= existing.length) {
        this._filterDesignerMessage = strings.PropFilterDesignerSelectValid;
        this.context.propertyPane.refresh();
        return;
      }

      var selected = existing[index] || {};
      this.properties.filterDesignerField = String(selected.field || '');
      this.properties.filterDesignerOperator = (String(selected.operator || 'contains').toLowerCase() as any);
      this.properties.filterDesignerLogical = (String(selected.logical || 'and').toLowerCase() as any);
      this.properties.filterDesignerValueType = String(selected.valueType || '').toLowerCase() === 'expression' ? 'expression' : 'static';
      this.properties.filterDesignerValue = String(selected.value === undefined || selected.value === null ? '' : selected.value);
      this._filterDesignerMessage = strings.PropFilterDesignerLoadSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && error.message ? error.message : strings.PropFilterDesignerActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleUpdateSelectedFilterCondition(): void {
    try {
      var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
      var existing = this.parseFilterJsonArray(this.properties.filterJson);
      if (isNaN(index) || index < 0 || index >= existing.length) {
        this._filterDesignerMessage = strings.PropFilterDesignerSelectValid;
        this.context.propertyPane.refresh();
        return;
      }

      var field = String(this.properties.filterDesignerField || '').trim();
      if (!field) {
        this._filterDesignerMessage = strings.PropFilterDesignerFieldRequired;
        this.context.propertyPane.refresh();
        return;
      }

      var value = String(this.properties.filterDesignerValue || '').trim();
      if (!value) {
        this._filterDesignerMessage = strings.PropFilterDesignerValueRequired;
        this.context.propertyPane.refresh();
        return;
      }
      var valueType = this.properties.filterDesignerValueType === 'expression' ? 'expression' : 'static';
      if (valueType === 'expression' && !this.isSupportedFilterExpression(value)) {
        this._filterDesignerMessage = strings.PropFilterDesignerExpressionInvalid;
        this.context.propertyPane.refresh();
        return;
      }

      existing[index] = {
        field: field,
        operator: String(this.properties.filterDesignerOperator || 'contains').trim().toLowerCase(),
        logical: String(this.properties.filterDesignerLogical || 'and').trim().toLowerCase(),
        valueType: valueType,
        value: value
      };

      this.properties.filterJson = JSON.stringify(existing, null, 2);
      this._filterDesignerMessage = strings.PropFilterDesignerUpdateSuccess;
      this._filterJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && error.message ? error.message : strings.PropFilterDesignerActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleRemoveSelectedFilterCondition(): void {
    try {
      var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
      var existing = this.parseFilterJsonArray(this.properties.filterJson);
      if (isNaN(index) || index < 0 || index >= existing.length) {
        this._filterDesignerMessage = strings.PropFilterDesignerSelectValid;
        this.context.propertyPane.refresh();
        return;
      }

      existing.splice(index, 1);
      this.properties.filterJson = JSON.stringify(existing, null, 2);
      this.properties.filterDesignerSelectedIndex = '';
      this._filterDesignerMessage = strings.PropFilterDesignerRemoveSuccess;
      this._filterJsonValidationMessage = existing.length === 0 ? strings.JsonValidationEmpty : strings.JsonValidationSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && error.message ? error.message : strings.PropFilterDesignerActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleResetFilterJson(): void {
    this.properties.filterJson = '';
    this.properties.filterDesignerSelectedIndex = '';
    this._filterJsonValidationMessage = strings.JsonValidationEmpty;
    this._filterDesignerMessage = strings.PropFilterDesignerResetSuccess;
    this.context.propertyPane.refresh();
  }

  private isSupportedFilterExpression(value: string): boolean {
    var normalized = String(value || '').trim().toLowerCase();
    return normalized === 'today'
      || normalized === 'now'
      || normalized === 'me'
      || normalized === 'me.email'
      || normalized === 'me.login'
      || normalized === 'me.id'
      || /^date\(\s*[+-]?\d+\s*\)$/.test(normalized);
  }

  private parseConditionalStyleJsonArray(raw: string | undefined): any[] {
    var source = String(raw || '').trim();
    if (!source) {
      return [];
    }

    var parsed = JSON.parse(source);
    if (!Array.isArray(parsed)) {
      throw new Error(strings.JsonValidationExpectedArray);
    }

    return parsed;
  }

  private handleValidateConditionalStyleJson(): void {
    var result = this.validateJsonText(this.properties.conditionalStyleJson, 'array');
    this._conditionalStyleJsonValidationMessage = result.message;
    this.context.propertyPane.refresh();
  }

  private parseConditionalStylePriority(raw: string | undefined): number {
    var parsed = parseInt(String(raw || '').trim(), 10);
    if (isNaN(parsed)) {
      return 100;
    }

    return parsed;
  }

  private createConditionalStyleEntryFromDesigner(): any {
    var ruleName = String(this.properties.conditionalStyleRuleName || '').trim();
    if (!ruleName) {
      throw new Error(strings.PropConditionalStyleRuleNameRequired);
    }

    var scope = String(this.properties.conditionalStyleScope || 'row').trim().toLowerCase();
    if (scope !== 'row' && scope !== 'column') {
      scope = 'row';
    }

    var applyField = String(this.properties.conditionalStyleApplyField || '').trim();
    if (scope === 'column' && !applyField) {
      throw new Error(strings.PropConditionalStyleApplyFieldRequired);
    }

    var conditionField = String(this.properties.conditionalStyleConditionField || '').trim();
    if (!conditionField) {
      throw new Error(strings.PropConditionalStyleConditionFieldRequired);
    }

    var conditionValue = String(this.properties.conditionalStyleConditionValue || '').trim();
    if (!conditionValue) {
      throw new Error(strings.PropConditionalStyleConditionValueRequired);
    }
    var conditionValueType = this.properties.conditionalStyleConditionValueType === 'expression' ? 'expression' : 'static';
    if (conditionValueType === 'expression' && !this.isSupportedFilterExpression(conditionValue)) {
      throw new Error(strings.PropFilterDesignerExpressionInvalid);
    }

    return {
      rule: ruleName,
      scope: scope,
      applyField: applyField,
      enabled: this.properties.conditionalStyleEnabled !== false,
      priority: this.parseConditionalStylePriority(this.properties.conditionalStylePriority),
      conditionField: conditionField,
      operator: String(this.properties.conditionalStyleConditionOperator || 'eq').trim().toLowerCase(),
      logical: String(this.properties.conditionalStyleConditionLogical || 'and').trim().toLowerCase(),
      valueType: conditionValueType,
      value: conditionValue,
      style: {
        backgroundColor: String(this.properties.conditionalStyleBackgroundColor || '').trim(),
        color: String(this.properties.conditionalStyleForegroundColor || '').trim(),
        fontFamily: String(this.properties.conditionalStyleFontFamily || '').trim(),
        fontSize: String(this.properties.conditionalStyleFontSize || '').trim(),
        fontStyle: String(this.properties.conditionalStyleFontStyle || '').trim(),
        fontWeight: String(this.properties.conditionalStyleFontWeight || '').trim(),
        textAlign: String(this.properties.conditionalStyleTextAlign || '').trim()
      }
    };
  }

  private resetConditionalStyleDesigner(): void {
    this.properties.conditionalStyleRuleName = '';
    this.properties.conditionalStyleScope = 'row';
    this.properties.conditionalStyleApplyField = '';
    this.properties.conditionalStyleConditionField = '';
    this.properties.conditionalStyleConditionOperator = 'eq';
    this.properties.conditionalStyleConditionLogical = 'and';
    this.properties.conditionalStyleConditionValueType = 'static';
    this.properties.conditionalStyleEnabled = true;
    this.properties.conditionalStylePriority = '100';
    this.properties.conditionalStyleConditionValue = '';
    this.properties.conditionalStyleBackgroundColor = '';
    this.properties.conditionalStyleForegroundColor = '';
    this.properties.conditionalStyleFontFamily = '';
    this.properties.conditionalStyleFontSize = '';
    this.properties.conditionalStyleFontStyle = '';
    this.properties.conditionalStyleFontWeight = '';
    this.properties.conditionalStyleTextAlign = '';
    this.properties.conditionalStyleSelectedIndex = '';
    this._conditionalStyleDesignerRevision += 1;
  }

  private handleNewConditionalStyle(): void {
    this._isEditingConditionalStyle = false;
    this.resetConditionalStyleDesigner();
    this._conditionalStyleDesignerMessage = strings.PropConditionalStyleNewSuccess;
    this.context.propertyPane.refresh();
  }

  private handleAddConditionalStyleCondition(): void {
    try {
      var nextEntry = this.createConditionalStyleEntryFromDesigner();
      var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
      existing.push(nextEntry);
      this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
      this.resetConditionalStyleDesigner();
      this._conditionalStyleDesignerMessage = strings.PropConditionalStyleAddSuccess;
      this._conditionalStyleJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleLoadSelectedConditionalStyleCondition(): void {
    try {
      var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
      var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
      if (isNaN(index) || index < 0 || index >= existing.length) {
        this._conditionalStyleDesignerMessage = strings.PropConditionalStyleSelectValid;
        this.context.propertyPane.refresh();
        return;
      }

      var selected = existing[index] || {};
      var selectedStyle = selected.style || {};
      this.properties.conditionalStyleRuleName = String(selected.rule || '');
      this.properties.conditionalStyleScope = (String(selected.scope || 'row').toLowerCase() as any);
      this.properties.conditionalStyleApplyField = String(selected.applyField || '');
      this.properties.conditionalStyleConditionField = String(selected.conditionField || '');
      this.properties.conditionalStyleConditionOperator = (String(selected.operator || 'eq').toLowerCase() as any);
      this.properties.conditionalStyleConditionLogical = (String(selected.logical || 'and').toLowerCase() as any);
      this.properties.conditionalStyleConditionValueType = String(selected.valueType || '').toLowerCase() === 'expression' ? 'expression' : 'static';
      this.properties.conditionalStyleEnabled = selected.enabled !== false;
      this.properties.conditionalStylePriority = String(selected.priority === undefined || selected.priority === null ? 100 : selected.priority);
      this.properties.conditionalStyleConditionValue = String(selected.value === undefined || selected.value === null ? '' : selected.value);
      this.properties.conditionalStyleBackgroundColor = String(selectedStyle.backgroundColor || '');
      this.properties.conditionalStyleForegroundColor = String(selectedStyle.color || '');
      this.properties.conditionalStyleFontFamily = String(selectedStyle.fontFamily || '');
      this.properties.conditionalStyleFontSize = String(selectedStyle.fontSize || '');
      this.properties.conditionalStyleFontStyle = String(selectedStyle.fontStyle || '');
      this.properties.conditionalStyleFontWeight = String(selectedStyle.fontWeight || '');
      this.properties.conditionalStyleTextAlign = String(selectedStyle.textAlign || '');
      this._isEditingConditionalStyle = true;
      this._conditionalStyleDesignerRevision += 1;
      this._conditionalStyleDesignerMessage = strings.PropConditionalStyleLoadSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleUpdateSelectedConditionalStyleCondition(): void {
    try {
      var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
      var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
      if (isNaN(index) || index < 0 || index >= existing.length) {
        this._conditionalStyleDesignerMessage = strings.PropConditionalStyleSelectValid;
        this.context.propertyPane.refresh();
        return;
      }

      existing[index] = this.createConditionalStyleEntryFromDesigner();
      this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
      this._conditionalStyleDesignerMessage = strings.PropConditionalStyleUpdateSuccess;
      this._conditionalStyleJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleRemoveSelectedConditionalStyleCondition(): void {
    try {
      var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
      var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
      if (isNaN(index) || index < 0 || index >= existing.length) {
        this._conditionalStyleDesignerMessage = strings.PropConditionalStyleSelectValid;
        this.context.propertyPane.refresh();
        return;
      }

      existing.splice(index, 1);
      this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
      this._isEditingConditionalStyle = false;
      this.resetConditionalStyleDesigner();
      this._conditionalStyleDesignerMessage = strings.PropConditionalStyleRemoveSuccess;
      this._conditionalStyleJsonValidationMessage = existing.length === 0 ? strings.JsonValidationEmpty : strings.JsonValidationSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleMoveConditionalStyleConditionUp(): void {
    try {
      var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
      var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
      if (isNaN(index) || index <= 0 || index >= existing.length) {
        this._conditionalStyleDesignerMessage = strings.PropConditionalStyleMoveBoundary;
        this.context.propertyPane.refresh();
        return;
      }

      var current = existing[index];
      existing[index] = existing[index - 1];
      existing[index - 1] = current;
      this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
      this.properties.conditionalStyleSelectedIndex = String(index - 1);
      this._conditionalStyleDesignerMessage = strings.PropConditionalStyleMoveUpSuccess;
      this._conditionalStyleJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleMoveConditionalStyleConditionDown(): void {
    try {
      var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
      var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
      if (isNaN(index) || index < 0 || index >= existing.length - 1) {
        this._conditionalStyleDesignerMessage = strings.PropConditionalStyleMoveBoundary;
        this.context.propertyPane.refresh();
        return;
      }

      var current = existing[index];
      existing[index] = existing[index + 1];
      existing[index + 1] = current;
      this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
      this.properties.conditionalStyleSelectedIndex = String(index + 1);
      this._conditionalStyleDesignerMessage = strings.PropConditionalStyleMoveDownSuccess;
      this._conditionalStyleJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
    }

    this.context.propertyPane.refresh();
  }

  private handleResetConditionalStyleJson(): void {
    this.properties.conditionalStyleJson = '';
    this._isEditingConditionalStyle = false;
    this.resetConditionalStyleDesigner();
    this._conditionalStyleJsonValidationMessage = strings.JsonValidationEmpty;
    this._conditionalStyleDesignerMessage = strings.PropConditionalStyleResetSuccess;
    this.context.propertyPane.refresh();
  }

  private initializeDynamicDataSource(): void {
    this._dynamicDataSourceManager = (this.context as any).dynamicDataSourceManager || (this.context as any)._dynamicDataSourceManager;
    this.logDiagnostic('Initializing dynamic data source manager. Available=' + String(!!this._dynamicDataSourceManager));
    if (this._dynamicDataSourceManager && this._dynamicDataSourceManager.initializeSource) {
      this._dynamicDataSourceManager.initializeSource(this);
    } else if (this._dynamicDataSourceManager && this._dynamicDataSourceManager.registerSource) {
      this._dynamicDataSourceManager.registerSource(this);
    }
  }

  private notifyDynamicSourceChanged(): void {
    if (!this._dynamicDataSourceManager) {
      return;
    }

    if (this._dynamicDataSourceManager.notifySourceChanged) {
      this._dynamicDataSourceManager.notifySourceChanged();
    } else if (this._dynamicDataSourceManager.notifyDataChanged) {
      this._dynamicDataSourceManager.notifyDataChanged();
    }
  }

  private notifyDynamicData(propertyId: string): void {
    if (!this._dynamicDataSourceManager) {
      return;
    }

    var hasNotified = false;

    if (this._dynamicDataSourceManager.notifyPropertyChanged) {
      this._dynamicDataSourceManager.notifyPropertyChanged(propertyId);
      hasNotified = true;
    }

    if (this._dynamicDataSourceManager.notifySourceChanged) {
      this._dynamicDataSourceManager.notifySourceChanged();
      hasNotified = true;
    }

    if (!hasNotified && this._dynamicDataSourceManager.notifyDataChanged) {
      this._dynamicDataSourceManager.notifyDataChanged();
    }
  }

  private handleSelectionChange(itemId: number, mode: string): void {
    this._selectedItemId = itemId > 0 ? itemId : 0;
    this._selectedMode = mode || 'view';
    this.logDiagnostic('Selection changed. itemId=' + String(this._selectedItemId) + ', mode=' + String(this._selectedMode));
    this.notifyDynamicData('selectedItemId');
    this.notifyDynamicData('selectedMode');
    this.render();
  }

  private async getJsonWithAcceptFallback(url: string): Promise<any> {
    var response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      response = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=verbose'
          }
        }
      );
    }

    if (!response.ok) {
      response = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );
    }

    if (!response.ok) {
      throw new Error('Request failed. HTTP ' + String(response.status) + ' ' + response.statusText);
    }

    return response.json();
  }

  private async loadLists(): Promise<void> {
    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      this.logDiagnostic('Loading lists from web URL: ' + webUrl);
      var data = await this.getJsonWithAcceptFallback(
        webUrl + "/_api/web/lists?$select=Title,Hidden,BaseTemplate&$filter=Hidden eq false"
      );
      var lists = data && data.value ? data.value : [];

      if (!lists || lists.length === 0) {
        lists = data && data.d && data.d.results ? data.d.results : [];
      }

      this._lists = lists.map(function(list: any) {
        return {
          key: list.Title,
          text: list.Title
        };
      });
      this.logDiagnostic('Loaded lists successfully. Count=' + String(this._lists.length));
      this.context.propertyPane.refresh();
    } catch (error) {
      this.logDiagnostic('Failed to load lists: ' + (error && error.message ? error.message : String(error)));
      this._lists = [];
    }
  }

  private getTargetPageOptions(): IDropdownOption[] {
    var options = this._sitePages.length > 0
      ? this._sitePages.slice()
      : [{ key: '__defaultForm__', text: strings.PropLinkTargetDefaultFormOption }];
    var configuredUrl = String(this.properties.linkTargetPageUrl || '').trim();
    if (configuredUrl && configuredUrl !== '__defaultForm__'
      && !options.some(function(option: IDropdownOption) { return String(option.key) === configuredUrl; })) {
      options.push({ key: configuredUrl, text: configuredUrl + ' (' + strings.PropLinkTargetSavedUrlLabel + ')' });
    }
    return options;
  }

  private async loadSitePages(): Promise<void> {
    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      var libraryData = await this.getJsonWithAcceptFallback(
        webUrl + '/_api/web/lists?$select=Id,Title,BaseTemplate&$filter=(BaseTemplate eq 119 or BaseTemplate eq 850)'
      );
      var libraries = libraryData && libraryData.value ? libraryData.value
        : (libraryData && libraryData.d && libraryData.d.results ? libraryData.d.results : []);
      var pageOptionsByUrl: { [url: string]: IDropdownOption } = {};
      for (var libraryIndex = 0; libraryIndex < libraries.length; libraryIndex += 1) {
        var libraryId = String(libraries[libraryIndex].Id || '').replace(/[{}]/g, '');
        if (!libraryId) {
          continue;
        }
        try {
          var data = await this.getJsonWithAcceptFallback(
            webUrl + "/_api/web/lists(guid'" + libraryId
            + "')/items?$select=File/Name,File/ServerRelativeUrl&$expand=File&$top=5000"
          );
          var pages = data && data.value ? data.value
            : (data && data.d && data.d.results ? data.d.results : []);
          pages.filter(function(page: any) {
            return page.File && page.File.ServerRelativeUrl
              && /\.aspx(?:$|[?#])/i.test(String(page.File.ServerRelativeUrl));
          }).forEach(function(page: any) {
            var fileRef = String(page.File.ServerRelativeUrl);
            var title = String(page.File.Name || fileRef);
            pageOptionsByUrl[fileRef.toLowerCase()] = { key: fileRef, text: title + ' (' + fileRef + ')' };
          });
        } catch (pageError) {
          this.logDiagnostic('Failed to load pages from library "' + String(libraries[libraryIndex].Title || '') + '": '
            + String(pageError && pageError.message ? pageError.message : pageError));
        }
      }
      var defaultPageOptions: IDropdownOption[] = [
        { key: '__defaultForm__', text: strings.PropLinkTargetDefaultFormOption }
      ];
      this._sitePages = defaultPageOptions
        .concat(Object.keys(pageOptionsByUrl).map(function(url: string) { return pageOptionsByUrl[url]; })
          .sort(function(left: IDropdownOption, right: IDropdownOption) {
            return String(left.text).localeCompare(String(right.text));
          }));
      this.context.propertyPane.refresh();
    } catch (error) {
      this.logDiagnostic('Failed to load Site Pages: ' + (error && error.message ? error.message : String(error)));
      this._sitePages = [{ key: '__defaultForm__', text: strings.PropLinkTargetDefaultFormOption }];
      this.context.propertyPane.refresh();
    }
  }

  private async loadViews(listName: string): Promise<void> {
    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      this.logDiagnostic('Loading views for list: ' + String(listName));
      var data = await this.getJsonWithAcceptFallback(
        webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/views?$select=Id,Title,DefaultView"
      );
      var views = data && data.value ? data.value : [];

      if (!views || views.length === 0) {
        views = data && data.d && data.d.results ? data.d.results : [];
      }

      this._views = views.map(function(view: any) {
        return {
          key: String(view.Id),
          text: view.Title,
          isDefault: view.DefaultView === true
        };
      });

      if (!this.properties.viewId && this._views.length > 0) {
        for (var i = 0; i < this._views.length; i += 1) {
          if (this._views[i].isDefault) {
            this.properties.viewId = String(this._views[i].key);
            break;
          }
        }
        if (!this.properties.viewId) {
          this.properties.viewId = String(this._views[0].key);
        }
        this.logDiagnostic('Auto-selected default viewId=' + String(this.properties.viewId));
      }

      if (this.properties.viewId && (!this.properties.viewColumns || this.properties.viewColumns.length === 0)) {
        await this.loadViewColumns(listName, this.properties.viewId);
      }

      this.logDiagnostic('Loaded views successfully. Count=' + String(this._views.length));
      this.context.propertyPane.refresh();
      this.render();
    } catch (error) {
      this.logDiagnostic('Failed to load views for list ' + String(listName) + ': ' + (error && error.message ? error.message : String(error)));
      this._views = [];
      this.context.propertyPane.refresh();
    }
  }

  private async loadViewColumns(listName: string, viewId: string): Promise<void> {
    if (!listName || !viewId) {
      this.properties.viewColumns = [];
      this.context.propertyPane.refresh();
      this.render();
      return;
    }

    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      var listPath = "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')";
      var normalizedViewId = String(viewId).replace(/[{}]/g, '');
      var viewData = await this.getJsonWithAcceptFallback(
        webUrl + listPath + "/views/getById('" + encodeURIComponent(normalizedViewId) + "')/ViewFields"
      );
      var viewFields = normalizeStringCollection(viewData && viewData.value);
      if (viewFields.length === 0) {
        viewFields = normalizeStringCollection(viewData && viewData.Items);
      }
      if (viewFields.length === 0) {
        viewFields = normalizeStringCollection(viewData && viewData.d && viewData.d.Items);
      }

      var fieldData = await this.getJsonWithAcceptFallback(webUrl + listPath + "/fields?$select=InternalName,Title");
      var fields = fieldData && fieldData.value ? fieldData.value : [];
      if (!fields || fields.length === 0) {
        fields = fieldData && fieldData.d && fieldData.d.results ? fieldData.d.results : [];
      }
      var titleByName: { [fieldName: string]: string } = {};
      fields.forEach(function(field: any) {
        var internalName = String(field.InternalName || '');
        if (internalName) {
          titleByName[internalName] = String(field.Title || internalName);
        }
      });

      if (this.properties.listName !== listName || this.properties.viewId !== viewId) {
        return;
      }

      this.properties.viewColumns = viewFields.map(function(fieldName: string) {
        return {
          fieldName: fieldName,
          displayName: titleByName[fieldName] || fieldName,
          width: ''
        };
      });
      this.logDiagnostic('Initialized view column collection. Count=' + String(this.properties.viewColumns.length));
      this.context.propertyPane.refresh();
      this.render();
    } catch (error) {
      this.properties.viewColumns = [];
      this.context.propertyPane.refresh();
      this.render();
      this.logDiagnostic('Failed to initialize columns for view ' + viewId + ': ' + (error && error.message ? error.message : String(error)));
    }
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }

    console.log('[ListControlWebPart] ' + message);
  }
}