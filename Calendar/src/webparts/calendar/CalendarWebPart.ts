import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'CalendarWebPartStrings';
import Calendar from './components/Calendar';
import { ICalendarDataSource, ICalendarProps } from './components/ICalendarProps';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface IDropdownOption {
  key: string;
  text: string;
}

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

export interface ICalendarWebPartProps {
  title: string;
  description: string;
  listName: string;
  dataSources: ICalendarDataSource[];
  defaultView: string;
  showWeekends: boolean;
  calendarHeight: number;
  enableSwimlanes: boolean;
  swimlaneFieldName: string;
  swimlaneDays: number;
  swimlaneUnassignedLabel: string;
  filterJson: string;
  filterDesignerSourceList?: string;
  filterDesignerField?: string;
  filterDesignerOperator?: string;
  filterDesignerLogical?: string;
  filterDesignerValueType?: string;
  filterDesignerValue?: string;
  filterDesignerSelectedIndex?: string;
  filterDesignerLookupPick?: string;
  conditionalStyleJson: string;
  conditionalStyleSourceList?: string;
  conditionalStyleRuleName?: string;
  conditionalStyleConditionField?: string;
  conditionalStyleConditionOperator?: string;
  conditionalStyleConditionLogical?: string;
  conditionalStyleConditionValueType?: string;
  conditionalStyleConditionValue?: string;
  conditionalStyleEnabled?: boolean;
  conditionalStylePriority?: string;
  conditionalStyleBackgroundColor?: string;
  conditionalStyleBorderColor?: string;
  conditionalStyleTextColor?: string;
  conditionalStyleFontFamily?: string;
  conditionalStyleFontSize?: string;
  conditionalStyleFontStyle?: string;
  conditionalStyleFontWeight?: string;
  conditionalStyleSelectedIndex?: string;
  showLinkToItem: boolean;
  enableEventDetails: boolean;
  eventDetailsFields: any[];
  eventDetailsFieldsJson: string;
  detailsLinkPresentation: string;
  linkTargetPageUrl: string;
  linkTargetIdParam: string;
  includeReturnUrlParam: boolean;
  calendarBackgroundColor: string;
  webPartBorderColor: string;
  webPartBorderWidth: number;
  webPartCornerRadius: number;
  webPartTitleTextColor: string;
  webPartTitleFontFamily: string;
  webPartTitleFontSize: string;
  webPartTitleFontStyle: string;
  webPartTitleFontBold: boolean;
  webPartTitleAlignment: string;
  webPartTitleBackgroundColor: string;
  webPartTitleMinHeight: number;
  webPartDescriptionTextColor: string;
  webPartDescriptionFontFamily: string;
  webPartDescriptionFontSize: string;
  webPartDescriptionFontStyle: string;
  webPartDescriptionFontBold: boolean;
  calendarTitleTextColor: string;
  calendarTitleFontFamily: string;
  calendarTitleFontSize: string;
  calendarTitleFontStyle: string;
  calendarTitleFontBold: boolean;
  toolbarButtonBackgroundColor: string;
  toolbarButtonTextColor: string;
  toolbarButtonBorderColor: string;
  toolbarButtonActiveBackgroundColor: string;
  toolbarButtonActiveTextColor: string;
  toolbarButtonFontFamily: string;
  toolbarButtonFontSize: string;
  toolbarButtonFontStyle: string;
  toolbarButtonFontBold: boolean;
  toolbarButtonBorderWidth: number;
  toolbarButtonCornerRadius: number;
  eventBackgroundColor: string;
  eventTextColor: string;
  eventBorderColor: string;
  eventFontFamily: string;
  eventFontSize: string;
  eventFontStyle: string;
  eventFontBold: boolean;
  eventBorderWidth: number;
  eventCornerRadius: number;
  selectedEventBackgroundColor: string;
  selectedEventTextColor: string;
  selectedEventBorderColor: string;
  selectedEventFontFamily: string;
  selectedEventFontSize: string;
  selectedEventFontStyle: string;
  selectedEventFontBold: boolean;
  selectedEventBorderWidth: number;
  selectedEventCornerRadius: number;
  detailsPanelBackgroundColor: string;
  detailsPanelBorderColor: string;
  detailsPanelBorderWidth: number;
  detailsPanelCornerRadius: number;
  detailsPanelMaxWidth: number;
  detailsPanelPadding: number;
  detailsPanelShadow: boolean;
  detailsOverlayColor: string;
  detailsOverlayOpacity: number;
  detailsTitleTextColor: string;
  detailsTitleFontFamily: string;
  detailsTitleFontSize: string;
  detailsTitleFontStyle: string;
  detailsTitleFontBold: boolean;
  detailsLabelTextColor: string;
  detailsLabelFontFamily: string;
  detailsLabelFontSize: string;
  detailsLabelFontStyle: string;
  detailsLabelFontBold: boolean;
  detailsValueTextColor: string;
  detailsValueFontFamily: string;
  detailsValueFontSize: string;
  detailsValueFontStyle: string;
  detailsValueFontBold: boolean;
  detailsDividerColor: string;
  detailsDividerWidth: number;
  detailsCloseButtonColor: string;
  detailsCloseButtonBackgroundColor: string;
  detailsCloseButtonHoverBackgroundColor: string;
  detailsCloseButtonCornerRadius: number;
  enableDiagnostics: boolean;
}

export default class CalendarWebPart extends BaseClientSideWebPart<ICalendarWebPartProps> {
  private _lists: IDropdownOption[] = [];
  private _sitePages: IDropdownOption[] = [];
  private _dateFieldOptions: IDropdownOption[] = [];
  private _listBaseTemplateByTitle: { [lowerTitle: string]: number } = {};
  private _calendarDateFieldsByListTitle: {
    [lowerTitle: string]: { startFieldName: string; endFieldName: string }
  } = {};
  private _listFields: IDropdownOption[] = [];
  private _eventDetailFieldsByListTitle: { [lowerTitle: string]: IDropdownOption[] } = {};
  private _fieldTypeByInternalName: { [internalName: string]: string } = {};
  private _fieldMetadataByInternalName: { [internalName: string]: { type: string; richText: boolean } } = {};
  private _fieldLookupListByInternalName: { [internalName: string]: string } = {};
  private _filterLookupItemOptions: IDropdownOption[] = [];
  private _filterLookupMessage: string = '';
  private _filterJsonValidationMessage: string = '';
  private _filterDesignerMessage: string = '';
  private _selectedItemId: number = 0;
  private _selectedMode: string = 'view';
  private _dynamicDataSourceManager: any;
  private _conditionalStyleJsonValidationMessage: string = '';
  private _conditionalStyleDesignerMessage: string = '';
  private _isEditingConditionalStyle: boolean = false;
  private _conditionalStyleDesignerRevision: number = 0;

  public get id(): string {
    return this.context.instanceId;
  }

  public get metadata(): IDynamicDataSourceMetadataCompat {
    return {
      title: this.properties.title || 'Calendar',
      description: 'The event selected in this Calendar web part.',
      alias: this.context.manifest.alias,
      componentId: this.context.manifest.id,
      instanceId: this.context.instanceId
    };
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinitionCompat> {
    return [
      {
        id: 'selectedItemId',
        title: 'Selected Item ID',
        description: 'The SharePoint item ID of the selected calendar event.'
      },
      {
        id: 'selectedMode',
        title: 'Selected Mode',
        description: 'The desired form mode for the selected event.'
      }
    ];
  }

  public getPropertyValue(propertyId: string): any {
    if (propertyId === 'selectedItemId') {
      return this._selectedItemId;
    }

    if (propertyId === 'selectedMode') {
      return this._selectedMode;
    }

    throw new Error('Bad property id');
  }

  private getPrimaryListName(): string {
    var sources = Array.isArray(this.properties.dataSources) ? this.properties.dataSources : [];
    for (var index = 0; index < sources.length; index += 1) {
      var sourceListName = String(sources[index] && sources[index].listName || '').trim();
      if (sourceListName) {
        return sourceListName;
      }
    }
    return String(this.properties.listName || '').trim();
  }

  public render(): void {
    var primaryListName = this.getPrimaryListName();
    this.logDiagnostic('render() called. primaryListName=' + String(primaryListName || '(none)'));
    const element: React.ReactElement<ICalendarProps> = React.createElement(
      Calendar,
      {
        title: this.properties.title,
        description: this.properties.description || '',
        displayMode: this.displayMode,
        spfxContext: this.context,
        listName: primaryListName,
        dataSources: (Array.isArray(this.properties.dataSources) ? this.properties.dataSources : []).map((source: ICalendarDataSource) => {
          var listTitleKey = String(source.listName || '').toLowerCase();
          var inferredDateFields = this._calendarDateFieldsByListTitle[listTitleKey];
          var configuredStartField = String(source.startFieldName || '');
          var configuredEndField = String(source.endFieldName || '');
          return {
            listName: source.listName,
            targetPageUrl: source.targetPageUrl,
            startFieldName: inferredDateFields
              && (!configuredStartField || configuredStartField === '__eventsDefault__')
              ? inferredDateFields.startFieldName : source.startFieldName,
            endFieldName: inferredDateFields
              && (!configuredEndField || configuredEndField === '__eventsDefault__')
              ? inferredDateFields.endFieldName : source.endFieldName,
            backgroundColor: source.backgroundColor,
            baseTemplate: this._listBaseTemplateByTitle[listTitleKey]
          };
        }),
        defaultView: this.properties.defaultView || 'dayGridMonth',
        showWeekends: this.properties.showWeekends !== false,
        calendarHeight: typeof this.properties.calendarHeight === 'number' ? this.properties.calendarHeight : 600,
        enableSwimlanes: this.properties.enableSwimlanes === true,
        swimlaneFieldName: this.properties.swimlaneFieldName || '',
        swimlaneDays: typeof this.properties.swimlaneDays === 'number' ? this.properties.swimlaneDays : 7,
        swimlaneUnassignedLabel: this.properties.swimlaneUnassignedLabel || strings.SwimlaneUnassignedDefault,
        filterJson: this.properties.filterJson || '',
        conditionalStyleJson: this.properties.conditionalStyleJson || '',
        showLinkToItem: this.properties.showLinkToItem === true,
        enableEventDetails: this.properties.enableEventDetails !== false,
        eventDetailsFields: Array.isArray(this.properties.eventDetailsFields) ? this.properties.eventDetailsFields : [],
        eventDetailsFieldsJson: this.properties.eventDetailsFieldsJson || '',
        fieldMetadata: this._fieldMetadataByInternalName,
        detailsLinkPresentation: this.properties.detailsLinkPresentation || 'button',
        linkTargetPageUrl: this.properties.linkTargetPageUrl || '',
        linkTargetIdParam: this.properties.linkTargetIdParam || 'itemid',
        includeReturnUrlParam: this.properties.includeReturnUrlParam === true,
        calendarBackgroundColor: this.properties.calendarBackgroundColor || '#ffffff',
        webPartBorderColor: this.properties.webPartBorderColor || '#d2d0ce',
        webPartBorderWidth: typeof this.properties.webPartBorderWidth === 'number' ? this.properties.webPartBorderWidth : 0,
        webPartCornerRadius: typeof this.properties.webPartCornerRadius === 'number' ? this.properties.webPartCornerRadius : 0,
        webPartTitleTextColor: this.properties.webPartTitleTextColor || '#323130',
        webPartTitleFontFamily: this.properties.webPartTitleFontFamily || 'inherit',
        webPartTitleFontSize: this.properties.webPartTitleFontSize || '20px',
        webPartTitleFontStyle: this.properties.webPartTitleFontStyle || 'normal',
        webPartTitleFontBold: this.properties.webPartTitleFontBold !== false,
        webPartTitleAlignment: this.properties.webPartTitleAlignment || 'left',
        webPartTitleBackgroundColor: this.properties.webPartTitleBackgroundColor || '#ffffff',
        webPartTitleMinHeight: typeof this.properties.webPartTitleMinHeight === 'number'
          ? this.properties.webPartTitleMinHeight : 0,
        webPartDescriptionTextColor: this.properties.webPartDescriptionTextColor || '#605e5c',
        webPartDescriptionFontFamily: this.properties.webPartDescriptionFontFamily || 'inherit',
        webPartDescriptionFontSize: this.properties.webPartDescriptionFontSize || '14px',
        webPartDescriptionFontStyle: this.properties.webPartDescriptionFontStyle || 'normal',
        webPartDescriptionFontBold: this.properties.webPartDescriptionFontBold === true,
        calendarTitleTextColor: this.properties.calendarTitleTextColor || '#323130',
        calendarTitleFontFamily: this.properties.calendarTitleFontFamily || 'inherit',
        calendarTitleFontSize: this.properties.calendarTitleFontSize || '28px',
        calendarTitleFontStyle: this.properties.calendarTitleFontStyle || 'normal',
        calendarTitleFontBold: this.properties.calendarTitleFontBold !== false,
        toolbarButtonBackgroundColor: this.properties.toolbarButtonBackgroundColor || '#2f4358',
        toolbarButtonTextColor: this.properties.toolbarButtonTextColor || '#ffffff',
        toolbarButtonBorderColor: this.properties.toolbarButtonBorderColor || '#2f4358',
        toolbarButtonActiveBackgroundColor: this.properties.toolbarButtonActiveBackgroundColor || '#172536',
        toolbarButtonActiveTextColor: this.properties.toolbarButtonActiveTextColor || '#ffffff',
        toolbarButtonFontFamily: this.properties.toolbarButtonFontFamily || 'inherit',
        toolbarButtonFontSize: this.properties.toolbarButtonFontSize || '16px',
        toolbarButtonFontStyle: this.properties.toolbarButtonFontStyle || 'normal',
        toolbarButtonFontBold: this.properties.toolbarButtonFontBold === true,
        toolbarButtonBorderWidth: typeof this.properties.toolbarButtonBorderWidth === 'number' ? this.properties.toolbarButtonBorderWidth : 1,
        toolbarButtonCornerRadius: typeof this.properties.toolbarButtonCornerRadius === 'number' ? this.properties.toolbarButtonCornerRadius : 4,
        eventBackgroundColor: this.properties.eventBackgroundColor || '#3788d8',
        eventTextColor: this.properties.eventTextColor || '#ffffff',
        eventBorderColor: this.properties.eventBorderColor || '#2c6faa',
        eventFontFamily: this.properties.eventFontFamily || 'inherit',
        eventFontSize: this.properties.eventFontSize || 'inherit',
        eventFontStyle: this.properties.eventFontStyle || 'normal',
        eventFontBold: this.properties.eventFontBold === true,
        eventBorderWidth: typeof this.properties.eventBorderWidth === 'number' ? this.properties.eventBorderWidth : 1,
        eventCornerRadius: typeof this.properties.eventCornerRadius === 'number' ? this.properties.eventCornerRadius : 3,
        selectedEventBackgroundColor: this.properties.selectedEventBackgroundColor || '#ffb900',
        selectedEventTextColor: this.properties.selectedEventTextColor || '#201f1e',
        selectedEventBorderColor: this.properties.selectedEventBorderColor || '#8a4b00',
        selectedEventFontFamily: this.properties.selectedEventFontFamily || 'inherit',
        selectedEventFontSize: this.properties.selectedEventFontSize || 'inherit',
        selectedEventFontStyle: this.properties.selectedEventFontStyle || 'normal',
        selectedEventFontBold: this.properties.selectedEventFontBold !== false,
        selectedEventBorderWidth: typeof this.properties.selectedEventBorderWidth === 'number'
          ? this.properties.selectedEventBorderWidth : 3,
        selectedEventCornerRadius: typeof this.properties.selectedEventCornerRadius === 'number'
          ? this.properties.selectedEventCornerRadius : 5,
        detailsPanelBackgroundColor: this.properties.detailsPanelBackgroundColor || '#ffffff',
        detailsPanelBorderColor: this.properties.detailsPanelBorderColor || '#8a8886',
        detailsPanelBorderWidth: typeof this.properties.detailsPanelBorderWidth === 'number' ? this.properties.detailsPanelBorderWidth : 1,
        detailsPanelCornerRadius: typeof this.properties.detailsPanelCornerRadius === 'number' ? this.properties.detailsPanelCornerRadius : 6,
        detailsPanelMaxWidth: typeof this.properties.detailsPanelMaxWidth === 'number' ? this.properties.detailsPanelMaxWidth : 680,
        detailsPanelPadding: typeof this.properties.detailsPanelPadding === 'number' ? this.properties.detailsPanelPadding : 30,
        detailsPanelShadow: this.properties.detailsPanelShadow !== false,
        detailsOverlayColor: this.properties.detailsOverlayColor || '#000000',
        detailsOverlayOpacity: typeof this.properties.detailsOverlayOpacity === 'number' ? this.properties.detailsOverlayOpacity : 28,
        detailsTitleTextColor: this.properties.detailsTitleTextColor || '#201f1e',
        detailsTitleFontFamily: this.properties.detailsTitleFontFamily || 'inherit',
        detailsTitleFontSize: this.properties.detailsTitleFontSize || '24px',
        detailsTitleFontStyle: this.properties.detailsTitleFontStyle || 'normal',
        detailsTitleFontBold: this.properties.detailsTitleFontBold !== false,
        detailsLabelTextColor: this.properties.detailsLabelTextColor || '#323130',
        detailsLabelFontFamily: this.properties.detailsLabelFontFamily || 'inherit',
        detailsLabelFontSize: this.properties.detailsLabelFontSize || '14px',
        detailsLabelFontStyle: this.properties.detailsLabelFontStyle || 'normal',
        detailsLabelFontBold: this.properties.detailsLabelFontBold !== false,
        detailsValueTextColor: this.properties.detailsValueTextColor || '#201f1e',
        detailsValueFontFamily: this.properties.detailsValueFontFamily || 'inherit',
        detailsValueFontSize: this.properties.detailsValueFontSize || '14px',
        detailsValueFontStyle: this.properties.detailsValueFontStyle || 'normal',
        detailsValueFontBold: this.properties.detailsValueFontBold === true,
        detailsDividerColor: this.properties.detailsDividerColor || '#edebe9',
        detailsDividerWidth: typeof this.properties.detailsDividerWidth === 'number' ? this.properties.detailsDividerWidth : 1,
        detailsCloseButtonColor: this.properties.detailsCloseButtonColor || '#323130',
        detailsCloseButtonBackgroundColor: this.properties.detailsCloseButtonBackgroundColor || '#ffffff',
        detailsCloseButtonHoverBackgroundColor: this.properties.detailsCloseButtonHoverBackgroundColor || '#edebe9',
        detailsCloseButtonCornerRadius: typeof this.properties.detailsCloseButtonCornerRadius === 'number'
          ? this.properties.detailsCloseButtonCornerRadius : 2,
        enableDiagnostics: this.properties.enableDiagnostics !== false,
        onEventSelectionChange: (itemId: number, mode: string) => this.handleEventSelectionChange(itemId, mode),
        fUpdateProperty: (value: string) => {
          this.properties.title = value;
        },
        fPropertyPaneOpen: this.context.propertyPane.open
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.logDiagnostic('onInit started. listName=' + String(this.properties.listName || '(none)'));
    this.initializeDynamicDataSource();
    if (!Array.isArray(this.properties.eventDetailsFields) && String(this.properties.eventDetailsFieldsJson || '').trim()) {
      try {
        var legacyDetailFields = JSON.parse(this.properties.eventDetailsFieldsJson);
        this.properties.eventDetailsFields = Array.isArray(legacyDetailFields) ? legacyDetailFields : [];
      } catch (_error) {
        this.properties.eventDetailsFields = [];
      }
    }
    return Promise.all([this.loadLists(), this.loadSitePages()]).then(() => {
      var primaryListName = this.getPrimaryListName();
      return Promise.all([
        primaryListName ? this.loadListFields(primaryListName) : this.loadEventDetailFields(''),
        this.loadAllEventDetailFields()
      ]).then(() => undefined);
    });
  }

  protected onDispose(): void {
    this.logDiagnostic('onDispose called.');
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected onPropertyPaneConfigurationStart(): void {
    var primaryListName = this.getPrimaryListName();
    this.logDiagnostic('Property pane opened. primaryListName=' + String(primaryListName || '(none)'));
    if (this._lists.length === 0) {
      this.loadLists();
    }
    if (this._sitePages.length === 0) {
      this.loadSitePages();
    }
    if (primaryListName && this._listFields.length === 0) {
      this.loadListFields(primaryListName);
    }
    this.loadAllEventDetailFields();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
    if (propertyPath === 'linkTargetPageUrl' && oldValue !== newValue) {
      this.properties.includeReturnUrlParam = !!newValue && newValue !== '__defaultForm__';
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
      this.render();
      return;
    }
    if (propertyPath === 'listName' && oldValue !== newValue) {
      this.properties.swimlaneFieldName = '';
      this._listFields = [];
      this._fieldTypeByInternalName = {};
      this._fieldMetadataByInternalName = {};
      this._fieldLookupListByInternalName = {};
      var designerListName = String(newValue || this.getPrimaryListName());
      if (designerListName) {
        this.loadListFields(designerListName);
      }
    }
    if (propertyPath === 'dataSources' && oldValue !== newValue) {
      this.properties.swimlaneFieldName = '';
      this._listFields = [];
      this._eventDetailFieldsByListTitle = {};
      this._fieldTypeByInternalName = {};
      this._fieldMetadataByInternalName = {};
      this._fieldLookupListByInternalName = {};
      var primaryListName = this.getPrimaryListName();
      if (primaryListName) {
        this.loadListFields(primaryListName);
      }
      this.loadAllEventDetailFields();
    }
    if (propertyPath === 'filterDesignerField' && oldValue !== newValue) {
      this._filterLookupItemOptions = [];
      this.properties.filterDesignerLookupPick = '';
      this._filterLookupMessage = '';
      if (this.isLookupTypeField(String(newValue || ''))) {
        this.handleLoadFilterLookupItems();
      }
    }
    if ((propertyPath === 'filterDesignerSourceList' || propertyPath === 'conditionalStyleSourceList')
      && oldValue !== newValue) {
      if (propertyPath === 'filterDesignerSourceList') {
        this.properties.filterDesignerField = '';
        this.properties.filterDesignerLookupPick = '';
        this._filterLookupItemOptions = [];
        this._filterLookupMessage = '';
      } else {
        this.properties.conditionalStyleConditionField = '';
      }
      this._listFields = [];
      this._fieldTypeByInternalName = {};
      this._fieldMetadataByInternalName = {};
      this._fieldLookupListByInternalName = {};
      if (newValue) {
        this.loadListFields(String(newValue));
      }
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private validateJson(value: string): string {
    if (!String(value || '').trim()) {
      return '';
    }
    try {
      var parsed = JSON.parse(value);
      return Array.isArray(parsed) ? '' : 'Value must be a JSON array.';
    } catch (error) {
      return 'Invalid JSON: ' + (error && (error as any).message ? (error as any).message : String(error));
    }
  }

  private parseConditionalStyleJsonArray(): any[] {
    var source = String(this.properties.conditionalStyleJson || '').trim();
    if (!source) {
      return [];
    }
    var parsed = JSON.parse(source);
    if (!Array.isArray(parsed)) {
      throw new Error(strings.JsonValidationExpectedArray);
    }
    return parsed;
  }

  private parseFilterJsonArray(): any[] {
    var source = String(this.properties.filterJson || '').trim();
    if (!source) {
      return [];
    }
    var parsed = JSON.parse(source);
    if (!Array.isArray(parsed)) {
      throw new Error(strings.JsonValidationExpectedArray);
    }
    return parsed;
  }

  private persistFilters(entries: any[]): void {
    this.properties.filterJson = entries.length > 0 ? JSON.stringify(entries, null, 2) : '';
    this._filterJsonValidationMessage = entries.length > 0 ? strings.JsonValidationSuccess : strings.JsonValidationEmpty;
    this.render();
  }

  private createFilterEntry(): any {
    var field = String(this.properties.filterDesignerField || '').trim();
    var value = String(this.properties.filterDesignerValue || '').trim();
    if (!field) {
      throw new Error(strings.FilterDesignerFieldRequired);
    }
    if (!value) {
      throw new Error(strings.FilterDesignerValueRequired);
    }
    var valueType = this.properties.filterDesignerValueType === 'expression' ? 'expression' : 'static';
    if (valueType === 'expression' && !this.isSupportedConditionExpression(value)) {
      throw new Error(strings.ConditionalStyleExpressionInvalid);
    }
    return {
      sourceListName: String(this.properties.filterDesignerSourceList || '').trim(),
      field: field,
      operator: String(this.properties.filterDesignerOperator || 'contains'),
      logical: String(this.properties.filterDesignerLogical || 'and'),
      valueType: valueType,
      value: value
    };
  }

  private handleAddFilter(): void {
    try {
      var entries = this.parseFilterJsonArray();
      entries.push(this.createFilterEntry());
      this.persistFilters(entries);
      this._filterDesignerMessage = strings.FilterDesignerAddSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && (error as any).message ? (error as any).message : strings.FilterDesignerActionFailed;
    }
    this.context.propertyPane.refresh();
  }

  private handleLoadFilter(): void {
    try {
      var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
      var entries = this.parseFilterJsonArray();
      if (isNaN(index) || index < 0 || index >= entries.length) {
        throw new Error(strings.FilterDesignerSelectValid);
      }
      var selected = entries[index] || {};
      this.properties.filterDesignerSourceList = String(selected.sourceListName || '');
      this.properties.filterDesignerField = String(selected.field || '');
      this.properties.filterDesignerOperator = String(selected.operator || 'contains');
      this.properties.filterDesignerLogical = String(selected.logical || 'and');
      this.properties.filterDesignerValueType = selected.valueType === 'expression' ? 'expression' : 'static';
      this.properties.filterDesignerValue = String(selected.value === undefined ? '' : selected.value);
      var filterSourceList = String(this.properties.filterDesignerSourceList || this.getPrimaryListName());
      if (filterSourceList) {
        this.loadListFields(filterSourceList);
      }
      this._filterDesignerMessage = strings.FilterDesignerLoadSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && (error as any).message ? (error as any).message : strings.FilterDesignerActionFailed;
    }
    this.context.propertyPane.refresh();
  }

  private handleUpdateFilter(): void {
    try {
      var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
      var entries = this.parseFilterJsonArray();
      if (isNaN(index) || index < 0 || index >= entries.length) {
        throw new Error(strings.FilterDesignerSelectValid);
      }
      entries[index] = this.createFilterEntry();
      this.persistFilters(entries);
      this._filterDesignerMessage = strings.FilterDesignerUpdateSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && (error as any).message ? (error as any).message : strings.FilterDesignerActionFailed;
    }
    this.context.propertyPane.refresh();
  }

  private handleRemoveFilter(): void {
    try {
      var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
      var entries = this.parseFilterJsonArray();
      if (isNaN(index) || index < 0 || index >= entries.length) {
        throw new Error(strings.FilterDesignerSelectValid);
      }
      entries.splice(index, 1);
      this.properties.filterDesignerSelectedIndex = '';
      this.persistFilters(entries);
      this._filterDesignerMessage = strings.FilterDesignerRemoveSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && (error as any).message ? (error as any).message : strings.FilterDesignerActionFailed;
    }
    this.context.propertyPane.refresh();
  }

  private handleResetFilters(): void {
    this.persistFilters([]);
    this.properties.filterDesignerSelectedIndex = '';
    this._filterDesignerMessage = strings.FilterDesignerResetSuccess;
    this.context.propertyPane.refresh();
  }

  private handleValidateFilters(): void {
    this._filterJsonValidationMessage = this.validateJson(this.properties.filterJson || '') || strings.JsonValidationSuccess;
    this.context.propertyPane.refresh();
  }

  private isSupportedConditionExpression(value: string): boolean {
    var normalized = String(value || '').trim().toLowerCase();
    return normalized === 'today'
      || normalized === 'now'
      || normalized === 'me'
      || normalized === 'me.email'
      || normalized === 'me.login'
      || normalized === 'me.id'
      || normalized === 'me.title'
      || /^date\(\s*[+-]?\d+\s*\)$/.test(normalized);
  }

  private createConditionalStyleEntry(): any {
    var ruleName = String(this.properties.conditionalStyleRuleName || '').trim();
    var fieldName = String(this.properties.conditionalStyleConditionField || '').trim();
    var conditionValue = String(this.properties.conditionalStyleConditionValue || '').trim();
    if (!ruleName) {
      throw new Error(strings.ConditionalStyleRuleNameRequired);
    }
    if (!fieldName) {
      throw new Error(strings.ConditionalStyleFieldRequired);
    }
    if (!conditionValue) {
      throw new Error(strings.ConditionalStyleValueRequired);
    }

    var valueType = this.properties.conditionalStyleConditionValueType === 'expression' ? 'expression' : 'static';
    if (valueType === 'expression' && !this.isSupportedConditionExpression(conditionValue)) {
      throw new Error(strings.ConditionalStyleExpressionInvalid);
    }

    var priority = parseInt(String(this.properties.conditionalStylePriority || '100'), 10);
    return {
      rule: ruleName,
      sourceListName: String(this.properties.conditionalStyleSourceList || '').trim(),
      enabled: this.properties.conditionalStyleEnabled !== false,
      priority: isNaN(priority) ? 100 : priority,
      conditionField: fieldName,
      operator: String(this.properties.conditionalStyleConditionOperator || 'eq'),
      logical: String(this.properties.conditionalStyleConditionLogical || 'and'),
      valueType: valueType,
      value: conditionValue,
      style: {
        backgroundColor: String(this.properties.conditionalStyleBackgroundColor || '').trim(),
        borderColor: String(this.properties.conditionalStyleBorderColor || '').trim(),
        textColor: String(this.properties.conditionalStyleTextColor || '').trim(),
        fontFamily: String(this.properties.conditionalStyleFontFamily || '').trim(),
        fontSize: String(this.properties.conditionalStyleFontSize || '').trim(),
        fontStyle: String(this.properties.conditionalStyleFontStyle || '').trim(),
        fontWeight: String(this.properties.conditionalStyleFontWeight || '').trim()
      }
    };
  }

  private resetConditionalStyleDesigner(): void {
    this.properties.conditionalStyleSourceList = '';
    this.properties.conditionalStyleRuleName = '';
    this.properties.conditionalStyleConditionField = '';
    this.properties.conditionalStyleConditionOperator = 'eq';
    this.properties.conditionalStyleConditionLogical = 'and';
    this.properties.conditionalStyleConditionValueType = 'static';
    this.properties.conditionalStyleConditionValue = '';
    this.properties.conditionalStyleEnabled = true;
    this.properties.conditionalStylePriority = '100';
    this.properties.conditionalStyleBackgroundColor = '';
    this.properties.conditionalStyleBorderColor = '';
    this.properties.conditionalStyleTextColor = '';
    this.properties.conditionalStyleFontFamily = '';
    this.properties.conditionalStyleFontSize = '';
    this.properties.conditionalStyleFontStyle = '';
    this.properties.conditionalStyleFontWeight = '';
    this.properties.conditionalStyleSelectedIndex = '';
    this._conditionalStyleDesignerRevision += 1;
    var primaryListName = this.getPrimaryListName();
    if (primaryListName) {
      this.loadListFields(primaryListName);
    }
  }

  private persistConditionalStyles(entries: any[]): void {
    this.properties.conditionalStyleJson = entries.length > 0 ? JSON.stringify(entries, null, 2) : '';
    this._conditionalStyleJsonValidationMessage = entries.length > 0 ? strings.JsonValidationSuccess : strings.JsonValidationEmpty;
    this.render();
  }

  private handleNewConditionalStyle(): void {
    this._isEditingConditionalStyle = false;
    this.resetConditionalStyleDesigner();
    this._conditionalStyleDesignerMessage = strings.ConditionalStyleNewSuccess;
    this.context.propertyPane.refresh();
  }

  private handleAddConditionalStyle(): void {
    try {
      var entries = this.parseConditionalStyleJsonArray();
      entries.push(this.createConditionalStyleEntry());
      this.persistConditionalStyles(entries);
      this.resetConditionalStyleDesigner();
      this._conditionalStyleDesignerMessage = strings.ConditionalStyleAddSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && (error as any).message ? (error as any).message : strings.ConditionalStyleActionFailed;
    }
    this.context.propertyPane.refresh();
  }

  private handleLoadConditionalStyle(): void {
    try {
      var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
      var entries = this.parseConditionalStyleJsonArray();
      if (isNaN(index) || index < 0 || index >= entries.length) {
        throw new Error(strings.ConditionalStyleSelectValid);
      }
      var selected = entries[index] || {};
      var style = selected.style || {};
      this.properties.conditionalStyleSourceList = String(selected.sourceListName || '');
      this.properties.conditionalStyleRuleName = String(selected.rule || selected.ruleName || '');
      this.properties.conditionalStyleConditionField = String(selected.conditionField || selected.field || '');
      this.properties.conditionalStyleConditionOperator = String(selected.operator || 'eq');
      this.properties.conditionalStyleConditionLogical = String(selected.logical || 'and');
      this.properties.conditionalStyleConditionValueType = selected.valueType === 'expression' ? 'expression' : 'static';
      this.properties.conditionalStyleConditionValue = String(selected.value === undefined ? '' : selected.value);
      this.properties.conditionalStyleEnabled = selected.enabled !== false;
      this.properties.conditionalStylePriority = String(selected.priority === undefined ? 100 : selected.priority);
      this.properties.conditionalStyleBackgroundColor = String(style.backgroundColor || selected.backgroundColor || '');
      this.properties.conditionalStyleBorderColor = String(style.borderColor || selected.borderColor || '');
      this.properties.conditionalStyleTextColor = String(style.textColor || style.color || selected.textColor || selected.color || '');
      this.properties.conditionalStyleFontFamily = String(style.fontFamily || selected.fontFamily || '');
      this.properties.conditionalStyleFontSize = String(style.fontSize || selected.fontSize || '');
      this.properties.conditionalStyleFontStyle = String(style.fontStyle || selected.fontStyle || '');
      this.properties.conditionalStyleFontWeight = String(style.fontWeight || selected.fontWeight || '');
      var styleSourceList = String(this.properties.conditionalStyleSourceList || this.getPrimaryListName());
      if (styleSourceList) {
        this.loadListFields(styleSourceList);
      }
      this._isEditingConditionalStyle = true;
      this._conditionalStyleDesignerRevision += 1;
      this._conditionalStyleDesignerMessage = strings.ConditionalStyleLoadSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && (error as any).message ? (error as any).message : strings.ConditionalStyleActionFailed;
    }
    this.context.propertyPane.refresh();
  }

  private handleUpdateConditionalStyle(): void {
    try {
      var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
      var entries = this.parseConditionalStyleJsonArray();
      if (isNaN(index) || index < 0 || index >= entries.length) {
        throw new Error(strings.ConditionalStyleSelectValid);
      }
      entries[index] = this.createConditionalStyleEntry();
      this.persistConditionalStyles(entries);
      this._conditionalStyleDesignerMessage = strings.ConditionalStyleUpdateSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && (error as any).message ? (error as any).message : strings.ConditionalStyleActionFailed;
    }
    this.context.propertyPane.refresh();
  }

  private handleRemoveConditionalStyle(): void {
    try {
      var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
      var entries = this.parseConditionalStyleJsonArray();
      if (isNaN(index) || index < 0 || index >= entries.length) {
        throw new Error(strings.ConditionalStyleSelectValid);
      }
      entries.splice(index, 1);
      this.persistConditionalStyles(entries);
      this._isEditingConditionalStyle = false;
      this.resetConditionalStyleDesigner();
      this._conditionalStyleDesignerMessage = strings.ConditionalStyleRemoveSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && (error as any).message ? (error as any).message : strings.ConditionalStyleActionFailed;
    }
    this.context.propertyPane.refresh();
  }

  private moveConditionalStyle(direction: number): void {
    try {
      var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
      var entries = this.parseConditionalStyleJsonArray();
      var target = index + direction;
      if (isNaN(index) || index < 0 || index >= entries.length || target < 0 || target >= entries.length) {
        throw new Error(strings.ConditionalStyleMoveBoundary);
      }
      var current = entries[index];
      entries[index] = entries[target];
      entries[target] = current;
      this.properties.conditionalStyleSelectedIndex = String(target);
      this.persistConditionalStyles(entries);
      this._conditionalStyleDesignerMessage = direction < 0 ? strings.ConditionalStyleMoveUpSuccess : strings.ConditionalStyleMoveDownSuccess;
    } catch (error) {
      this._conditionalStyleDesignerMessage = error && (error as any).message ? (error as any).message : strings.ConditionalStyleActionFailed;
    }
    this.context.propertyPane.refresh();
  }

  private handleResetConditionalStyles(): void {
    this.persistConditionalStyles([]);
    this._isEditingConditionalStyle = false;
    this.resetConditionalStyleDesigner();
    this._conditionalStyleDesignerMessage = strings.ConditionalStyleResetSuccess;
    this.context.propertyPane.refresh();
  }

  private handleValidateConditionalStyles(): void {
    this._conditionalStyleJsonValidationMessage = this.validateJson(this.properties.conditionalStyleJson || '') || strings.JsonValidationSuccess;
    this.context.propertyPane.refresh();
  }

  private handleConditionalStyleColorChange(propertyPath: string, oldValue: string, newValue: string): void {
    (this.properties as any)[propertyPath] = newValue;
  }

  private handleAppearanceColorChange(propertyPath: string, oldValue: any, newValue: any): void {
    (this.properties as any)[propertyPath] = newValue;
    this.render();
  }

  private getColorPickerValue(value: any, fallback: string): string {
    var color = String(value || '').trim();
    return /^#[0-9a-f]{6}$/i.test(color) || /^#[0-9a-f]{3}$/i.test(color) ? color : fallback;
  }

  private renderDataSourceBackgroundColor(field: any, value: any,
    onUpdate: (fieldId: string, fieldValue: any) => void): React.ReactElement<any> {
    var color = this.getColorPickerValue(value, '#3788d8');
    return React.createElement('div', { style: { display: 'flex', alignItems: 'center' } },
      React.createElement('input', {
        type: 'color',
        value: color,
        title: strings.DataSourceBackgroundColorLabel,
        'aria-label': strings.DataSourceBackgroundColorLabel,
        style: { width: '44px', height: '32px', padding: '2px', border: '1px solid #8a8886', cursor: 'pointer' },
        onChange: (event: any) => onUpdate(field.id, String(event.target.value || color))
      }),
      React.createElement('span', { style: { marginLeft: '8px', fontSize: '12px' } }, color));
  }

  private getFontFamilyOptions(): IPropertyPaneDropdownOption[] {
    return [
      { key: '', text: strings.AppearanceDefaultLabel },
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

  private getFontSizeOptions(): IPropertyPaneDropdownOption[] {
    return [
      { key: '', text: strings.AppearanceDefaultLabel },
      { key: '11px', text: '11px' },
      { key: '12px', text: '12px' },
      { key: '14px', text: '14px' },
      { key: '16px', text: '16px' },
      { key: '18px', text: '18px' },
      { key: '20px', text: '20px' },
      { key: '24px', text: '24px' },
      { key: '28px', text: '28px' },
      { key: '32px', text: '32px' },
      { key: '36px', text: '36px' }
    ];
  }

  private getFontStyleOptions(includeDefault: boolean): IPropertyPaneDropdownOption[] {
    var options: IPropertyPaneDropdownOption[] = [
      { key: 'normal', text: strings.FontStyleNormal },
      { key: 'italic', text: strings.FontStyleItalic },
      { key: 'oblique', text: strings.FontStyleOblique },
      { key: 'initial', text: strings.FontStyleInitial },
      { key: 'inherit', text: strings.FontStyleInherit },
      { key: 'unset', text: strings.FontStyleUnset }
    ];
    if (includeDefault) {
      options.unshift({ key: '', text: strings.AppearanceDefaultLabel });
    }
    return options;
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }

    console.log('[CalendarWebPart] ' + message);
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

  private handleEventSelectionChange(itemId: number, mode: string): void {
    this._selectedItemId = itemId > 0 ? itemId : 0;
    this._selectedMode = mode || 'view';
    this.logDiagnostic('Event selection changed. itemId=' + String(this._selectedItemId));
    this.notifyDynamicData('selectedItemId');
    this.notifyDynamicData('selectedMode');
  }

  private getWebPartVersion(): string {
    const solutionVersion = packageSolutionConfig && packageSolutionConfig.solution
      ? String(packageSolutionConfig.solution.version || '')
      : '';
    if (solutionVersion) {
      return solutionVersion;
    }

    const manifestVersion = this.context && this.context.manifest ? String(this.context.manifest.version || '') : '';
    if (manifestVersion && manifestVersion !== '*') {
      return manifestVersion;
    }

    return 'Unknown';
  }

  private async getJsonWithAcceptFallback(url: string): Promise<any> {
    let response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=verbose' }
      });
    }

    if (!response.ok) {
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      });
    }

    if (!response.ok) {
      throw new Error('Request failed. HTTP ' + String(response.status) + ' ' + response.statusText);
    }

    return response.json();
  }

  private getCalendarDateFields(viewQuery: string): { startFieldName: string; endFieldName: string } | undefined {
    if (!viewQuery || typeof DOMParser === 'undefined') {
      return undefined;
    }
    try {
      var queryDocument = new DOMParser().parseFromString('<ViewQuery>' + viewQuery + '</ViewQuery>', 'text/xml');
      var overlapNodes = queryDocument.getElementsByTagName('DateRangesOverlap');
      if (!overlapNodes.length) {
        return undefined;
      }
      var fieldReferences = overlapNodes[0].getElementsByTagName('FieldRef');
      var startFieldName = fieldReferences.length > 0 ? String(fieldReferences[0].getAttribute('Name') || '') : '';
      var endFieldName = fieldReferences.length > 1 ? String(fieldReferences[1].getAttribute('Name') || '') : '';
      return startFieldName && endFieldName ? {
        startFieldName: startFieldName,
        endFieldName: endFieldName
      } : undefined;
    } catch (_error) {
      return undefined;
    }
  }

  private async loadLists(): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      this.logDiagnostic('Loading Events lists and lists with calendar views from web URL: ' + webUrl);
      const data = await this.getJsonWithAcceptFallback(
        webUrl + "/_api/web/lists?$select=Id,Title,Hidden,BaseTemplate,BaseType&$filter=Hidden eq false and BaseType eq 0"
      );
      const lists = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      var calendarLists: any[] = [];
      this._calendarDateFieldsByListTitle = {};
      for (var listIndex = 0; listIndex < lists.length; listIndex += 1) {
        var list = lists[listIndex];
        var isEventsList = Number(list.BaseTemplate) === 106;
        var hasCalendarView = false;
        if (!isEventsList && list.Id) {
          try {
            var listId = String(list.Id).replace(/[{}]/g, '');
            var viewData = await this.getJsonWithAcceptFallback(
              webUrl + "/_api/web/lists(guid'" + listId + "')/views?$select=Title,ViewType,ViewType2,ViewQuery"
            );
            var views = viewData && viewData.value ? viewData.value
              : (viewData && viewData.d && viewData.d.results ? viewData.d.results : []);
            hasCalendarView = views.some((view: any) => {
              var viewType = String(view.ViewType || '').toLowerCase();
              var viewType2 = String(view.ViewType2 || '').toLowerCase();
              var isCalendarView = viewType === 'calendar'
                || viewType2 === 'calendar'
                || viewType2 === 'moderncalendar';
              if (isCalendarView) {
                this.logDiagnostic('Calendar view detected for list "' + String(list.Title) + '": "'
                  + String(view.Title || '') + '" (ViewType=' + viewType + ', ViewType2=' + viewType2 + ')');
                var inferredDateFields = this.getCalendarDateFields(String(view.ViewQuery || ''));
                if (inferredDateFields) {
                  this._calendarDateFieldsByListTitle[String(list.Title || '').toLowerCase()] = inferredDateFields;
                  this.logDiagnostic('Using calendar view date fields for list "' + String(list.Title) + '": '
                    + inferredDateFields.startFieldName + ' / ' + inferredDateFields.endFieldName);
                }
              }
              return isCalendarView;
            });
          } catch (viewError) {
            this.logDiagnostic('Failed to inspect views for list "' + String(list.Title) + '": '
              + String(viewError && viewError.message ? viewError.message : viewError));
          }
        }
        if (isEventsList || hasCalendarView) {
          calendarLists.push(list);
        }
      }

      this._listBaseTemplateByTitle = {};
      var dateOptionsByName: { [internalName: string]: IDropdownOption } = {};
      for (var calendarListIndex = 0; calendarListIndex < calendarLists.length; calendarListIndex += 1) {
        var calendarList = calendarLists[calendarListIndex];
        this._listBaseTemplateByTitle[String(calendarList.Title || '').toLowerCase()] = Number(calendarList.BaseTemplate || 0);
        try {
          var fieldData = await this.getJsonWithAcceptFallback(
            webUrl + "/_api/web/lists/getByTitle('" + String(calendarList.Title || '').replace(/'/g, "''")
            + "')/fields?$select=InternalName,Title,TypeAsString,Hidden&$filter=Hidden eq false"
          );
          var dateFields = fieldData && fieldData.value ? fieldData.value
            : (fieldData && fieldData.d && fieldData.d.results ? fieldData.d.results : []);
          dateFields.filter((field: any) => String(field.TypeAsString || '').toLowerCase() === 'datetime')
            .forEach((field: any) => {
              var internalName = String(field.InternalName || '');
              if (internalName && !dateOptionsByName[internalName]) {
                dateOptionsByName[internalName] = {
                  key: internalName,
                  text: String(field.Title || internalName) + ' (' + internalName + ')'
                };
              }
            });
        } catch (fieldError) {
          this.logDiagnostic('Failed to load DateTime fields for list "' + String(calendarList.Title) + '": '
            + String(fieldError && fieldError.message ? fieldError.message : fieldError));
        }
      }
      this._dateFieldOptions = [{ key: '__eventsDefault__', text: strings.DataSourceEventsDateFieldsOption }]
        .concat(Object.keys(dateOptionsByName).map((internalName: string) => dateOptionsByName[internalName])
          .sort((left: IDropdownOption, right: IDropdownOption) => left.text.localeCompare(right.text)));
      this._lists = calendarLists.map((calendarListOption: any) => {
        var typeLabel = Number(calendarListOption.BaseTemplate) === 106
          ? strings.DataSourceEventsListType
          : strings.DataSourceCalendarViewListType;
        return { key: calendarListOption.Title, text: String(calendarListOption.Title) + ' (' + typeLabel + ')' };
      }).sort((left: IDropdownOption, right: IDropdownOption) => left.text.localeCompare(right.text));
      this.logDiagnostic('Loaded calendar-capable lists successfully. Count=' + String(this._lists.length));
      this.context.propertyPane.refresh();
    } catch (error) {
      this.logDiagnostic('Failed to load lists: ' + (error && error.message ? error.message : String(error)));
      this._lists = [];
      this._dateFieldOptions = [];
      this._listBaseTemplateByTitle = {};
      this._calendarDateFieldsByListTitle = {};
    }
  }

  private async loadSitePages(): Promise<void> {
    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      var libraryData = await this.getJsonWithAcceptFallback(
        webUrl + '/_api/web/lists?$select=Id,Title,BaseTemplate&$filter=(BaseTemplate eq 119 or BaseTemplate eq 850)'
      );
      var libraries = libraryData && libraryData.value ? libraryData.value
        : (libraryData && libraryData.d && libraryData.d.results ? libraryData.d.results : []);
      if (!libraries.length || !libraries[0].Id) {
        this._sitePages = [{ key: '__defaultForm__', text: strings.DataSourceDefaultFormOption }];
        this.context.propertyPane.refresh();
        return;
      }
      var pageOptionsByUrl: { [url: string]: IDropdownOption } = {};
      for (var libraryIndex = 0; libraryIndex < libraries.length; libraryIndex += 1) {
        var libraryId = String(libraries[libraryIndex].Id).replace(/[{}]/g, '');
        try {
          var data = await this.getJsonWithAcceptFallback(
            webUrl + "/_api/web/lists(guid'" + libraryId + "')/items?$select=File/Name,File/ServerRelativeUrl&$expand=File&$top=5000"
          );
          var pages = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
          pages.filter((page: any) => page.File && page.File.ServerRelativeUrl
            && /\.aspx(?:$|[?#])/i.test(String(page.File.ServerRelativeUrl)))
            .forEach((page: any) => {
              var fileRef = String(page.File.ServerRelativeUrl);
              var title = String(page.File.Name || fileRef);
              pageOptionsByUrl[fileRef.toLowerCase()] = { key: fileRef, text: title + ' (' + fileRef + ')' };
            });
        } catch (pageError) {
          this.logDiagnostic('Failed to load pages from library "' + String(libraries[libraryIndex].Title || '') + '": '
            + String(pageError && pageError.message ? pageError.message : pageError));
        }
      }
      this._sitePages = [{ key: '__defaultForm__', text: strings.DataSourceDefaultFormOption }]
        .concat(Object.keys(pageOptionsByUrl).map((url: string) => pageOptionsByUrl[url])
          .sort((left: IDropdownOption, right: IDropdownOption) => left.text.localeCompare(right.text)));
      this.context.propertyPane.refresh();
    } catch (error) {
      this.logDiagnostic('Failed to load Site Pages: ' + (error && (error as any).message ? (error as any).message : String(error)));
      this._sitePages = [{ key: '__defaultForm__', text: strings.DataSourceDefaultFormOption }];
    }
  }

  private async loadListFields(listName: string): Promise<void> {
    if (!listName) {
      this._listFields = [];
      this._fieldTypeByInternalName = {};
      this._fieldMetadataByInternalName = {};
      this._fieldLookupListByInternalName = {};
      return;
    }

    try {
      const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const data = await this.getJsonWithAcceptFallback(
        webUrl + "/_api/web/lists/getByTitle('" + String(listName).replace(/'/g, "''") + "')/fields"
        + "?$select=InternalName,Title,TypeAsString,RichText,SchemaXml,LookupList,Hidden,ReadOnlyField,Sealed&$filter=Hidden eq false"
      );
      const fields = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      this._fieldTypeByInternalName = {};
      this._fieldMetadataByInternalName = {};
      this._fieldLookupListByInternalName = {};
      fields.forEach((field: any) => {
        const internalName = String(field.InternalName || '');
        if (!internalName) {
          return;
        }
        const fieldType = String(field.TypeAsString || '').toLowerCase();
        const schemaXml = String(field.SchemaXml || '');
        const isRichText = field.RichText === true || String(field.RichText).toLowerCase() === 'true'
          || /\bRichText\s*=\s*["']TRUE["']/i.test(schemaXml);
        this._fieldTypeByInternalName[internalName] = fieldType;
        this._fieldMetadataByInternalName[internalName.toLowerCase()] = {
          type: fieldType,
          richText: fieldType === 'note' && isRichText
        };
        const lookupListId = field.LookupList ? String(field.LookupList).replace(/[{}]/g, '') : '';
        if (lookupListId) {
          this._fieldLookupListByInternalName[internalName] = lookupListId;
        }
      });
      this._listFields = fields
        .filter((field: any) => {
          const fieldType = String(field.TypeAsString || '').toLowerCase();
          const isConfigurable = (field.ReadOnlyField !== true && field.Sealed !== true) || field.InternalName === 'ID';
          return field.InternalName && isConfigurable && fieldType !== 'attachments' && fieldType !== 'computed';
        })
        .map((field: any) => {
          return {
            key: String(field.InternalName),
            text: String(field.Title || field.InternalName) + ' (' + String(field.InternalName) + ')'
          };
        })
        .sort((left: IDropdownOption, right: IDropdownOption) => left.text.localeCompare(right.text));
      this.render();
      this.context.propertyPane.refresh();
    } catch (error) {
      this.logDiagnostic('Failed to load list fields: ' + (error && (error as any).message ? (error as any).message : String(error)));
      this._listFields = [];
      this._fieldTypeByInternalName = {};
      this._fieldMetadataByInternalName = {};
      this._fieldLookupListByInternalName = {};
    }
  }

  private getEventDetailSourceOptions(): IDropdownOption[] {
    var sourceNames: string[] = [];
    var configuredSources = Array.isArray(this.properties.dataSources) ? this.properties.dataSources : [];
    configuredSources.forEach((source: ICalendarDataSource) => {
      var listName = String(source && source.listName || '').trim();
      if (listName && sourceNames.map((name: string) => name.toLowerCase()).indexOf(listName.toLowerCase()) < 0) {
        sourceNames.push(listName);
      }
    });
    if (sourceNames.length === 0 && String(this.properties.listName || '').trim()) {
      sourceNames.push(String(this.properties.listName).trim());
    }
    return sourceNames.map((listName: string) => ({ key: listName, text: listName }));
  }

  private async loadEventDetailFields(listName: string): Promise<void> {
    var normalizedListName = String(listName || '').trim();
    var cacheKey = normalizedListName.toLowerCase();
    if (!normalizedListName || this._eventDetailFieldsByListTitle[cacheKey]) {
      return;
    }
    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      var data = await this.getJsonWithAcceptFallback(
        webUrl + "/_api/web/lists/getByTitle('" + normalizedListName.replace(/'/g, "''") + "')/fields"
        + "?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField,Sealed&$filter=Hidden eq false"
      );
      var fields = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      this._eventDetailFieldsByListTitle[cacheKey] = fields
        .filter((field: any) => {
          var fieldType = String(field.TypeAsString || '').toLowerCase();
          var isConfigurable = (field.ReadOnlyField !== true && field.Sealed !== true) || field.InternalName === 'ID';
          return field.InternalName && isConfigurable && fieldType !== 'attachments' && fieldType !== 'computed';
        })
        .map((field: any) => ({
          key: String(field.InternalName),
          text: String(field.Title || field.InternalName) + ' (' + String(field.InternalName) + ')'
        }))
        .sort((left: IDropdownOption, right: IDropdownOption) => left.text.localeCompare(right.text));
    } catch (error) {
      this.logDiagnostic('Failed to load event detail fields for "' + normalizedListName + '": '
        + (error && (error as any).message ? (error as any).message : String(error)));
      this._eventDetailFieldsByListTitle[cacheKey] = [];
    }
  }

  private async loadAllEventDetailFields(): Promise<void> {
    var sourceOptions = this.getEventDetailSourceOptions();
    await Promise.all(sourceOptions.map((source: IDropdownOption) => this.loadEventDetailFields(String(source.key))));
    this.context.propertyPane.refresh();
  }

  private renderEventDetailFieldSelection(field: any, value: any,
    onUpdate: (fieldId: string, fieldValue: any) => void): React.ReactElement<any> {
    var sourceOptions = this.getEventDetailSourceOptions();
    var primarySource = sourceOptions.length > 0 ? String(sourceOptions[0].key) : '';
    var selection = value && typeof value === 'object'
      ? value
      : { sourceListName: primarySource, field: String(value || '') };
    var sourceListName = String(selection.sourceListName || primarySource);
    var selectedField = String(selection.field || '');
    var fieldOptions = this._eventDetailFieldsByListTitle[sourceListName.toLowerCase()] || [];
    var selectStyle: any = { width: '100%', minWidth: '150px', height: '32px', marginBottom: '6px' };
    var sourceOptionElements: any[] = sourceOptions.map((option: IDropdownOption) => React.createElement('option' as any, {
      key: option.key,
      value: option.key
    }, option.text));
    var fieldOptionElements: any[] = [
      React.createElement('option' as any, { key: '', value: '' }, strings.EventDetailsSelectFieldOption)
    ].concat(fieldOptions.map((option: IDropdownOption) => React.createElement('option' as any, {
      key: option.key,
      value: option.key
    }, option.text)));
    return React.createElement('div' as any, undefined,
      React.createElement('select' as any, {
        value: sourceListName,
        title: strings.EventDetailsSourceLabel,
        'aria-label': strings.EventDetailsSourceLabel,
        style: selectStyle,
        onChange: (event: any) => {
          var nextSource = String(event.target.value || '');
          onUpdate(field.id, { sourceListName: nextSource, field: '' });
        }
      }, sourceOptionElements),
      React.createElement('select' as any, {
        value: selectedField,
        title: strings.EventDetailsFieldLabel,
        'aria-label': strings.EventDetailsFieldLabel,
        style: selectStyle,
        disabled: !sourceListName,
        onChange: (event: any) => onUpdate(field.id, {
          sourceListName: sourceListName,
          field: String(event.target.value || '')
        })
      }, fieldOptionElements));
  }

  private isLookupTypeField(internalName: string): boolean {
    var typeName = this._fieldTypeByInternalName[internalName || ''];
    return typeName === 'lookup' || typeName === 'lookupmulti';
  }

  private async loadLookupListItems(listId: string): Promise<IDropdownOption[]> {
    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      var data = await this.getJsonWithAcceptFallback(
        webUrl + "/_api/web/lists(guid'" + listId + "')/items?$select=Id,Title&$top=500&$orderby=Title"
      );
      var items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      return items.map((item: any) => {
        var title = item.Title ? String(item.Title) : strings.LookupItemNoTitle;
        return { key: title, text: title + ' (ID: ' + String(item.Id) + ')' };
      });
    } catch (error) {
      this.logDiagnostic('Failed to load lookup items: ' + (error && (error as any).message ? (error as any).message : String(error)));
      return [];
    }
  }

  private async handleLoadFilterLookupItems(): Promise<void> {
    var fieldName = String(this.properties.filterDesignerField || '');
    var listId = this._fieldLookupListByInternalName[fieldName];
    if (!listId) {
      this._filterLookupMessage = strings.LookupFieldUnavailable;
      this._filterLookupItemOptions = [];
      this.context.propertyPane.refresh();
      return;
    }
    this._filterLookupItemOptions = await this.loadLookupListItems(listId);
    this._filterLookupMessage = this._filterLookupItemOptions.length > 0
      ? String(this._filterLookupItemOptions.length) + ' ' + strings.LookupItemsLoaded
      : strings.LookupItemsEmpty;
    this.context.propertyPane.refresh();
  }

  private handleApplyFilterLookupPick(): void {
    var picked = String(this.properties.filterDesignerLookupPick || '');
    if (!picked) {
      this._filterLookupMessage = strings.LookupSelectFirst;
      this.context.propertyPane.refresh();
      return;
    }
    this.properties.filterDesignerValue = picked;
    this._filterLookupMessage = strings.LookupValueApplied;
    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const viewOptions: IPropertyPaneDropdownOption[] = [
      { key: 'dayGridMonth', text: strings.ViewMonthLabel },
      { key: 'timeGridWeek', text: strings.ViewWeekLabel },
      { key: 'timeGridDay', text: strings.ViewDayLabel },
      { key: 'listWeek', text: strings.ViewListLabel }
    ];
    const listOptions: IPropertyPaneDropdownOption[] = this._lists.map((list) => {
      return { key: list.key, text: list.text };
    });
    const savedLinkTargetPageUrl = String(this.properties.linkTargetPageUrl || '');
    const linkTargetPageOptions: IDropdownOption[] = this._sitePages.slice();
    if (savedLinkTargetPageUrl && savedLinkTargetPageUrl !== '__defaultForm__'
      && !linkTargetPageOptions.some((page: IDropdownOption) => String(page.key) === savedLinkTargetPageUrl)) {
      linkTargetPageOptions.push({ key: savedLinkTargetPageUrl, text: savedLinkTargetPageUrl + ' (saved URL)' });
    }
    const dataSourceScopeOptions: IPropertyPaneDropdownOption[] = [
      { key: '', text: strings.DataSourceAllOption }
    ].concat((Array.isArray(this.properties.dataSources) ? this.properties.dataSources : [])
      .filter((source: ICalendarDataSource) => !!String(source && source.listName || '').trim())
      .map((source: ICalendarDataSource) => ({
        key: String(source.listName),
        text: String(source.listName)
      })));
    const laneFieldOptions: IPropertyPaneDropdownOption[] = [{ key: '', text: strings.SwimlaneFieldNone }]
      .concat(this._listFields.map((field) => {
        return { key: field.key, text: field.text };
      }));
    const conditionalFieldOptions: IPropertyPaneDropdownOption[] = [{ key: '', text: strings.ConditionalStyleFieldNone }]
      .concat(this._listFields.map((field) => {
        return { key: field.key, text: field.text };
      }));
    const filterFieldOptions: IPropertyPaneDropdownOption[] = [{ key: '', text: strings.FilterDesignerFieldNone }]
      .concat(this._listFields.map((field) => {
        return { key: field.key, text: field.text };
      }));
    let filterEntries: any[] = [];
    let filterSummary = strings.FilterSummaryNone;
    try {
      filterEntries = this.parseFilterJsonArray();
      if (filterEntries.length > 0) {
        filterSummary = String(filterEntries.length) + ' ' + strings.FilterSummaryConditionsLabel;
      }
    } catch (_filterError) {
      filterSummary = strings.FilterSummaryInvalid;
    }
    const filterEntryOptions: IPropertyPaneDropdownOption[] = filterEntries.length > 0
      ? filterEntries.map((entry: any, index: number) => {
        return {
          key: String(index),
          text: String(index + 1) + '. [' + String(entry.sourceListName || strings.DataSourceAllShortLabel) + '] '
            + String(entry.field || '?') + ' '
            + String(entry.operator || 'contains') + ' ' + String(entry.value === undefined ? '' : entry.value)
        };
      })
      : [{ key: '', text: strings.FilterDesignerExistingNone }];
    let conditionalEntries: any[] = [];
    let conditionalSummary = strings.ConditionalStyleSummaryNone;
    try {
      conditionalEntries = this.parseConditionalStyleJsonArray();
      if (conditionalEntries.length > 0) {
        const ruleNames: string[] = [];
        conditionalEntries.forEach((entry: any) => {
          const ruleName = String(entry.rule || entry.ruleName || strings.ConditionalStyleUnnamedRule);
          const scopedRuleName = String(entry.sourceListName || '') + '|' + ruleName;
          if (ruleNames.indexOf(scopedRuleName) < 0) {
            ruleNames.push(scopedRuleName);
          }
        });
        conditionalSummary = String(ruleNames.length) + ' ' + strings.ConditionalStyleRulesLabel
          + ', ' + String(conditionalEntries.length) + ' ' + strings.ConditionalStyleConditionsLabel;
      }
    } catch (_error) {
      conditionalSummary = strings.ConditionalStyleSummaryInvalid;
    }
    const conditionalEntryOptions: IPropertyPaneDropdownOption[] = conditionalEntries.length > 0
      ? conditionalEntries.map((entry: any, index: number) => {
        return {
          key: String(index),
          text: String(index + 1) + '. [' + String(entry.sourceListName || strings.DataSourceAllShortLabel) + '] '
            + String(entry.rule || entry.ruleName || strings.ConditionalStyleUnnamedRule)
            + ': ' + String(entry.conditionField || entry.field || '?') + ' '
            + String(entry.operator || 'eq') + ' ' + String(entry.value === undefined ? '' : entry.value)
        };
      })
      : [{ key: '', text: strings.ConditionalStyleExistingNone }];
    const fontFamilyOptions = this.getFontFamilyOptions();
    const fontSizeOptions = this.getFontSizeOptions();

    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: 'Version: ' + this.getWebPartVersion(),
              groupFields: [
                PropertyPaneLabel('propertyPaneVersionInfo', {
                  text: ' '
                })
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true,
                  resizable: true,
                  rows: 3
                })
              ]
            },
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyFieldCollectionData('dataSources', {
                  key: 'calendarDataSourcesCollection',
                  label: strings.DataSourcesCollectionLabel,
                  panelHeader: strings.DataSourcesCollectionHeader,
                  panelDescription: strings.DataSourcesCollectionDescription,
                  manageBtnLabel: strings.DataSourcesCollectionManageButton,
                  value: Array.isArray(this.properties.dataSources) ? this.properties.dataSources : [],
                  enableSorting: true,
                  fields: [
                    {
                      id: 'listName',
                      title: strings.DataSourceListLabel,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: listOptions
                    },
                    {
                      id: 'startFieldName',
                      title: strings.DataSourceStartFieldLabel,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      defaultValue: '__eventsDefault__',
                      options: this._dateFieldOptions.map((field: IDropdownOption) => ({ key: field.key, text: field.text }))
                    },
                    {
                      id: 'endFieldName',
                      title: strings.DataSourceEndFieldLabel,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      defaultValue: '__eventsDefault__',
                      options: this._dateFieldOptions.map((field: IDropdownOption) => ({ key: field.key, text: field.text }))
                    },
                    {
                      id: 'backgroundColor',
                      title: strings.DataSourceBackgroundColorLabel,
                      type: CustomCollectionFieldType.custom,
                      defaultValue: '#3788d8',
                      onCustomRender: this.renderDataSourceBackgroundColor.bind(this)
                    }
                  ]
                }),
                PropertyPaneDropdown('defaultView', {
                  label: strings.DefaultViewFieldLabel,
                  options: viewOptions,
                  selectedKey: this.properties.defaultView || 'dayGridMonth'
                }),
                PropertyPaneToggle('showWeekends', {
                  label: strings.ShowWeekendsFieldLabel,
                  checked: this.properties.showWeekends !== false
                }),
                PropertyPaneSlider('calendarHeight', {
                  label: strings.CalendarHeightFieldLabel,
                  min: 300,
                  max: 1200,
                  step: 20,
                  value: typeof this.properties.calendarHeight === 'number' ? this.properties.calendarHeight : 600
                }),
                PropertyPaneToggle('enableSwimlanes', {
                  label: strings.EnableSwimlanesLabel,
                  checked: this.properties.enableSwimlanes === true
                }),
                PropertyPaneDropdown('swimlaneFieldName', {
                  label: strings.SwimlaneFieldLabel,
                  options: laneFieldOptions,
                  selectedKey: this.properties.swimlaneFieldName || '',
                  disabled: this.properties.enableSwimlanes !== true
                }),
                PropertyPaneDropdown('swimlaneDays', {
                  label: strings.SwimlaneDaysLabel,
                  options: [
                    { key: 7, text: strings.SwimlaneDaysSeven },
                    { key: 14, text: strings.SwimlaneDaysFourteen },
                    { key: 30, text: strings.SwimlaneDaysThirty }
                  ],
                  selectedKey: typeof this.properties.swimlaneDays === 'number' ? this.properties.swimlaneDays : 7,
                  disabled: this.properties.enableSwimlanes !== true
                }),
                PropertyPaneTextField('swimlaneUnassignedLabel', {
                  label: strings.SwimlaneUnassignedLabel,
                  placeholder: strings.SwimlaneUnassignedDefault,
                  disabled: this.properties.enableSwimlanes !== true
                })
              ]
            },
            {
              groupName: strings.LinkGroupName,
              groupFields: [
                PropertyPaneToggle('enableEventDetails', {
                  label: strings.EnableEventDetailsLabel,
                  checked: this.properties.enableEventDetails !== false
                }),
                PropertyFieldCollectionData('eventDetailsFields', {
                  key: 'eventDetailsFieldsCollection',
                  label: strings.EventDetailsCollectionLabel,
                  panelHeader: strings.EventDetailsCollectionHeader,
                  panelDescription: strings.EventDetailsCollectionDescription,
                  manageBtnLabel: strings.EventDetailsCollectionManageButton,
                  value: Array.isArray(this.properties.eventDetailsFields) ? this.properties.eventDetailsFields : [],
                  enableSorting: true,
                  disabled: this.properties.enableEventDetails === false,
                  fields: [
                    {
                      id: 'field',
                      title: strings.EventDetailsSourceAndFieldLabel,
                      type: CustomCollectionFieldType.custom,
                      required: true,
                      onCustomRender: this.renderEventDetailFieldSelection.bind(this)
                    },
                    {
                      id: 'label',
                      title: strings.EventDetailsLabelLabel,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'showLabel',
                      title: strings.EventDetailsShowLabelLabel,
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: true
                    },
                    {
                      id: 'format',
                      title: strings.EventDetailsFormatLabel,
                      type: CustomCollectionFieldType.dropdown,
                      defaultValue: 'auto',
                      options: [
                        { key: 'auto', text: strings.EventDetailsFormatAuto },
                        { key: 'text', text: strings.EventDetailsFormatText },
                        { key: 'date', text: strings.EventDetailsFormatDate },
                        { key: 'datetime', text: strings.EventDetailsFormatDateTime },
                        { key: 'time', text: strings.EventDetailsFormatTime },
                        { key: 'yesno', text: strings.EventDetailsFormatYesNo }
                      ]
                    },
                    {
                      id: 'hideWhenEmpty',
                      title: strings.EventDetailsHideEmptyLabel,
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: false
                    }
                  ]
                }),
                PropertyPaneCheckbox('showLinkToItem', {
                  text: strings.ShowLinkToItemLabel,
                  checked: this.properties.showLinkToItem === true
                }),
                PropertyPaneDropdown('linkTargetPageUrl', {
                  label: strings.LinkTargetPageUrlLabel,
                  options: linkTargetPageOptions,
                  selectedKey: this.properties.linkTargetPageUrl || '__defaultForm__',
                  disabled: this.properties.showLinkToItem !== true
                }),
                PropertyPaneTextField('linkTargetIdParam', {
                  label: strings.LinkTargetIdParamLabel,
                  placeholder: 'itemid',
                  disabled: this.properties.showLinkToItem !== true
                }),
                PropertyPaneCheckbox('includeReturnUrlParam', {
                  text: strings.IncludeReturnUrlParamLabel,
                  checked: this.properties.includeReturnUrlParam === true,
                  disabled: this.properties.showLinkToItem !== true
                })
              ]
            },
            {
              groupName: strings.DiagnosticsGroupName,
              groupFields: [
                PropertyPaneCheckbox('enableDiagnostics', {
                  text: strings.PropEnableDiagnosticsLabel,
                  checked: this.properties.enableDiagnostics !== false
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.AppearancePageDescription
          },
          groups: [
            {
              groupName: strings.CalendarSurfaceGroupName,
              groupFields: [
                PropertyFieldColorPicker('calendarBackgroundColor', {
                  label: strings.CalendarBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.calendarBackgroundColor, '#ffffff'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'calendarBackgroundColorField'
                }),
                PropertyFieldColorPicker('webPartBorderColor', {
                  label: strings.WebPartBorderColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.webPartBorderColor, '#d2d0ce'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'webPartBorderColorField'
                }),
                PropertyPaneSlider('webPartBorderWidth', {
                  label: strings.WebPartBorderWidthLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.webPartBorderWidth === 'number' ? this.properties.webPartBorderWidth : 0
                }),
                PropertyPaneSlider('webPartCornerRadius', {
                  label: strings.WebPartCornerRadiusLabel,
                  min: 0,
                  max: 40,
                  step: 1,
                  value: typeof this.properties.webPartCornerRadius === 'number' ? this.properties.webPartCornerRadius : 0
                })
              ]
            },
            {
              groupName: strings.WebPartTitleAppearanceGroupName,
              groupFields: [
                PropertyFieldColorPicker('webPartTitleTextColor', {
                  label: strings.TitleTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.webPartTitleTextColor, '#323130'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'webPartTitleTextColorField'
                }),
                PropertyFieldColorPicker('webPartTitleBackgroundColor', {
                  label: strings.TitleBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.webPartTitleBackgroundColor, '#ffffff'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'webPartTitleBackgroundColorField'
                }),
                PropertyPaneDropdown('webPartTitleFontFamily', {
                  label: strings.TitleFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.webPartTitleFontFamily || ''
                }),
                PropertyPaneDropdown('webPartTitleFontSize', {
                  label: strings.TitleFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.webPartTitleFontSize || '20px'
                }),
                PropertyPaneDropdown('webPartTitleFontStyle', {
                  label: strings.TitleFontStyleLabel,
                  options: this.getFontStyleOptions(false),
                  selectedKey: this.properties.webPartTitleFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('webPartTitleFontBold', {
                  text: strings.TitleFontBoldLabel,
                  checked: this.properties.webPartTitleFontBold !== false
                }),
                PropertyPaneDropdown('webPartTitleAlignment', {
                  label: strings.TitleAlignmentLabel,
                  options: [
                    { key: 'left', text: strings.AlignmentLeft },
                    { key: 'center', text: strings.AlignmentCenter },
                    { key: 'right', text: strings.AlignmentRight }
                  ],
                  selectedKey: this.properties.webPartTitleAlignment || 'left'
                }),
                PropertyPaneSlider('webPartTitleMinHeight', {
                  label: strings.TitleMinHeightLabel,
                  min: 0,
                  max: 200,
                  step: 4,
                  value: typeof this.properties.webPartTitleMinHeight === 'number'
                    ? this.properties.webPartTitleMinHeight : 0
                }),
                PropertyFieldColorPicker('webPartDescriptionTextColor', {
                  label: strings.DescriptionTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.webPartDescriptionTextColor, '#605e5c'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'webPartDescriptionTextColorField'
                }),
                PropertyPaneDropdown('webPartDescriptionFontFamily', {
                  label: strings.DescriptionFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.webPartDescriptionFontFamily || ''
                }),
                PropertyPaneDropdown('webPartDescriptionFontSize', {
                  label: strings.DescriptionFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.webPartDescriptionFontSize || '14px'
                }),
                PropertyPaneDropdown('webPartDescriptionFontStyle', {
                  label: strings.DescriptionFontStyleLabel,
                  options: this.getFontStyleOptions(false),
                  selectedKey: this.properties.webPartDescriptionFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('webPartDescriptionFontBold', {
                  text: strings.DescriptionFontBoldLabel,
                  checked: this.properties.webPartDescriptionFontBold === true
                })
              ]
            },
            {
              groupName: strings.CalendarTitleAppearanceGroupName,
              groupFields: [
                PropertyFieldColorPicker('calendarTitleTextColor', {
                  label: strings.TitleTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.calendarTitleTextColor, '#323130'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'calendarTitleTextColorField'
                }),
                PropertyPaneDropdown('calendarTitleFontFamily', {
                  label: strings.TitleFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.calendarTitleFontFamily || ''
                }),
                PropertyPaneDropdown('calendarTitleFontSize', {
                  label: strings.TitleFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.calendarTitleFontSize || '28px'
                }),
                PropertyPaneDropdown('calendarTitleFontStyle', {
                  label: strings.TitleFontStyleLabel,
                  options: this.getFontStyleOptions(false),
                  selectedKey: this.properties.calendarTitleFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('calendarTitleFontBold', {
                  text: strings.TitleFontBoldLabel,
                  checked: this.properties.calendarTitleFontBold !== false
                })
              ]
            },
            {
              groupName: strings.ToolbarButtonAppearanceGroupName,
              groupFields: [
                PropertyFieldColorPicker('toolbarButtonBackgroundColor', {
                  label: strings.ToolbarButtonBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.toolbarButtonBackgroundColor, '#2f4358'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'toolbarButtonBackgroundColorField'
                }),
                PropertyFieldColorPicker('toolbarButtonTextColor', {
                  label: strings.ToolbarButtonTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.toolbarButtonTextColor, '#ffffff'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'toolbarButtonTextColorField'
                }),
                PropertyFieldColorPicker('toolbarButtonBorderColor', {
                  label: strings.ToolbarButtonBorderColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.toolbarButtonBorderColor, '#2f4358'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'toolbarButtonBorderColorField'
                }),
                PropertyFieldColorPicker('toolbarButtonActiveBackgroundColor', {
                  label: strings.ToolbarButtonActiveBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.toolbarButtonActiveBackgroundColor, '#172536'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'toolbarButtonActiveBackgroundColorField'
                }),
                PropertyFieldColorPicker('toolbarButtonActiveTextColor', {
                  label: strings.ToolbarButtonActiveTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.toolbarButtonActiveTextColor, '#ffffff'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'toolbarButtonActiveTextColorField'
                }),
                PropertyPaneDropdown('toolbarButtonFontFamily', {
                  label: strings.ToolbarButtonFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.toolbarButtonFontFamily || ''
                }),
                PropertyPaneDropdown('toolbarButtonFontSize', {
                  label: strings.ToolbarButtonFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.toolbarButtonFontSize || '16px'
                }),
                PropertyPaneDropdown('toolbarButtonFontStyle', {
                  label: strings.ToolbarButtonFontStyleLabel,
                  options: this.getFontStyleOptions(false),
                  selectedKey: this.properties.toolbarButtonFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('toolbarButtonFontBold', {
                  text: strings.ToolbarButtonFontBoldLabel,
                  checked: this.properties.toolbarButtonFontBold === true
                }),
                PropertyPaneSlider('toolbarButtonBorderWidth', {
                  label: strings.ToolbarButtonBorderWidthLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.toolbarButtonBorderWidth === 'number' ? this.properties.toolbarButtonBorderWidth : 1
                }),
                PropertyPaneSlider('toolbarButtonCornerRadius', {
                  label: strings.ToolbarButtonCornerRadiusLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.toolbarButtonCornerRadius === 'number' ? this.properties.toolbarButtonCornerRadius : 4
                })
              ]
            },
            {
              groupName: strings.EventAppearanceGroupName,
              groupFields: [
                PropertyFieldColorPicker('eventBackgroundColor', {
                  label: strings.EventBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.eventBackgroundColor, '#3788d8'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'eventBackgroundColorField'
                }),
                PropertyFieldColorPicker('eventTextColor', {
                  label: strings.EventTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.eventTextColor, '#ffffff'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'eventTextColorField'
                }),
                PropertyFieldColorPicker('eventBorderColor', {
                  label: strings.EventBorderColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.eventBorderColor, '#2c6faa'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'eventBorderColorField'
                }),
                PropertyPaneDropdown('eventFontFamily', {
                  label: strings.EventFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.eventFontFamily || ''
                }),
                PropertyPaneDropdown('eventFontSize', {
                  label: strings.EventFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.eventFontSize || ''
                }),
                PropertyPaneDropdown('eventFontStyle', {
                  label: strings.EventFontStyleLabel,
                  options: this.getFontStyleOptions(false),
                  selectedKey: this.properties.eventFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('eventFontBold', {
                  text: strings.EventFontBoldLabel,
                  checked: this.properties.eventFontBold === true
                }),
                PropertyPaneSlider('eventBorderWidth', {
                  label: strings.EventBorderWidthLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.eventBorderWidth === 'number' ? this.properties.eventBorderWidth : 1
                }),
                PropertyPaneSlider('eventCornerRadius', {
                  label: strings.EventCornerRadiusLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.eventCornerRadius === 'number' ? this.properties.eventCornerRadius : 3
                })
              ]
            },
            {
              groupName: strings.SelectedEventAppearanceGroupName,
              groupFields: [
                PropertyFieldColorPicker('selectedEventBackgroundColor', {
                  label: strings.SelectedEventBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.selectedEventBackgroundColor, '#ffb900'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'selectedEventBackgroundColorField'
                }),
                PropertyFieldColorPicker('selectedEventTextColor', {
                  label: strings.SelectedEventTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.selectedEventTextColor, '#201f1e'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'selectedEventTextColorField'
                }),
                PropertyFieldColorPicker('selectedEventBorderColor', {
                  label: strings.SelectedEventBorderColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.selectedEventBorderColor, '#8a4b00'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'selectedEventBorderColorField'
                }),
                PropertyPaneDropdown('selectedEventFontFamily', {
                  label: strings.SelectedEventFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.selectedEventFontFamily || ''
                }),
                PropertyPaneDropdown('selectedEventFontSize', {
                  label: strings.SelectedEventFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.selectedEventFontSize || ''
                }),
                PropertyPaneDropdown('selectedEventFontStyle', {
                  label: strings.SelectedEventFontStyleLabel,
                  options: this.getFontStyleOptions(false),
                  selectedKey: this.properties.selectedEventFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('selectedEventFontBold', {
                  text: strings.SelectedEventFontBoldLabel,
                  checked: this.properties.selectedEventFontBold !== false
                }),
                PropertyPaneSlider('selectedEventBorderWidth', {
                  label: strings.SelectedEventBorderWidthLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.selectedEventBorderWidth === 'number'
                    ? this.properties.selectedEventBorderWidth : 3
                }),
                PropertyPaneSlider('selectedEventCornerRadius', {
                  label: strings.SelectedEventCornerRadiusLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.selectedEventCornerRadius === 'number'
                    ? this.properties.selectedEventCornerRadius : 5
                })
              ]
            },
            {
              groupName: strings.EventDetailsAppearanceGroupName,
              groupFields: [
                PropertyFieldColorPicker('detailsOverlayColor', {
                  label: strings.DetailsOverlayColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsOverlayColor, '#000000'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsOverlayColorField'
                }),
                PropertyPaneSlider('detailsOverlayOpacity', {
                  label: strings.DetailsOverlayOpacityLabel,
                  min: 0,
                  max: 90,
                  step: 1,
                  value: typeof this.properties.detailsOverlayOpacity === 'number' ? this.properties.detailsOverlayOpacity : 28
                }),
                PropertyFieldColorPicker('detailsPanelBackgroundColor', {
                  label: strings.DetailsPanelBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsPanelBackgroundColor, '#ffffff'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsPanelBackgroundColorField'
                }),
                PropertyFieldColorPicker('detailsPanelBorderColor', {
                  label: strings.DetailsPanelBorderColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsPanelBorderColor, '#8a8886'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsPanelBorderColorField'
                }),
                PropertyPaneSlider('detailsPanelBorderWidth', {
                  label: strings.DetailsPanelBorderWidthLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.detailsPanelBorderWidth === 'number' ? this.properties.detailsPanelBorderWidth : 1
                }),
                PropertyPaneSlider('detailsPanelCornerRadius', {
                  label: strings.DetailsPanelCornerRadiusLabel,
                  min: 0,
                  max: 60,
                  step: 1,
                  value: typeof this.properties.detailsPanelCornerRadius === 'number' ? this.properties.detailsPanelCornerRadius : 6
                }),
                PropertyPaneSlider('detailsPanelMaxWidth', {
                  label: strings.DetailsPanelMaxWidthLabel,
                  min: 320,
                  max: 1200,
                  step: 20,
                  value: typeof this.properties.detailsPanelMaxWidth === 'number' ? this.properties.detailsPanelMaxWidth : 680
                }),
                PropertyPaneSlider('detailsPanelPadding', {
                  label: strings.DetailsPanelPaddingLabel,
                  min: 8,
                  max: 80,
                  step: 2,
                  value: typeof this.properties.detailsPanelPadding === 'number' ? this.properties.detailsPanelPadding : 30
                }),
                PropertyPaneCheckbox('detailsPanelShadow', {
                  text: strings.DetailsPanelShadowLabel,
                  checked: this.properties.detailsPanelShadow !== false
                }),
                PropertyFieldColorPicker('detailsTitleTextColor', {
                  label: strings.DetailsTitleTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsTitleTextColor, '#201f1e'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsTitleTextColorField'
                }),
                PropertyPaneDropdown('detailsTitleFontFamily', {
                  label: strings.DetailsTitleFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.detailsTitleFontFamily || ''
                }),
                PropertyPaneDropdown('detailsTitleFontSize', {
                  label: strings.DetailsTitleFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.detailsTitleFontSize || '24px'
                }),
                PropertyPaneDropdown('detailsTitleFontStyle', {
                  label: strings.DetailsTitleFontStyleLabel,
                  options: this.getFontStyleOptions(false),
                  selectedKey: this.properties.detailsTitleFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('detailsTitleFontBold', {
                  text: strings.DetailsTitleFontBoldLabel,
                  checked: this.properties.detailsTitleFontBold !== false
                }),
                PropertyFieldColorPicker('detailsLabelTextColor', {
                  label: strings.DetailsLabelTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsLabelTextColor, '#323130'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsLabelTextColorField'
                }),
                PropertyPaneDropdown('detailsLabelFontFamily', {
                  label: strings.DetailsLabelFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.detailsLabelFontFamily || ''
                }),
                PropertyPaneDropdown('detailsLabelFontSize', {
                  label: strings.DetailsLabelFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.detailsLabelFontSize || '14px'
                }),
                PropertyPaneDropdown('detailsLabelFontStyle', {
                  label: strings.DetailsLabelFontStyleLabel,
                  options: this.getFontStyleOptions(false),
                  selectedKey: this.properties.detailsLabelFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('detailsLabelFontBold', {
                  text: strings.DetailsLabelFontBoldLabel,
                  checked: this.properties.detailsLabelFontBold !== false
                }),
                PropertyFieldColorPicker('detailsValueTextColor', {
                  label: strings.DetailsValueTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsValueTextColor, '#201f1e'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsValueTextColorField'
                }),
                PropertyPaneDropdown('detailsValueFontFamily', {
                  label: strings.DetailsValueFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.detailsValueFontFamily || ''
                }),
                PropertyPaneDropdown('detailsValueFontSize', {
                  label: strings.DetailsValueFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.detailsValueFontSize || '14px'
                }),
                PropertyPaneDropdown('detailsValueFontStyle', {
                  label: strings.DetailsValueFontStyleLabel,
                  options: this.getFontStyleOptions(false),
                  selectedKey: this.properties.detailsValueFontStyle || 'normal'
                }),
                PropertyPaneCheckbox('detailsValueFontBold', {
                  text: strings.DetailsValueFontBoldLabel,
                  checked: this.properties.detailsValueFontBold === true
                }),
                PropertyFieldColorPicker('detailsDividerColor', {
                  label: strings.DetailsDividerColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsDividerColor, '#edebe9'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsDividerColorField'
                }),
                PropertyPaneSlider('detailsDividerWidth', {
                  label: strings.DetailsDividerWidthLabel,
                  min: 0,
                  max: 10,
                  step: 1,
                  value: typeof this.properties.detailsDividerWidth === 'number' ? this.properties.detailsDividerWidth : 1
                }),
                PropertyFieldColorPicker('detailsCloseButtonColor', {
                  label: strings.DetailsCloseButtonColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsCloseButtonColor, '#323130'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsCloseButtonColorField'
                }),
                PropertyFieldColorPicker('detailsCloseButtonBackgroundColor', {
                  label: strings.DetailsCloseButtonBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsCloseButtonBackgroundColor, '#ffffff'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsCloseButtonBackgroundColorField'
                }),
                PropertyFieldColorPicker('detailsCloseButtonHoverBackgroundColor', {
                  label: strings.DetailsCloseButtonHoverBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.detailsCloseButtonHoverBackgroundColor, '#edebe9'),
                  onPropertyChange: this.handleAppearanceColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'detailsCloseButtonHoverBackgroundColorField'
                }),
                PropertyPaneSlider('detailsCloseButtonCornerRadius', {
                  label: strings.DetailsCloseButtonCornerRadiusLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: typeof this.properties.detailsCloseButtonCornerRadius === 'number'
                    ? this.properties.detailsCloseButtonCornerRadius : 2
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.FilterPageDescription
          },
          groups: [
            {
              groupName: strings.FilterSummaryGroupName,
              groupFields: [
                PropertyPaneLabel('filterSummary', { text: filterSummary })
              ]
            },
            {
              groupName: strings.FilterDesignerGroupName,
              groupFields: [
                PropertyPaneDropdown('filterDesignerSourceList', {
                  label: strings.DataSourceRuleScopeLabel,
                  options: dataSourceScopeOptions,
                  selectedKey: this.properties.filterDesignerSourceList || ''
                }),
                PropertyPaneDropdown('filterDesignerField', {
                  label: strings.FilterDesignerFieldLabel,
                  options: filterFieldOptions,
                  selectedKey: this.properties.filterDesignerField || ''
                }),
                PropertyPaneDropdown('filterDesignerOperator', {
                  label: strings.FilterDesignerOperatorLabel,
                  options: [
                    { key: 'eq', text: strings.OperatorEquals },
                    { key: 'ne', text: strings.OperatorNotEquals },
                    { key: 'contains', text: strings.OperatorContains },
                    { key: 'notcontains', text: strings.OperatorNotContains },
                    { key: 'startswith', text: strings.OperatorStartsWith },
                    { key: 'endswith', text: strings.OperatorEndsWith },
                    { key: 'gt', text: strings.OperatorGreaterThan },
                    { key: 'ge', text: strings.OperatorGreaterThanOrEqual },
                    { key: 'lt', text: strings.OperatorLessThan },
                    { key: 'le', text: strings.OperatorLessThanOrEqual }
                  ],
                  selectedKey: this.properties.filterDesignerOperator || 'contains'
                }),
                PropertyPaneDropdown('filterDesignerLogical', {
                  label: strings.FilterDesignerLogicalLabel,
                  options: [
                    { key: 'and', text: strings.LogicalAnd },
                    { key: 'or', text: strings.LogicalOr }
                  ],
                  selectedKey: this.properties.filterDesignerLogical || 'and'
                }),
                PropertyPaneDropdown('filterDesignerValueType', {
                  label: strings.FilterDesignerValueTypeLabel,
                  options: [
                    { key: 'static', text: strings.ConditionalStyleValueTypeStatic },
                    { key: 'expression', text: strings.ConditionalStyleValueTypeExpression }
                  ],
                  selectedKey: this.properties.filterDesignerValueType || 'static'
                }),
                PropertyPaneTextField('filterDesignerValue', {
                  label: strings.FilterDesignerValueLabel,
                  placeholder: this.properties.filterDesignerValueType === 'expression'
                    ? strings.ConditionalStyleExpressionPlaceholder
                    : strings.FilterDesignerValuePlaceholder,
                  description: this.properties.filterDesignerValueType === 'expression'
                    ? strings.ConditionalStyleExpressionHelp
                    : ''
                }),
                ...(this.properties.filterDesignerValueType !== 'expression'
                  && this.isLookupTypeField(this.properties.filterDesignerField || '') ? [
                  PropertyPaneLabel('filterLookupHelp', { text: strings.LookupPickerHelp }),
                  PropertyPaneButton('loadFilterLookupItems', {
                    text: strings.LookupRefreshButton,
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleLoadFilterLookupItems.bind(this)
                  }),
                  PropertyPaneDropdown('filterDesignerLookupPick', {
                    label: strings.LookupSelectLabel,
                    options: this._filterLookupItemOptions.length > 0
                      ? this._filterLookupItemOptions
                      : [{ key: '', text: strings.LookupItemsNotLoaded }],
                    selectedKey: this.properties.filterDesignerLookupPick || ''
                  }),
                  PropertyPaneButton('applyFilterLookupPick', {
                    text: strings.LookupApplyButton,
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleApplyFilterLookupPick.bind(this)
                  }),
                  PropertyPaneLabel('filterLookupMessage', {
                    text: this._filterLookupMessage || strings.LookupStatusPlaceholder
                  })
                ] : []),
                PropertyPaneButton('addFilter', {
                  text: strings.FilterDesignerAddButton,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.handleAddFilter.bind(this)
                }),
                PropertyPaneDropdown('filterDesignerSelectedIndex', {
                  label: strings.FilterDesignerExistingLabel,
                  options: filterEntryOptions,
                  selectedKey: this.properties.filterDesignerSelectedIndex || ''
                }),
                PropertyPaneButton('loadFilter', {
                  text: strings.FilterDesignerLoadButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleLoadFilter.bind(this)
                }),
                PropertyPaneButton('updateFilter', {
                  text: strings.FilterDesignerUpdateButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleUpdateFilter.bind(this)
                }),
                PropertyPaneButton('removeFilter', {
                  text: strings.FilterDesignerRemoveButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleRemoveFilter.bind(this)
                }),
                PropertyPaneButton('resetFilters', {
                  text: strings.FilterDesignerResetButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleResetFilters.bind(this)
                }),
                PropertyPaneLabel('filterDesignerMessage', {
                  text: this._filterDesignerMessage || strings.FilterDesignerStatusPlaceholder
                })
              ]
            },
            {
              groupName: strings.FilterJsonGroupName,
              groupFields: [
                PropertyPaneTextField('filterJson', {
                  label: strings.FilterJsonLabel,
                  description: strings.FilterJsonDescription,
                  placeholder: strings.FilterJsonPlaceholder,
                  value: this.properties.filterJson || '',
                  multiline: true,
                  resizable: true,
                  rows: 10
                }),
                PropertyPaneButton('validateFilters', {
                  text: strings.FilterValidateButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleValidateFilters.bind(this)
                }),
                PropertyPaneLabel('filterValidationMessage', {
                  text: this._filterJsonValidationMessage || strings.FilterValidationPlaceholder
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.ConditionalStylePageDescription
          },
          groups: [
            {
              groupName: strings.ConditionalStyleSummaryGroupName,
              groupFields: [
                PropertyPaneLabel('conditionalStyleSummary', { text: conditionalSummary })
              ]
            },
            {
              groupName: strings.ConditionalStyleDesignerGroupName,
              groupFields: [
                PropertyPaneDropdown('conditionalStyleSourceList', {
                  label: strings.DataSourceRuleScopeLabel,
                  options: dataSourceScopeOptions,
                  selectedKey: this.properties.conditionalStyleSourceList || ''
                }),
                PropertyPaneTextField('conditionalStyleRuleName', {
                  label: strings.ConditionalStyleRuleNameLabel,
                  placeholder: strings.ConditionalStyleRuleNamePlaceholder,
                  value: this.properties.conditionalStyleRuleName || ''
                }),
                PropertyPaneDropdown('conditionalStyleConditionField', {
                  label: strings.ConditionalStyleConditionFieldLabel,
                  options: conditionalFieldOptions,
                  selectedKey: this.properties.conditionalStyleConditionField || ''
                }),
                PropertyPaneDropdown('conditionalStyleConditionOperator', {
                  label: strings.ConditionalStyleOperatorLabel,
                  options: [
                    { key: 'eq', text: strings.OperatorEquals },
                    { key: 'ne', text: strings.OperatorNotEquals },
                    { key: 'contains', text: strings.OperatorContains },
                    { key: 'notcontains', text: strings.OperatorNotContains },
                    { key: 'startswith', text: strings.OperatorStartsWith },
                    { key: 'endswith', text: strings.OperatorEndsWith },
                    { key: 'gt', text: strings.OperatorGreaterThan },
                    { key: 'ge', text: strings.OperatorGreaterThanOrEqual },
                    { key: 'lt', text: strings.OperatorLessThan },
                    { key: 'le', text: strings.OperatorLessThanOrEqual }
                  ],
                  selectedKey: this.properties.conditionalStyleConditionOperator || 'eq'
                }),
                PropertyPaneDropdown('conditionalStyleConditionLogical', {
                  label: strings.ConditionalStyleLogicalLabel,
                  options: [
                    { key: 'and', text: strings.LogicalAnd },
                    { key: 'or', text: strings.LogicalOr }
                  ],
                  selectedKey: this.properties.conditionalStyleConditionLogical || 'and'
                }),
                PropertyPaneDropdown('conditionalStyleConditionValueType', {
                  label: strings.ConditionalStyleValueTypeLabel,
                  options: [
                    { key: 'static', text: strings.ConditionalStyleValueTypeStatic },
                    { key: 'expression', text: strings.ConditionalStyleValueTypeExpression }
                  ],
                  selectedKey: this.properties.conditionalStyleConditionValueType || 'static'
                }),
                PropertyPaneTextField('conditionalStyleConditionValue', {
                  label: strings.ConditionalStyleValueLabel,
                  placeholder: this.properties.conditionalStyleConditionValueType === 'expression'
                    ? strings.ConditionalStyleExpressionPlaceholder
                    : strings.ConditionalStyleValuePlaceholder,
                  description: this.properties.conditionalStyleConditionValueType === 'expression'
                    ? strings.ConditionalStyleExpressionHelp
                    : ''
                }),
                PropertyPaneCheckbox('conditionalStyleEnabled', {
                  text: strings.ConditionalStyleEnabledLabel,
                  checked: this.properties.conditionalStyleEnabled !== false
                }),
                PropertyPaneTextField('conditionalStylePriority', {
                  label: strings.ConditionalStylePriorityLabel,
                  placeholder: '100',
                  value: this.properties.conditionalStylePriority || '100'
                }),
                PropertyFieldColorPicker('conditionalStyleBackgroundColor', {
                  label: strings.ConditionalStyleBackgroundColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.conditionalStyleBackgroundColor, '#ffffff'),
                  onPropertyChange: this.handleConditionalStyleColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'calendarConditionalBackground-' + String(this._conditionalStyleDesignerRevision)
                }),
                PropertyFieldColorPicker('conditionalStyleBorderColor', {
                  label: strings.ConditionalStyleBorderColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.conditionalStyleBorderColor, '#8a8886'),
                  onPropertyChange: this.handleConditionalStyleColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'calendarConditionalBorder-' + String(this._conditionalStyleDesignerRevision)
                }),
                PropertyFieldColorPicker('conditionalStyleTextColor', {
                  label: strings.ConditionalStyleTextColorLabel,
                  selectedColor: this.getColorPickerValue(this.properties.conditionalStyleTextColor, '#201f1e'),
                  onPropertyChange: this.handleConditionalStyleColorChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'calendarConditionalText-' + String(this._conditionalStyleDesignerRevision)
                }),
                PropertyPaneDropdown('conditionalStyleFontFamily', {
                  label: strings.ConditionalStyleFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.conditionalStyleFontFamily || ''
                }),
                PropertyPaneDropdown('conditionalStyleFontSize', {
                  label: strings.ConditionalStyleFontSizeLabel,
                  options: fontSizeOptions,
                  selectedKey: this.properties.conditionalStyleFontSize || ''
                }),
                PropertyPaneDropdown('conditionalStyleFontStyle', {
                  label: strings.ConditionalStyleFontStyleLabel,
                  options: this.getFontStyleOptions(true),
                  selectedKey: this.properties.conditionalStyleFontStyle || ''
                }),
                PropertyPaneDropdown('conditionalStyleFontWeight', {
                  label: strings.ConditionalStyleFontWeightLabel,
                  options: [
                    { key: '', text: strings.ConditionalStyleDefault },
                    { key: 'normal', text: strings.FontWeightNormal },
                    { key: 'bold', text: strings.FontWeightBold }
                  ],
                  selectedKey: this.properties.conditionalStyleFontWeight || ''
                }),
                PropertyPaneButton('addConditionalStyle', {
                  text: strings.ConditionalStyleAddButton,
                  buttonType: PropertyPaneButtonType.Primary,
                  disabled: this._isEditingConditionalStyle,
                  onClick: this.handleAddConditionalStyle.bind(this)
                }),
                ...(this._isEditingConditionalStyle ? [
                  PropertyPaneButton('newConditionalStyle', {
                    text: strings.ConditionalStyleNewButton,
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleNewConditionalStyle.bind(this)
                  })
                ] : []),
                PropertyPaneDropdown('conditionalStyleSelectedIndex', {
                  label: strings.ConditionalStyleExistingLabel,
                  options: conditionalEntryOptions,
                  selectedKey: this.properties.conditionalStyleSelectedIndex || ''
                }),
                PropertyPaneButton('loadConditionalStyle', {
                  text: strings.ConditionalStyleLoadButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleLoadConditionalStyle.bind(this)
                }),
                PropertyPaneButton('updateConditionalStyle', {
                  text: strings.ConditionalStyleUpdateButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  disabled: !this._isEditingConditionalStyle,
                  onClick: this.handleUpdateConditionalStyle.bind(this)
                }),
                PropertyPaneButton('removeConditionalStyle', {
                  text: strings.ConditionalStyleRemoveButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleRemoveConditionalStyle.bind(this)
                }),
                PropertyPaneButton('moveConditionalStyleUp', {
                  text: strings.ConditionalStyleMoveUpButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.moveConditionalStyle.bind(this, -1)
                }),
                PropertyPaneButton('moveConditionalStyleDown', {
                  text: strings.ConditionalStyleMoveDownButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.moveConditionalStyle.bind(this, 1)
                }),
                PropertyPaneButton('resetConditionalStyles', {
                  text: strings.ConditionalStyleResetButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleResetConditionalStyles.bind(this)
                }),
                PropertyPaneLabel('conditionalStyleDesignerMessage', {
                  text: this._conditionalStyleDesignerMessage || strings.ConditionalStyleStatusPlaceholder
                })
              ]
            },
            {
              groupName: strings.ConditionalStyleJsonGroupName,
              groupFields: [
                PropertyPaneTextField('conditionalStyleJson', {
                  label: strings.ConditionalStyleJsonLabel,
                  description: strings.ConditionalStyleJsonDescription,
                  placeholder: strings.ConditionalStyleJsonPlaceholder,
                  multiline: true,
                  resizable: true,
                  rows: 12
                }),
                PropertyPaneButton('validateConditionalStyles', {
                  text: strings.ConditionalStyleValidateButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleValidateConditionalStyles.bind(this)
                }),
                PropertyPaneLabel('conditionalStyleValidationMessage', {
                  text: this._conditionalStyleJsonValidationMessage || strings.ConditionalStyleValidationPlaceholder
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
