import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import DynamicProperty from '@microsoft/sp-component-base/lib/DynamicProperty';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneTextField,
} from '@microsoft/sp-webpart-base';
import { PropertyPaneCustomField } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField';
import { PropertyPaneCheckbox } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCheckbox/PropertyPaneCheckbox';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SharePointDynamicFormWebPartStrings';
import { FormSchema, FormMode } from '../../formEngine/core/types';
import { SharePointDynamicFormContainer } from './components/SharePointDynamicForm';
import { releaseOptionalFullWidth, updateResponsiveOptionalFullWidth } from '../shared/deterministicFullWidth';

var packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface ISharePointDynamicFormWebPartProps {
  forceFullWidth?: boolean;
  fixedWidth?: number;
  formSchemaJson: string;
  listName: string;
  mode: FormMode;
  dynamicItemId?: any;
  dynamicItemIdReference?: string;
  dynamicPreferredSourceInstanceId?: string;
  dynamicItemMode?: any;
  dynamicItemModeReference?: string;
  useDynamicItemIdAsItemId?: boolean;
  dynamicItemTargetField?: string;
  permissionBaseLookupField?: string;
  permissionScope?: 'list' | 'item';
  permissionScopeItemId?: string;
  permissionScopeLookupPick?: string;
  useDynamicValueAsFilter?: boolean;
  filterJson?: string;
  useItemId: boolean;
  itemId: number;
  itemIdQueryParam?: string;
  isInDesignerMode: boolean;
  labelPosition?: 'top' | 'left';
  showFieldDescription?: boolean;
  enableDynamicDiagnostics?: boolean;
  // Button configuration
  submitButtonLabel?: string;
  addSubmitButtonLabel?: string;
  editSubmitButtonLabel?: string;
  showCancelButton?: boolean;
  cancelButtonLabel?: string;
  cancelRedirectUrl?: string;
  submitRedirectUrl?: string;
  onSubmitMessage?: string;
  buttonTextColor?: string;
  buttonBackgroundColor?: string;
  buttonBorderColor?: string;
  buttonBorderWidth?: number | string;
  buttonFontFamily?: string;
  buttonFontSize?: number | string;
  buttonFontStyle?: string;
  buttonFontBold?: boolean;
  buttonCornerStyle?: 'square' | 'rounded';
  buttonCornerRadius?: number;
  defaultValuesJson?: string;
  advancedValidationEnabled?: boolean;
  advancedValidationJson?: string;
  validationDesignerExpression?: string;
  validationDesignerExpressionField?: string;
  validationDesignerExpressionTemplate?: string;
  validationDesignerOneClickTemplate?: string;
  validationDesignerMessage?: string;
  validationDesignerTargetField?: string;
  validationDesignerSelectedIndex?: string;
  requiredRulesJson?: string;
  requiredDesignerField?: string;
  requiredDesignerRequired?: 'required' | 'optional';
  requiredDesignerMessage?: string;
  requiredDesignerSelectedIndex?: string;
  filterDesignerField?: string;
  filterDesignerType?: 'text' | 'number' | 'lookup' | 'boolean' | 'datetime';
  filterDesignerOperator?: 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le' | 'contains' | 'startswith' | 'endswith';
  filterDesignerLogical?: 'and' | 'or';
  filterDesignerValueSource?: 'dynamic' | 'static' | 'expression' | 'fieldValue';
  filterDesignerValue?: string;
  filterDesignerSelectedIndex?: string;
  filterDesignerLookupPick?: string;
  defaultDesignerField?: string;
  defaultDesignerValueType?: 'string' | 'number' | 'boolean' | 'json' | 'expression';
  defaultDesignerValue?: string;
  defaultDesignerSelectedField?: string;
  defaultDesignerLookupPick?: string;
}

export interface IDropdownOption {
  key: string | number;
  text: string;
}

function normalizeColorValue(value: any, fallback: string): string {
  if (typeof value === 'string') {
    var trimmed = value.trim();
    return trimmed ? trimmed : fallback;
  }

  if (value && typeof value === 'object') {
    var fromStr = (value as any).str;
    if (typeof fromStr === 'string' && fromStr.trim()) {
      return fromStr;
    }

    var fromHex = (value as any).hex;
    if (typeof fromHex === 'string' && fromHex.trim()) {
      return fromHex;
    }

    var fromColor = (value as any).color;
    if (typeof fromColor === 'string' && fromColor.trim()) {
      return fromColor;
    }
  }

  return fallback;
}

function toPositiveNumber(value: any, fallback: number): number {
  var parsed = parseInt(String(value || ''), 10);
  return !isNaN(parsed) && parsed > 0 ? parsed : fallback;
}

function toPositiveItemId(value: any): number {
  if (value === undefined || value === null) {
    return 0;
  }

  if (typeof value === 'number') {
    return !isNaN(value) && value > 0 ? value : 0;
  }

  if (typeof value === 'string') {
    var trimmed = value.trim();
    if (trimmed.charAt(0) === '{' || trimmed.charAt(0) === '[') {
      try {
        return toPositiveItemId(JSON.parse(trimmed));
      } catch (_parseError) {
        // Continue with primitive parsing.
      }
    }
    var parsed = parseInt(value, 10);
    return !isNaN(parsed) && parsed > 0 ? parsed : 0;
  }

  if (typeof value === 'object') {
    if (Array.isArray(value)) {
      for (var i = 0; i < value.length; i += 1) {
        var fromArrayEntry = toPositiveItemId(value[i]);
        if (fromArrayEntry > 0) {
          return fromArrayEntry;
        }
      }
      return 0;
    }

    var idCandidate =
      (value as any).ID ||
      (value as any).Id ||
      (value as any).id ||
      (value as any).ItemId ||
      (value as any).itemId;
    if (idCandidate !== undefined && idCandidate !== null && idCandidate !== '') {
      return toPositiveItemId(idCandidate);
    }

    var resultsCandidate = (value as any).results || (value as any).value || (value as any).items;
    if (resultsCandidate !== undefined) {
      return toPositiveItemId(resultsCandidate);
    }
  }

  return 0;
}

function getSerializedDynamicReference(dynamicProperty: any): string | undefined {
  if (!dynamicProperty) {
    return undefined;
  }

  if (typeof dynamicProperty === 'string') {
    return dynamicProperty;
  }

  if (dynamicProperty.sourceId && (dynamicProperty.property || dynamicProperty.propertyId)) {
    var directProperty = dynamicProperty.property || dynamicProperty.propertyId;
    var directPath = dynamicProperty.propertyPath || '';
    return String(dynamicProperty.sourceId) + ':' + String(directProperty) + (directPath ? ':' + String(directPath) : '');
  }

  if (typeof dynamicProperty.reference === 'string' && dynamicProperty.reference) {
    return dynamicProperty.reference;
  }

  if (dynamicProperty.reference && typeof dynamicProperty.reference.reference === 'string' && dynamicProperty.reference.reference) {
    return dynamicProperty.reference.reference;
  }

  if (dynamicProperty.reference && typeof dynamicProperty.reference._reference === 'string' && dynamicProperty.reference._reference) {
    return dynamicProperty.reference._reference;
  }

  if (typeof dynamicProperty._reference === 'string' && dynamicProperty._reference) {
    return dynamicProperty._reference;
  }

  if (dynamicProperty.reference && typeof dynamicProperty.reference === 'object') {
    var sourceId = dynamicProperty.reference.sourceId || dynamicProperty.reference._sourceId;
    var propertyId = dynamicProperty.reference.property || dynamicProperty.reference._property;
    var propertyPath = dynamicProperty.reference.propertyPath || dynamicProperty.reference._propertyPath || '';
    if (sourceId && propertyId) {
      return String(sourceId) + ':' + String(propertyId) + (propertyPath ? ':' + String(propertyPath) : '');
    }
  }

  return undefined;
}

function toFormMode(value: any, fallbackMode: FormMode): FormMode {
  if (typeof value === 'string') {
    var normalized = value.trim().toLowerCase();
    if (normalized === 'new' || normalized === 'edit' || normalized === 'view') {
      return normalized as FormMode;
    }
  }

  if (value && typeof value === 'object') {
    var candidate = (value as any).mode || (value as any).Mode || (value as any).selectedMode || (value as any).value;
    if (candidate !== undefined) {
      return toFormMode(candidate, fallbackMode);
    }
  }

  return fallbackMode;
}

function escapeODataText(value: string): string {
  return String(value || '').replace(/'/g, "''");
}

function normalizeQueryParamName(value: string | undefined, fallback: string): string {
  var normalized = String(value || '').trim();
  if (!normalized || normalized.toLowerCase() === 'id') {
    return fallback;
  }

  return normalized;
}

export default class SharePointDynamicFormWebPart extends BaseClientSideWebPart<ISharePointDynamicFormWebPartProps> {
  private static readonly LIST_CONTROL_COMPONENT_ID: string = '3fbfa88a-4ac3-4b8e-b7a6-eb5256cc7001';
  private static readonly LIST_CONTROL_ALIAS: string = 'ListControlWebPart';
  private static readonly GRID_CONTROL_COMPONENT_ID: string = 'de5f92e8-570c-4a7c-ba96-7ae89a098723';
  private static readonly GRID_CONTROL_ALIAS: string = 'GridControlWebPart';
  private static readonly CALENDAR_COMPONENT_ID: string = '7e7cc010-893d-4d1f-b120-4659c2fc2806';
  private static readonly CALENDAR_ALIAS: string = 'CalendarWebPart';

  private _isDarkTheme: boolean = false;
  private _lists: IDropdownOption[] = [];
  private _listFields: IDropdownOption[] = [];
  private _dynamicTargetLookupFields: IDropdownOption[] = [];
  private _lookupPermissionFields: IDropdownOption[] = [];
  private _listControlSources: IDropdownOption[] = [];
  private _isInDesignerMode: boolean = false;
  private _runtimeDynamicItemId?: DynamicProperty<number | string>;
  private _runtimeDynamicItemRef?: string;
  private _runtimeDynamicItemValue?: any;
  private _runtimeDynamicItemMode?: DynamicProperty<string>;
  private _runtimeDynamicItemModeRef?: string;
  private _runtimeDynamicItemModeValue?: any;
  private _dynamicSourcesChangedHandler?: () => void;
  private _dynamicSourcesChangedRegistered: boolean = false;
  private _dynamicValueChangedHandler?: () => void;
  private _dynamicValueChangedSourceId?: string;
  private _dynamicValueChangedRegistered: boolean = false;
  private _dynamicDiagnostics: string[] = [];
  private _filterJsonValidationMessage: string = '';
  private _defaultValuesJsonValidationMessage: string = '';
  private _advancedValidationJsonValidationMessage: string = '';
  private _requiredRulesJsonValidationMessage: string = '';
  private _filterDesignerMessage: string = '';
  private _defaultDesignerMessage: string = '';
  private _showFilterExpressionHelp: boolean = false;
  private _showDefaultExpressionHelp: boolean = false;
  private _fieldTypeByInternalName: { [internalName: string]: string } = {};
  private _fieldChoicesByInternalName: { [internalName: string]: string[] } = {};
  private _fieldLookupListByInternalName: { [internalName: string]: string } = {};
  private _filterLookupItemOptions: IDropdownOption[] = [];
  private _defaultLookupItemOptions: IDropdownOption[] = [];
  private _filterLookupMessage: string = '';
  private _defaultLookupMessage: string = '';
  private _permissionLookupItemOptions: IDropdownOption[] = [];
  private _permissionLookupMessage: string = '';
  private _validationDesignerMessage: string = '';
  private _requiredDesignerMessage: string = '';

  private handleColorPropertyChange(propertyPath: string, oldValue: any, newValue: any): void {
    var fallback = propertyPath === 'buttonBackgroundColor' ? '#f0f0f0' : '#000000';
    var normalized = normalizeColorValue(newValue, fallback);
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, normalized);
  }

  private isDesignerAvailable(): boolean {
    return true;
  }

  public render(): void {
    var forceFullWidth = this.properties.forceFullWidth === true;
    updateResponsiveOptionalFullWidth(this.domElement, this.context.instanceId, forceFullWidth);
    this.tryRebindDynamicReferences();
    this.ensureDynamicValueChangeHandler();

    var designerAvailable = this.isDesignerAvailable();
    var configuredItemId = parseInt(String(this.properties.itemId || ''), 10);
    var safeItemId = !isNaN(configuredItemId) && configuredItemId > 0 ? configuredItemId : 0;
    var dynamicItemIdValue = 0;
    var dynamicItemRawValue: any = undefined;
    var dynamicItemProperty = this.getRuntimeDynamicItemProperty();
    if (dynamicItemProperty && dynamicItemProperty.tryGetValue) {
      try {
        dynamicItemRawValue = dynamicItemProperty.tryGetValue();
        dynamicItemIdValue = toPositiveItemId(dynamicItemRawValue);
      } catch (_error) {
        dynamicItemIdValue = 0;
      }
    }
    var useDynamicItemAsItemId = this.properties.useDynamicItemIdAsItemId !== false;
    var useDynamicValueAsFilter = useDynamicItemAsItemId ? false : true;
    var safeDynamicItemId = !isNaN(dynamicItemIdValue) && dynamicItemIdValue > 0 ? dynamicItemIdValue : 0;
    var dynamicModeProperty = this.getRuntimeDynamicItemModeProperty();
    var effectiveMode: FormMode = this.properties.mode || 'new';
    if (dynamicModeProperty && dynamicModeProperty.tryGetValue) {
      try {
        effectiveMode = toFormMode(dynamicModeProperty.tryGetValue(), effectiveMode);
      } catch (_dynamicModeError) {
        // Ignore dynamic mode resolution errors and keep configured mode.
      }
    }
    const element: React.ReactElement<any> = React.createElement(
      SharePointDynamicFormContainer,
      {
        isInDesignerMode: designerAvailable ? (this.properties.isInDesignerMode || false) : false,
        isDesignerAvailable: designerAvailable,
        onToggleDesignerMode: () => this.toggleDesignerMode(),
        formSchemaJson: this.properties.formSchemaJson || '',
        labelPosition: this.properties.labelPosition || 'top',
        listName: this.properties.listName || '',
        mode: effectiveMode,
        dynamicItemId: useDynamicItemAsItemId ? safeDynamicItemId : 0,
        linkedFieldTarget: this.properties.dynamicItemTargetField || '',
        permissionBaseLookupField: this.properties.permissionBaseLookupField || '',
        permissionScope: this.properties.permissionScope || 'list',
        permissionScopeItemId: this.properties.permissionScopeItemId || '',
        linkedFieldValue: dynamicItemRawValue !== undefined ? dynamicItemRawValue : safeDynamicItemId,
        useDynamicValueAsFilter: useDynamicValueAsFilter,
        filterJson: this.properties.filterJson || '',
        useItemId: this.properties.useItemId || false,
        itemId: safeItemId,
        itemIdQueryParam: this.properties.itemIdQueryParam || 'itemid',
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: false,
        context: this.context,
        onSaveSchema: (schema) => this.saveSchema(schema),
        isPageEditMode: this.displayMode === DisplayMode.Edit,
        showFieldDescription: this.properties.showFieldDescription,
        containerWidth: forceFullWidth ? 0 : toPositiveNumber(this.properties.fixedWidth, 800),
        // Button configuration
        submitButtonLabel: this.properties.submitButtonLabel,
        addSubmitButtonLabel: this.properties.addSubmitButtonLabel,
        editSubmitButtonLabel: this.properties.editSubmitButtonLabel,
        showCancelButton: this.properties.showCancelButton,
        cancelButtonLabel: this.properties.cancelButtonLabel,
        cancelRedirectUrl: this.properties.cancelRedirectUrl,
        submitRedirectUrl: this.properties.submitRedirectUrl,
        onSubmitMessage: this.properties.onSubmitMessage,
        buttonTextColor: normalizeColorValue(this.properties.buttonTextColor, '#000000'),
        buttonBackgroundColor: normalizeColorValue(this.properties.buttonBackgroundColor, '#f0f0f0'),
        buttonBorderColor: normalizeColorValue(this.properties.buttonBorderColor, '#f0f0f0'),
        buttonBorderWidth: toPositiveNumber(this.properties.buttonBorderWidth, 1),
        buttonFontFamily: this.properties.buttonFontFamily,
        buttonFontSize: toPositiveNumber(this.properties.buttonFontSize, 14),
        buttonFontStyle: this.properties.buttonFontStyle,
        buttonFontBold: this.properties.buttonFontBold,
        buttonCornerStyle: this.properties.buttonCornerStyle || 'square',
        buttonCornerRadius: toPositiveNumber(this.properties.buttonCornerRadius, 4),
        defaultValuesJson: this.properties.defaultValuesJson || '',
        advancedValidationEnabled: this.properties.advancedValidationEnabled === true,
        advancedValidationJson: this.properties.advancedValidationJson || '',
        dynamicDiagnostics: this.properties.enableDynamicDiagnostics === false ? [] : this._dynamicDiagnostics.slice(0),
        dynamicItemReference: this.properties.dynamicItemIdReference || '',
        dynamicModeReference: this.properties.dynamicItemModeReference || '',
        dynamicPreferredSourceInstanceId: this.properties.dynamicPreferredSourceInstanceId || '',
        enableDynamicDiagnostics: this.properties.enableDynamicDiagnostics !== false,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.properties.itemIdQueryParam = normalizeQueryParamName(this.properties.itemIdQueryParam, 'itemid');
    this.getRuntimeDynamicItemProperty();
    this.getRuntimeDynamicItemModeProperty();
    this.promoteRuntimeDynamicProperties();
    this.enforceAllowedDynamicSources();
    this.registerDynamicSourceChangeHandler();

    return this._getEnvironmentMessage().then(() => {
      return this.loadLists().then(() => {
        if (this.properties.listName) {
          return this.loadListFields(this.properties.listName).then(() => {
            this.refreshAvailableListControlSources();
          });
        }
        this.refreshAvailableListControlSources();
        return Promise.resolve();
      });
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.properties.listName && this._listFields.length === 0) {
      this.loadListFields(this.properties.listName);
    }
    this.refreshAvailableListControlSources();
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this.syncPersistedDynamicReference();
    this.syncPersistedDynamicModeReference();
    this.getRuntimeDynamicItemProperty(true);
    this.getRuntimeDynamicItemModeProperty(true);
    this.promoteRuntimeDynamicProperties();
    this.pushDynamicDiagnostic('Property pane changes applied. itemRef=' + String(this.properties.dynamicItemIdReference || '(empty)') + ', modeRef=' + String(this.properties.dynamicItemModeReference || '(empty)'));
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    var incomingReference = getSerializedDynamicReference(newValue) || '';

    if ((propertyPath === 'dynamicItemId' || propertyPath === 'dynamicItemMode') && incomingReference) {
      incomingReference = this.normalizeDynamicReference(propertyPath, incomingReference);
      if (!incomingReference) {
        this.pushDynamicDiagnostic('Rejected dynamic reference for ' + propertyPath + ' because it was not a valid ListControl binding.');
        newValue = undefined;
      }
    }

    if (propertyPath === 'dynamicPreferredSourceInstanceId') {
      this.properties.dynamicPreferredSourceInstanceId = String(newValue || '');
      this.applySelectedListControlSource(this.properties.dynamicPreferredSourceInstanceId);
      this.pushDynamicDiagnostic('Selected ListControl source: ' + String(this.properties.dynamicPreferredSourceInstanceId || '(empty)'));
    } else if (propertyPath === 'dynamicItemId') {
      this.properties.dynamicItemIdReference = incomingReference || '';
      var parsedSelected = this.parseSerializedDynamicReference(this.properties.dynamicItemIdReference || '');
      this.properties.dynamicPreferredSourceInstanceId = parsedSelected ? parsedSelected.sourceId : (this.properties.dynamicPreferredSourceInstanceId || '');
      this.applyDynamicItemReference(this.properties.dynamicItemIdReference || '');
      this.syncDynamicModeReferenceFromItemReference(this.properties.dynamicItemIdReference || '');
      this.pushDynamicDiagnostic('Selected Connect to Source reference: ' + String(this.properties.dynamicItemIdReference || '(empty)'));
    } else if (propertyPath === 'dynamicItemMode') {
      this.properties.dynamicItemModeReference = incomingReference || '';
      this.applyDynamicModeReference(this.properties.dynamicItemModeReference || '');
    } else if (propertyPath === 'useDynamicItemIdAsItemId') {
      var usesDynamicItemIdDirectly = newValue !== false;
      this.properties.useDynamicValueAsFilter = !usesDynamicItemIdDirectly;
      if (usesDynamicItemIdDirectly) {
        this.context.propertyPane.refresh();
      }
    } else if (propertyPath === 'listName') {
      this.properties.dynamicItemTargetField = '';
      this.properties.permissionBaseLookupField = '';
      this.properties.permissionScope = 'list';
      this._listFields = [];
      this._dynamicTargetLookupFields = [];
      this._lookupPermissionFields = [];
      if (newValue) {
        this.loadListFields(String(newValue));
      } else {
        this.context.propertyPane.refresh();
      }
    } else if (propertyPath === 'permissionBaseLookupField') {
      if (!newValue) {
        this.properties.permissionScope = 'list';
      }
      this.properties.permissionScopeItemId = '';
      this._permissionLookupItemOptions = [];
      this.properties.permissionScopeLookupPick = '';
      this._permissionLookupMessage = '';
      if (newValue && this.properties.permissionScope === 'item' && this.isLookupTypeField(String(newValue))) {
        this.handleLoadPermissionLookupItems();
      }
      this.context.propertyPane.refresh();
    } else if (propertyPath === 'permissionScope') {
      if (newValue !== 'item') {
        this.properties.permissionScopeItemId = '';
        this._permissionLookupItemOptions = [];
        this.properties.permissionScopeLookupPick = '';
        this._permissionLookupMessage = '';
      } else if (this.isLookupTypeField(this.properties.permissionBaseLookupField || '')) {
        this.handleLoadPermissionLookupItems();
      }
      this.context.propertyPane.refresh();
    } else if (propertyPath === 'permissionScopeLookupPick') {
      var pickedPermissionItemId = String(newValue || '');
      this.properties.permissionScopeItemId = pickedPermissionItemId;
      this._permissionLookupMessage = pickedPermissionItemId
        ? 'Using item ID ' + pickedPermissionItemId + ' as the fixed permission item. This item now determines permissions on both Add and Edit.'
        : '';
      this.context.propertyPane.refresh();
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.promoteRuntimeDynamicProperties();
    this.ensureDynamicValueChangeHandler();

    if ((propertyPath === 'dynamicItemId' || propertyPath === 'dynamicItemMode' || propertyPath === 'dynamicPreferredSourceInstanceId') && !incomingReference) {
      this.enforceAllowedDynamicSources();
      this.context.propertyPane.refresh();
      this.render();
    }

    if (propertyPath === 'buttonCornerStyle') {
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'filterDesignerField' || propertyPath === 'filterDesignerType') {
      if (propertyPath === 'filterDesignerField') {
        this.properties.filterDesignerType = this.getFilterDesignerType(String(newValue || ''));
        this.properties.filterDesignerValue = '';
      }
      this._filterLookupItemOptions = [];
      this.properties.filterDesignerLookupPick = '';
      this._filterLookupMessage = '';
      if ((this.properties.filterDesignerType || 'text') === 'lookup' && this.isLookupTypeField(this.properties.filterDesignerField || '')) {
        this.handleLoadFilterLookupItems();
      }
    }

    if (propertyPath === 'filterDesignerValueSource') {
      this.properties.filterDesignerValue = '';
      this.properties.filterDesignerLookupPick = '';
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'filterDesignerLookupPick' && (this.properties.filterDesignerValueSource || 'dynamic') === 'fieldValue') {
      this.properties.filterDesignerValue = String(newValue || '');
      this._filterLookupMessage = newValue ? 'Selected item ID ' + String(newValue) + '.' : '';
    }

    if (propertyPath === 'defaultDesignerField') {
      this._defaultLookupItemOptions = [];
      this.properties.defaultDesignerLookupPick = '';
      this._defaultLookupMessage = '';
      if (this.isLookupTypeField(String(newValue || ''))) {
        this.handleLoadDefaultLookupItems();
      }
    }

    if (
      propertyPath === 'filterDesignerValueSource'
      || propertyPath === 'filterDesignerField'
      || propertyPath === 'filterDesignerType'
      || propertyPath === 'filterDesignerOperator'
      || propertyPath === 'filterDesignerLogical'
      || propertyPath === 'filterDesignerSelectedIndex'
      || propertyPath === 'filterDesignerLookupPick'
      || propertyPath === 'validationDesignerSelectedIndex'
      || propertyPath === 'validationDesignerExpressionField'
      || propertyPath === 'validationDesignerExpressionTemplate'
      || propertyPath === 'validationDesignerOneClickTemplate'
      || propertyPath === 'validationDesignerTargetField'
      || propertyPath === 'requiredDesignerSelectedIndex'
      || propertyPath === 'requiredDesignerField'
      || propertyPath === 'requiredDesignerRequired'
      || propertyPath === 'defaultDesignerField'
      || propertyPath === 'defaultDesignerValueType'
      || propertyPath === 'defaultDesignerSelectedField'
      || propertyPath === 'defaultDesignerLookupPick'
      || propertyPath === 'permissionScopeItemId'
      || propertyPath === 'permissionScopeLookupPick'
    ) {
      this.context.propertyPane.refresh();
    }
  }

  protected onAfterDeserialize(deserializedObject: ISharePointDynamicFormWebPartProps, dataVersion: Version): ISharePointDynamicFormWebPartProps {
    var restored = super.onAfterDeserialize(deserializedObject, dataVersion) as ISharePointDynamicFormWebPartProps;
    var dynamicItemProperty = restored.dynamicItemId as any;
    var hasRuntimeDynamicProperty = !!(dynamicItemProperty && typeof dynamicItemProperty.tryGetValue === 'function');

    if (!hasRuntimeDynamicProperty && restored.dynamicItemIdReference) {
      var provider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
      if (provider) {
        var nextDynamicProperty = new DynamicProperty<number | string>(provider, this.render.bind(this));
        try {
          nextDynamicProperty.setReference(restored.dynamicItemIdReference);
        } catch (_setReferenceError) {
          // Source may not be available yet during deserialize; keep reference persisted and retry later.
        }
        restored.dynamicItemId = nextDynamicProperty as any;
      }
    }

    var dynamicModeProperty = restored.dynamicItemMode as any;
    var hasRuntimeDynamicModeProperty = !!(dynamicModeProperty && typeof dynamicModeProperty.tryGetValue === 'function');

    if (!hasRuntimeDynamicModeProperty && restored.dynamicItemModeReference) {
      var modeProvider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
      if (modeProvider) {
        var nextDynamicModeProperty = new DynamicProperty<string>(modeProvider, this.render.bind(this));
        try {
          nextDynamicModeProperty.setReference(restored.dynamicItemModeReference);
        } catch (_setModeReferenceError) {
          // Source may not be available yet during deserialize; keep reference persisted and retry later.
        }
        restored.dynamicItemMode = nextDynamicModeProperty as any;
      }
    }

    return restored;
  }

  protected onBeforeSerialize(): void {
    this.syncPersistedDynamicReference();
    this.syncPersistedDynamicModeReference();
    var dynamicProperty = this.properties.dynamicItemId as any;
    if (!dynamicProperty || typeof dynamicProperty.tryGetValue === 'function') {
      // Continue and validate mode dynamic property as well.
    } else {
      // Ensure SPFx dynamic data serializer sees a DynamicProperty instance.
      var runtimeProperty = this.getRuntimeDynamicItemProperty(true);
      if (runtimeProperty && typeof runtimeProperty.tryGetValue === 'function') {
        this.properties.dynamicItemId = runtimeProperty;
      }
    }

    var dynamicModeProperty = this.properties.dynamicItemMode as any;
    if (dynamicModeProperty && typeof dynamicModeProperty.tryGetValue !== 'function') {
      var runtimeModeProperty = this.getRuntimeDynamicItemModeProperty(true);
      if (runtimeModeProperty && typeof runtimeModeProperty.tryGetValue === 'function') {
        this.properties.dynamicItemMode = runtimeModeProperty;
      }
    }
  }

  private syncPersistedDynamicReference(): void {
    var runtimeProperty = this.getRuntimeDynamicItemProperty();
    var nextReference = getSerializedDynamicReference(runtimeProperty) || getSerializedDynamicReference(this.properties.dynamicItemId as any);
    this.properties.dynamicItemIdReference = nextReference || this.properties.dynamicItemIdReference || '';
    var parsedRef = this.parseSerializedDynamicReference(this.properties.dynamicItemIdReference);
    if (parsedRef && parsedRef.sourceId) {
      this.properties.dynamicPreferredSourceInstanceId = parsedRef.sourceId;
    }
  }

  private syncPersistedDynamicModeReference(): void {
    var runtimeProperty = this.getRuntimeDynamicItemModeProperty();
    var nextReference = getSerializedDynamicReference(runtimeProperty) || getSerializedDynamicReference(this.properties.dynamicItemMode as any);
    this.properties.dynamicItemModeReference = nextReference || this.properties.dynamicItemModeReference || '';
  }

  private enforceAllowedDynamicSources(): void {
    if (this.properties.dynamicPreferredSourceInstanceId) {
      this.applySelectedListControlSource(this.properties.dynamicPreferredSourceInstanceId);
    }

    if (this.properties.dynamicItemIdReference) {
      this.properties.dynamicItemIdReference = this.normalizeDynamicReference('dynamicItemId', this.properties.dynamicItemIdReference);
    }

    if (!this.properties.dynamicItemIdReference) {
      this.properties.dynamicItemIdReference = '';
      this.properties.dynamicItemId = undefined;
      this._runtimeDynamicItemId = undefined;
      this._runtimeDynamicItemRef = undefined;
      this._runtimeDynamicItemValue = undefined;
    } else {
      this.applyDynamicItemReference(this.properties.dynamicItemIdReference);
      this.pushDynamicDiagnostic('Enforced/retained item reference: ' + String(this.properties.dynamicItemIdReference));
    }

    if (this.properties.dynamicItemModeReference) {
      this.properties.dynamicItemModeReference = this.normalizeDynamicReference('dynamicItemMode', this.properties.dynamicItemModeReference);
    }

    if (!this.properties.dynamicItemModeReference) {
      this.properties.dynamicItemModeReference = '';
      this.properties.dynamicItemMode = undefined;
      this._runtimeDynamicItemMode = undefined;
      this._runtimeDynamicItemModeRef = undefined;
      this._runtimeDynamicItemModeValue = undefined;
    } else {
      this.applyDynamicModeReference(this.properties.dynamicItemModeReference);
    }

    if (this.properties.dynamicItemIdReference) {
      this.syncDynamicModeReferenceFromItemReference(this.properties.dynamicItemIdReference);
    }
  }

  private syncDynamicModeReferenceFromItemReference(itemReference: string): void {
    var nextModeReference = this.buildSiblingReference(itemReference, 'selectedMode');
    if (nextModeReference) {
      nextModeReference = this.normalizeDynamicReference('dynamicItemMode', nextModeReference);
    }

    if (!nextModeReference) {
      this.properties.dynamicItemModeReference = '';
      this.properties.dynamicItemMode = undefined;
      this._runtimeDynamicItemMode = undefined;
      this._runtimeDynamicItemModeRef = undefined;
      this._runtimeDynamicItemModeValue = undefined;
      return;
    }

    this.properties.dynamicItemModeReference = nextModeReference;
    this.applyDynamicModeReference(nextModeReference);
  }

  private buildSiblingReference(serializedReference: string, siblingPropertyId: string): string {
    var parsed = this.parseSerializedDynamicReference(serializedReference);
    if (!parsed || !siblingPropertyId) {
      return '';
    }

    return parsed.sourceId + ':' + siblingPropertyId;
  }

  private normalizeDynamicReference(targetProperty: string, serializedReference: string): string {
    if (!serializedReference) {
      return '';
    }

    var sourceAndProperty = this.parseSerializedDynamicReference(serializedReference);
    if (!sourceAndProperty) {
      return '';
    }

    if (targetProperty === 'dynamicItemId' && sourceAndProperty.propertyId !== 'selectedItemId') {
      return '';
    }

    if (targetProperty === 'dynamicItemMode' && sourceAndProperty.propertyId !== 'selectedMode') {
      return '';
    }

    var provider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
    if (!provider) {
      return serializedReference;
    }

    if (provider.tryGetSource) {
      var source = provider.tryGetSource(sourceAndProperty.sourceId);
      if (source && source.metadata) {
        return this.isListControlSourceMetadata(source.metadata) ? serializedReference : '';
      }
    }

    if (provider.getAvailableSources) {
      var availableSources = provider.getAvailableSources() || [];
      var listControlSources: any[] = [];
      for (var i = 0; i < availableSources.length; i += 1) {
        var availableSource = availableSources[i];
        if (availableSource && availableSource.metadata && this.isListControlSourceMetadata(availableSource.metadata)) {
          listControlSources.push(availableSource);
        }
      }

      var preferredSourceId = this.properties.dynamicPreferredSourceInstanceId || '';
      if (preferredSourceId) {
        for (var j = 0; j < listControlSources.length; j += 1) {
          if (String(listControlSources[j].id) === preferredSourceId) {
            return preferredSourceId + ':' + sourceAndProperty.propertyId;
          }
        }
      }

      if (listControlSources.length > 0 && listControlSources[0].id) {
        return String(listControlSources[0].id) + ':' + sourceAndProperty.propertyId;
      }
    }

    return serializedReference;
  }

  private isListControlSourceMetadata(metadata: any): boolean {
    if (!metadata) {
      return false;
    }

    var componentId = metadata.componentId ? String(metadata.componentId).toLowerCase() : '';
    var alias = metadata.alias ? String(metadata.alias).toLowerCase() : '';

    return componentId === SharePointDynamicFormWebPart.LIST_CONTROL_COMPONENT_ID
      || alias === SharePointDynamicFormWebPart.LIST_CONTROL_ALIAS.toLowerCase()
      || componentId === SharePointDynamicFormWebPart.GRID_CONTROL_COMPONENT_ID
      || alias === SharePointDynamicFormWebPart.GRID_CONTROL_ALIAS.toLowerCase()
      || componentId === SharePointDynamicFormWebPart.CALENDAR_COMPONENT_ID
      || alias === SharePointDynamicFormWebPart.CALENDAR_ALIAS.toLowerCase();
  }

  private buildDynamicReference(sourceId: string, propertyId: string): string {
    if (!sourceId || !propertyId) {
      return '';
    }

    return String(sourceId) + ':' + String(propertyId);
  }

  private applySelectedListControlSource(sourceId: string): void {
    if (!sourceId) {
      this.properties.dynamicItemIdReference = '';
      this.properties.dynamicItemModeReference = '';
      return;
    }

    var itemRef = this.buildDynamicReference(sourceId, 'selectedItemId');
    var modeRef = this.buildDynamicReference(sourceId, 'selectedMode');
    this.properties.dynamicItemIdReference = itemRef;
    this.properties.dynamicItemModeReference = modeRef;
    this.applyDynamicItemReference(itemRef);
    this.applyDynamicModeReference(modeRef);
  }

  private refreshAvailableListControlSources(): void {
    var provider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
    var options: IDropdownOption[] = [];

    if (provider && provider.getAvailableSources) {
      var availableSources = provider.getAvailableSources() || [];
      for (var i = 0; i < availableSources.length; i += 1) {
        var source = availableSources[i];
        if (!source || !source.metadata || !this.isListControlSourceMetadata(source.metadata)) {
          continue;
        }

        var instanceName = '';
        if (source.getPropertyValue) {
          try {
            instanceName = String(source.getPropertyValue('instanceName') || '').trim();
          } catch (_instanceNameError) {
            instanceName = '';
          }
        }

        var title = instanceName || source.metadata.title || source.metadata.alias || source.id;
        options.push({
          key: String(source.id),
          text: String(title) + ' (' + String(source.id) + ')'
        });
      }
    }

    if (options.length === 0) {
      options = [{ key: '', text: strings.PropDynamicSourceNone }];
    }

    this._listControlSources = options;
    this.context.propertyPane.refresh();
  }

  private applyDynamicItemReference(reference: string): void {
    if (!reference) {
      return;
    }

    var itemProperty = this.getRuntimeDynamicItemProperty(true);
    if (itemProperty && itemProperty.setReference) {
      try {
        itemProperty.setReference(reference);
        this.properties.dynamicItemId = itemProperty;
        this.pushDynamicDiagnostic('Applied item dynamic reference: ' + reference);
      } catch (_itemReferenceError) {
        this.pushDynamicDiagnostic('Failed to apply item dynamic reference immediately, will retry: ' + reference);
      }
    }
  }

  private applyDynamicModeReference(reference: string): void {
    if (!reference) {
      return;
    }

    var modeProperty = this.getRuntimeDynamicItemModeProperty(true);
    if (modeProperty && modeProperty.setReference) {
      try {
        modeProperty.setReference(reference);
        this.properties.dynamicItemMode = modeProperty;
      } catch (_modeReferenceError) {
        this.pushDynamicDiagnostic('Failed to apply mode dynamic reference immediately, will retry: ' + reference);
      }
    }
  }

  private promoteRuntimeDynamicProperties(): void {
    if (this.properties.dynamicItemIdReference) {
      var runtimeItemProperty = this.getRuntimeDynamicItemProperty(true);
      if (runtimeItemProperty && typeof runtimeItemProperty.tryGetValue === 'function') {
        this.properties.dynamicItemId = runtimeItemProperty;
      }
    }

    if (this.properties.dynamicItemModeReference) {
      var runtimeModeProperty = this.getRuntimeDynamicItemModeProperty(true);
      if (runtimeModeProperty && typeof runtimeModeProperty.tryGetValue === 'function') {
        this.properties.dynamicItemMode = runtimeModeProperty;
      }
    }
  }

  private parseSerializedDynamicReference(serializedReference: string): { sourceId: string; propertyId: string } | undefined {
    var parts = String(serializedReference || '').split(':');
    if (parts.length < 2) {
      return undefined;
    }

    var sourceId = String(parts[0] || '').trim();
    var propertyId = String(parts[1] || '').trim();

    if (!sourceId || !propertyId) {
      return undefined;
    }

    return {
      sourceId: sourceId,
      propertyId: propertyId
    };
  }

  private getRuntimeDynamicItemProperty(forceSync: boolean = false): any {
    var dynamicProperty = this.properties.dynamicItemId as any;
    var provider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;

    if (dynamicProperty && typeof dynamicProperty.tryGetValue === 'function') {
      this._runtimeDynamicItemId = dynamicProperty as DynamicProperty<number | string>;
      if (forceSync) {
        var persistedReference = this.properties.dynamicItemIdReference || getSerializedDynamicReference(dynamicProperty);
        if (persistedReference && dynamicProperty.setReference) {
          try {
            dynamicProperty.setReference(persistedReference);
          } catch (_itemRebindError) {
            // Ignore rebind errors and keep the existing runtime binding.
          }
        }
      }
      this._runtimeDynamicItemRef = this.properties.dynamicItemIdReference || getSerializedDynamicReference(dynamicProperty);
      this._runtimeDynamicItemValue = undefined;
      this.properties.dynamicItemId = dynamicProperty;
      return dynamicProperty;
    }

    if (!provider) {
      return dynamicProperty;
    }

    if (!this._runtimeDynamicItemId) {
      this._runtimeDynamicItemId = new DynamicProperty<number | string>(provider, this.render.bind(this));
      forceSync = true;
    }

    this.properties.dynamicItemId = this._runtimeDynamicItemId as any;

    var nextReference = getSerializedDynamicReference(dynamicProperty) || this.properties.dynamicItemIdReference;
    var nextValue = dynamicProperty && dynamicProperty.value !== undefined ? dynamicProperty.value : '';

    if (!forceSync && nextReference === this._runtimeDynamicItemRef && nextValue === this._runtimeDynamicItemValue) {
      return this._runtimeDynamicItemId;
    }

    if (nextReference) {
      try {
        this._runtimeDynamicItemId.setReference(nextReference);
      } catch (_setReferenceError) {
        this._runtimeDynamicItemId.setValue(nextValue);
      }
    } else {
      this._runtimeDynamicItemId.setValue(nextValue);
    }

    this._runtimeDynamicItemRef = nextReference;
    this._runtimeDynamicItemValue = nextValue;
    this.properties.dynamicItemId = this._runtimeDynamicItemId as any;
    return this._runtimeDynamicItemId;
  }

  private getRuntimeDynamicItemModeProperty(forceSync: boolean = false): any {
    var dynamicProperty = this.properties.dynamicItemMode as any;
    var provider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;

    if (dynamicProperty && typeof dynamicProperty.tryGetValue === 'function') {
      this._runtimeDynamicItemMode = dynamicProperty as DynamicProperty<string>;
      if (forceSync) {
        var persistedModeReference = this.properties.dynamicItemModeReference || getSerializedDynamicReference(dynamicProperty);
        if (persistedModeReference && dynamicProperty.setReference) {
          try {
            dynamicProperty.setReference(persistedModeReference);
          } catch (_modeRebindError) {
            // Ignore rebind errors and keep the existing runtime binding.
          }
        }
      }
      this._runtimeDynamicItemModeRef = this.properties.dynamicItemModeReference || getSerializedDynamicReference(dynamicProperty);
      this._runtimeDynamicItemModeValue = undefined;
      this.properties.dynamicItemMode = dynamicProperty;
      return dynamicProperty;
    }

    if (!provider) {
      return dynamicProperty;
    }

    if (!this._runtimeDynamicItemMode) {
      this._runtimeDynamicItemMode = new DynamicProperty<string>(provider, this.render.bind(this));
      forceSync = true;
    }

    this.properties.dynamicItemMode = this._runtimeDynamicItemMode as any;

    var nextReference = getSerializedDynamicReference(dynamicProperty) || this.properties.dynamicItemModeReference;
    var nextValue = dynamicProperty && dynamicProperty.value !== undefined ? dynamicProperty.value : '';

    if (!forceSync && nextReference === this._runtimeDynamicItemModeRef && nextValue === this._runtimeDynamicItemModeValue) {
      return this._runtimeDynamicItemMode;
    }

    if (nextReference) {
      try {
        this._runtimeDynamicItemMode.setReference(nextReference);
      } catch (_setReferenceError) {
        this._runtimeDynamicItemMode.setValue(nextValue);
      }
    } else {
      this._runtimeDynamicItemMode.setValue(nextValue);
    }

    this._runtimeDynamicItemModeRef = nextReference;
    this._runtimeDynamicItemModeValue = nextValue;
    this.properties.dynamicItemMode = this._runtimeDynamicItemMode as any;
    return this._runtimeDynamicItemMode;
  }

  private async loadLists(): Promise<void> {
    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      var response = await this.getWithAcceptFallback(
        webUrl + "/_api/web/lists?$select=Title,Hidden,BaseTemplate&$filter=Hidden eq false"
      );
      var data = await response.json();
      var lists = data && data.value ? data.value : [];

      if (!lists || lists.length === 0) {
        lists = data && data.d && data.d.results ? data.d.results : [];
      }

      this._lists = lists.map((list) => ({
        key: list.Title,
        text: list.Title,
      }));

      this.context.propertyPane.refresh();
    } catch (error) {
      // List loading failure is non-critical; dropdown will be empty
      console.warn('[SharePointDynamicFormWebPart] loadLists: Failed to load site lists: ', error);
    }
  }

  private async loadListFields(listName: string): Promise<void> {
    if (!listName) {
      this._listFields = [{ key: '', text: strings.PropDynamicTargetFieldNone }];
      this._dynamicTargetLookupFields = [{ key: '', text: strings.PropDynamicTargetFieldNone }];
      this._lookupPermissionFields = [{ key: '', text: strings.PropPermissionBaseLookupFieldNone }];
      this._fieldTypeByInternalName = {};
      this._fieldChoicesByInternalName = {};
      this._fieldLookupListByInternalName = {};
      this.context.propertyPane.refresh();
      return;
    }

    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      var response = await this.getWithAcceptFallback(
        webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/fields?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField,Sealed,LookupList,Choices&$filter=Hidden eq false and ((ReadOnlyField eq false and Sealed eq false) or InternalName eq 'Attachments')"
      );
      var data = await response.json();
      var fields = data && data.value ? data.value : [];

      if (!fields || fields.length === 0) {
        fields = data && data.d && data.d.results ? data.d.results : [];
      }

      this._fieldTypeByInternalName = {};
      this._fieldChoicesByInternalName = {};
      this._fieldLookupListByInternalName = {};
      fields.forEach((field: any) => {
        var fieldInternalName = String(field.InternalName || '');
        if (!fieldInternalName) { return; }
        this._fieldTypeByInternalName[fieldInternalName] = String(field.TypeAsString || '').toLowerCase();
        var rawChoices = field.Choices;
        this._fieldChoicesByInternalName[fieldInternalName] = Array.isArray(rawChoices)
          ? rawChoices.map(function(choice: any) { return String(choice); })
          : (rawChoices && Array.isArray(rawChoices.results) ? rawChoices.results.map(function(choice: any) { return String(choice); }) : []);
        var lookupListId = field.LookupList ? String(field.LookupList).replace(/[{}]/g, '') : '';
        if (lookupListId) {
          this._fieldLookupListByInternalName[fieldInternalName] = lookupListId;
        }
      });

      var options = fields.map(function(field: any) {
        var internalName = String(field.InternalName || '');
        var title = String(field.Title || internalName);
        return {
          key: internalName,
          text: title + (title !== internalName ? ' (' + internalName + ')' : '')
        };
      }).sort(function(a: IDropdownOption, b: IDropdownOption) {
        return String(a.text).localeCompare(String(b.text));
      });

      this._listFields = options.length > 0 ? options : [{ key: '', text: strings.PropDynamicTargetFieldNone }];

      var lookupOptions = fields
        .filter(function(field: any) {
          var typeAsString = String(field.TypeAsString || '').toLowerCase();
          return typeAsString === 'lookup' || typeAsString === 'lookupmulti';
        })
        .map(function(field: any) {
          var internalName = String(field.InternalName || '');
          var title = String(field.Title || internalName);
          return {
            key: internalName,
            text: title + (title !== internalName ? ' (' + internalName + ')' : '')
          };
        })
        .sort(function(a: IDropdownOption, b: IDropdownOption) {
          return String(a.text).localeCompare(String(b.text));
        });

      this._lookupPermissionFields = lookupOptions.length > 0
        ? [{ key: '', text: strings.PropPermissionBaseLookupFieldClear }].concat(lookupOptions)
        : [{ key: '', text: strings.PropPermissionBaseLookupFieldNone }];
      this._dynamicTargetLookupFields = lookupOptions.length > 0
        ? lookupOptions
        : [{ key: '', text: strings.PropDynamicTargetFieldNone }];

      if (this.properties.dynamicItemTargetField) {
        var hasSelectedDynamicTarget = lookupOptions.some((option: IDropdownOption) => {
          return String(option.key).toLowerCase() === String(this.properties.dynamicItemTargetField).toLowerCase();
        });
        if (!hasSelectedDynamicTarget) {
          this.properties.dynamicItemTargetField = '';
        }
      }

      if (this.properties.permissionBaseLookupField) {
        var hasSelectedPermissionField = lookupOptions.some((option: IDropdownOption) => {
          return String(option.key) === String(this.properties.permissionBaseLookupField);
        });

        if (!hasSelectedPermissionField) {
          this.properties.permissionBaseLookupField = '';
          this.properties.permissionScope = 'list';
        }
      }

      this.context.propertyPane.refresh();
    } catch (error) {
      this._listFields = [{ key: '', text: strings.PropDynamicTargetFieldNone }];
      this._dynamicTargetLookupFields = [{ key: '', text: strings.PropDynamicTargetFieldNone }];
      this._lookupPermissionFields = [{ key: '', text: strings.PropPermissionBaseLookupFieldNone }];
      this._fieldTypeByInternalName = {};
      this._fieldChoicesByInternalName = {};
      this._fieldLookupListByInternalName = {};
      this.context.propertyPane.refresh();
      console.warn('[SharePointDynamicFormWebPart] loadListFields: Failed to load fields for list "' + listName + '": ', error);
    }
  }

  private isLookupTypeField(internalName: string): boolean {
    var typeName = this._fieldTypeByInternalName[internalName || ''];
    return typeName === 'lookup' || typeName === 'lookupmulti';
  }

  private supportsFilterFieldValue(internalName: string): boolean {
    var typeName = String(this._fieldTypeByInternalName[internalName || ''] || '').toLowerCase();
    return typeName === 'choice' || typeName === 'multichoice' || typeName === 'lookup' || typeName === 'lookupmulti' || typeName === 'user' || typeName === 'usermulti';
  }

  private isMultiFilterField(internalName: string): boolean {
    var typeName = String(this._fieldTypeByInternalName[internalName || ''] || '').toLowerCase();
    return typeName === 'multichoice' || typeName === 'lookupmulti' || typeName === 'usermulti';
  }

  private getFilterDesignerType(internalName: string): 'text' | 'number' | 'lookup' | 'boolean' | 'datetime' {
    var typeName = String(this._fieldTypeByInternalName[internalName || ''] || '').toLowerCase();
    if (typeName === 'lookup' || typeName === 'lookupmulti' || typeName === 'user' || typeName === 'usermulti') { return 'lookup'; }
    if (typeName === 'number' || typeName === 'currency' || typeName === 'integer' || typeName === 'counter') { return 'number'; }
    if (typeName === 'boolean') { return 'boolean'; }
    if (typeName === 'datetime') { return 'datetime'; }
    return 'text';
  }

  private getSelectedFilterFieldValues(): string[] {
    var rawValue = String(this.properties.filterDesignerValue || '').trim();
    if (!rawValue) { return []; }
    try {
      var parsed = JSON.parse(rawValue);
      return Array.isArray(parsed) ? parsed.map(function(value: any) { return String(value); }) : [String(parsed)];
    } catch (_parseError) {
      return rawValue.split(';').map(function(value: string) { return value.trim(); }).filter(function(value: string) { return !!value; });
    }
  }

  private createFilterMultiValuePicker(options: IDropdownOption[]): any {
    return PropertyPaneCustomField({
      key: 'filterDesignerFieldValueMultiPicker',
      onRender: (domElement: HTMLElement): void => {
        while (domElement.firstChild) { domElement.removeChild(domElement.firstChild); }
        var title = document.createElement('div');
        title.textContent = 'Select one or more field values';
        title.style.fontWeight = '600';
        title.style.marginBottom = '8px';
        domElement.appendChild(title);
        var selected = this.getSelectedFilterFieldValues();
        options.forEach((option: IDropdownOption) => {
          var value = String(option.key);
          var label = document.createElement('label');
          label.style.display = 'block';
          label.style.marginBottom = '6px';
          var checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.checked = selected.indexOf(value) >= 0;
          checkbox.style.marginRight = '6px';
          checkbox.addEventListener('change', () => {
            var nextValues = this.getSelectedFilterFieldValues();
            var existingIndex = nextValues.indexOf(value);
            if (checkbox.checked && existingIndex < 0) { nextValues.push(value); }
            if (!checkbox.checked && existingIndex >= 0) { nextValues.splice(existingIndex, 1); }
            this.properties.filterDesignerValue = JSON.stringify(nextValues);
          });
          label.appendChild(checkbox);
          label.appendChild(document.createTextNode(option.text));
          domElement.appendChild(label);
        });
      }
    });
  }

  private async loadLookupListItems(listId: string): Promise<IDropdownOption[]> {
    try {
      var webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      var url = webUrl + "/_api/web/lists(guid'" + listId + "')/items?$select=Id,Title&$top=500&$orderby=Title";
      var response = await this.getWithAcceptFallback(url);
      var data = await response.json();
      var items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      return items.map(function(item: any) {
        var title = item.Title ? String(item.Title) : '(no title)';
        return { key: String(item.Id), text: title + ' (ID: ' + item.Id + ')' };
      });
    } catch (error) {
      console.warn('[SharePointDynamicFormWebPart] loadLookupListItems: Failed to load items for list "' + listId + '": ', error);
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
      ? 'Loaded ' + this._filterLookupItemOptions.length + ' item(s). Pick one and click Use selected ID.'
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
    this._filterLookupMessage = 'Applied item ID ' + picked + ' to the filter value.';
    this.context.propertyPane.refresh();
  }

  private async handleLoadDefaultLookupItems(): Promise<void> {
    var fieldName = String(this.properties.defaultDesignerField || '');
    var listId = this._fieldLookupListByInternalName[fieldName];
    if (!listId) {
      this._defaultLookupMessage = 'Selected field is not a lookup column, or its target list could not be resolved.';
      this._defaultLookupItemOptions = [];
      this.context.propertyPane.refresh();
      return;
    }

    this._defaultLookupItemOptions = await this.loadLookupListItems(listId);
    this._defaultLookupMessage = this._defaultLookupItemOptions.length > 0
      ? 'Loaded ' + this._defaultLookupItemOptions.length + ' item(s). Pick one and click Use selected ID.'
      : 'The target list has no items.';
    this.context.propertyPane.refresh();
  }

  private handleApplyDefaultLookupPick(): void {
    var picked = String(this.properties.defaultDesignerLookupPick || '');
    if (!picked) {
      this._defaultLookupMessage = 'Select an item from the dropdown before applying.';
      this.context.propertyPane.refresh();
      return;
    }

    this.properties.defaultDesignerValue = picked;
    this.properties.defaultDesignerValueType = 'string';
    this._defaultLookupMessage = 'Applied item ID ' + picked + ' to the default value.';
    this.context.propertyPane.refresh();
  }

  private async handleLoadPermissionLookupItems(): Promise<void> {
    var fieldName = String(this.properties.permissionBaseLookupField || '');
    var listId = this._fieldLookupListByInternalName[fieldName];
    if (!listId) {
      this._permissionLookupMessage = 'Select a lookup column before loading its items.';
      this._permissionLookupItemOptions = [];
      this.context.propertyPane.refresh();
      return;
    }

    this._permissionLookupItemOptions = await this.loadLookupListItems(listId);
    this._permissionLookupMessage = this._permissionLookupItemOptions.length > 0
      ? 'Loaded ' + this._permissionLookupItemOptions.length + ' item(s). Pick one and click Use selected ID.'
      : 'The target list has no items.';
    this.context.propertyPane.refresh();
  }

  private async getWithAcceptFallback(url: string): Promise<any> {
    var response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=verbose'
        }
      });
    }

    if (!response.ok) {
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=minimalmetadata'
        }
      });
    }

    if (!response.ok) {
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      });
    }

    return response;
  }

  private _getEnvironmentMessage(): Promise<string> {
    return Promise.resolve(strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: any): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    this.unregisterDynamicValueChangeHandler();
    this.unregisterDynamicSourceChangeHandler();
    releaseOptionalFullWidth(this.domElement.ownerDocument, this.context.instanceId);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private getActiveDynamicSourceId(): string {
    var parsed = this.parseSerializedDynamicReference(this.properties.dynamicItemIdReference || '');
    if (parsed && parsed.sourceId) {
      return String(parsed.sourceId);
    }

    if (this.properties.dynamicPreferredSourceInstanceId) {
      return String(this.properties.dynamicPreferredSourceInstanceId);
    }

    return '';
  }

  private ensureDynamicValueChangeHandler(): void {
    var provider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
    if (!provider || !provider.registerSourceChanged) {
      return;
    }

    var sourceId = this.getActiveDynamicSourceId();
    if (!sourceId) {
      this.unregisterDynamicValueChangeHandler();
      return;
    }

    if (this._dynamicValueChangedRegistered && this._dynamicValueChangedSourceId === sourceId) {
      return;
    }

    this.unregisterDynamicValueChangeHandler();

    if (!this._dynamicValueChangedHandler) {
      this._dynamicValueChangedHandler = () => {
        this.pushDynamicDiagnostic('Dynamic source value changed for source ' + String(this._dynamicValueChangedSourceId || '(unknown)') + '. Re-rendering.');
        this.render();
      };
    }

    provider.registerSourceChanged(sourceId, this._dynamicValueChangedHandler);
    this._dynamicValueChangedSourceId = sourceId;
    this._dynamicValueChangedRegistered = true;
    this.pushDynamicDiagnostic('Registered source-changed handler for source ' + sourceId);
  }

  private unregisterDynamicValueChangeHandler(): void {
    if (!this._dynamicValueChangedRegistered || !this._dynamicValueChangedHandler || !this._dynamicValueChangedSourceId) {
      this._dynamicValueChangedRegistered = false;
      this._dynamicValueChangedSourceId = undefined;
      return;
    }

    var provider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
    if (provider && provider.unregisterSourceChanged) {
      provider.unregisterSourceChanged(this._dynamicValueChangedSourceId, this._dynamicValueChangedHandler);
    }

    this._dynamicValueChangedRegistered = false;
    this._dynamicValueChangedSourceId = undefined;
  }

  private registerDynamicSourceChangeHandler(): void {
    var provider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
    if (!provider || !provider.registerAvailableSourcesChanged || this._dynamicSourcesChangedRegistered) {
      return;
    }

    if (!this._dynamicSourcesChangedHandler) {
      this._dynamicSourcesChangedHandler = () => {
        this.pushDynamicDiagnostic('Dynamic available sources changed event received. Retrying binding.');
        this.refreshAvailableListControlSources();
        this.tryRebindDynamicReferences();
        this.promoteRuntimeDynamicProperties();
        this.render();
      };
    }

    provider.registerAvailableSourcesChanged(this._dynamicSourcesChangedHandler);
    this._dynamicSourcesChangedRegistered = true;
  }

  private unregisterDynamicSourceChangeHandler(): void {
    var provider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
    if (!provider || !provider.unregisterAvailableSourcesChanged || !this._dynamicSourcesChangedRegistered || !this._dynamicSourcesChangedHandler) {
      return;
    }

    provider.unregisterAvailableSourcesChanged(this._dynamicSourcesChangedHandler);
    this._dynamicSourcesChangedRegistered = false;
  }

  private tryRebindDynamicReferences(): void {
    if (this.properties.dynamicItemIdReference) {
      var normalizedItemRef = this.normalizeDynamicReference('dynamicItemId', this.properties.dynamicItemIdReference);
      if (normalizedItemRef) {
        if (normalizedItemRef !== this.properties.dynamicItemIdReference) {
          this.pushDynamicDiagnostic('Normalized item reference from ' + this.properties.dynamicItemIdReference + ' to ' + normalizedItemRef);
        }
        this.properties.dynamicItemIdReference = normalizedItemRef;
        this.applyDynamicItemReference(normalizedItemRef);
      }
    }

    if (this.properties.dynamicItemIdReference) {
      this.syncDynamicModeReferenceFromItemReference(this.properties.dynamicItemIdReference);
    }

    if (this.properties.dynamicItemModeReference) {
      var normalizedModeRef = this.normalizeDynamicReference('dynamicItemMode', this.properties.dynamicItemModeReference);
      if (normalizedModeRef) {
        if (normalizedModeRef !== this.properties.dynamicItemModeReference) {
          this.pushDynamicDiagnostic('Normalized mode reference from ' + this.properties.dynamicItemModeReference + ' to ' + normalizedModeRef);
        }
        this.properties.dynamicItemModeReference = normalizedModeRef;
        this.applyDynamicModeReference(normalizedModeRef);
      }
    }
  }

  private pushDynamicDiagnostic(message: string): void {
    if (this.properties.enableDynamicDiagnostics === false) {
      return;
    }

    var entry = '[' + new Date().toLocaleTimeString() + '] ' + message;
    console.log('[SharePointDynamicFormWebPart] ' + entry);
    
    if (this._dynamicDiagnostics.length > 0 && this._dynamicDiagnostics[0] === entry) {
      return;
    }

    this._dynamicDiagnostics.unshift(entry);
    if (this._dynamicDiagnostics.length > 12) {
      this._dynamicDiagnostics = this._dynamicDiagnostics.slice(0, 12);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private toggleDesignerMode(): void {
    if (!this.isDesignerAvailable()) {
      this._isInDesignerMode = false;
      this.properties.isInDesignerMode = false;
      this.render();
      return;
    }

    this._isInDesignerMode = !this._isInDesignerMode;
    this.properties.isInDesignerMode = this._isInDesignerMode;
    this.render();
  }

  private saveSchema(schema: FormSchema): void {
    this.properties.formSchemaJson = JSON.stringify(schema);
    this.context.propertyPane.refresh();
  }

  private getFontFamilyOptions(): Array<{ key: string; text: string }> {
    return [
      { key: 'inherit', text: 'Inherit' },
      { key: 'Segoe UI', text: 'Segoe UI' },
      { key: 'Arial, Helvetica, sans-serif', text: 'Arial' },
      { key: 'Verdana, Geneva, sans-serif', text: 'Verdana' },
      { key: 'Tahoma, Geneva, sans-serif', text: 'Tahoma' },
      { key: '"Trebuchet MS", Helvetica, sans-serif', text: 'Trebuchet MS' },
      { key: '"Times New Roman", Times, serif', text: 'Times New Roman' },
      { key: 'Georgia, serif', text: 'Georgia' },
      { key: '"Courier New", Courier, monospace', text: 'Courier New' }
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
      var errorMessage = error && error.message ? error.message : strings.CommonUnknownError;
      return {
        valid: false,
        message: strings.JsonValidationFailedPrefix + ': ' + errorMessage
      };
    }
  }

  private handleValidateFilterJson(): void {
    var result = this.validateJsonText(this.properties.filterJson, 'array');
    this._filterJsonValidationMessage = result.message;
    if (result.valid) {
      this.pushDynamicDiagnostic('Record filter JSON validation passed.');
    } else {
      this.pushDynamicDiagnostic('Record filter JSON validation failed: ' + result.message);
    }
    this.context.propertyPane.refresh();
  }

  private handleValidateDefaultValuesJson(): void {
    var result = this.validateJsonText(this.properties.defaultValuesJson, 'object');
    this._defaultValuesJsonValidationMessage = result.message;
    if (result.valid) {
      this.pushDynamicDiagnostic('Default values JSON validation passed.');
    } else {
      this.pushDynamicDiagnostic('Default values JSON validation failed: ' + result.message);
    }
    this.context.propertyPane.refresh();
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

  private parseDefaultValuesObject(raw: string | undefined): { [key: string]: any } {
    var source = String(raw || '').trim();
    if (!source) {
      return {};
    }

    var parsed = JSON.parse(source);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error(strings.JsonValidationExpectedObject);
    }

    return parsed;
  }

  private isSupportedSafeExpression(value: string): boolean {
    var normalized = String(value || '').trim().toLowerCase();
    return normalized === 'today'
      || normalized === 'now'
      || normalized === 'me'
      || normalized === 'me.email'
      || normalized === 'me.login'
      || normalized === 'me.id'
      || /^date\(\s*[+-]?\d+\s*\)$/.test(normalized);
  }

  private createExpressionHelpField(key: string, isDefault: boolean): any {
    var isExpanded = isDefault ? this._showDefaultExpressionHelp : this._showFilterExpressionHelp;
    return PropertyPaneCustomField({
      key: key + '-' + String(isExpanded),
      onRender: (domElement: HTMLElement): void => {
        domElement.innerHTML = '';
        var button = document.createElement('button');
        button.type = 'button';
        button.title = isExpanded ? strings.PropExpressionHelpHide : strings.PropExpressionHelpShow;
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
          if (isDefault) {
            this._showDefaultExpressionHelp = !this._showDefaultExpressionHelp;
          } else {
            this._showFilterExpressionHelp = !this._showFilterExpressionHelp;
          }
          this.context.propertyPane.refresh();
        };
        domElement.appendChild(button);
        if (isExpanded) {
          var help = document.createElement('div');
          help.textContent = strings.PropExpressionHelp;
          help.style.fontSize = '12px';
          help.style.lineHeight = '18px';
          help.style.marginTop = '4px';
          domElement.appendChild(help);
        }
      }
    });
  }

  private parseDesignerValue(raw: string, valueType: 'string' | 'number' | 'boolean' | 'json' | 'expression'): any {
    var text = String(raw || '').trim();
    if (valueType === 'expression') {
      if (!this.isSupportedSafeExpression(text)) {
        throw new Error(strings.PropExpressionInvalid);
      }
      return { $expression: text };
    }
    if (valueType === 'number') {
      var parsed = Number(text);
      if (isNaN(parsed)) {
        throw new Error('Default value must be a valid number.');
      }
      return parsed;
    }
    if (valueType === 'boolean') {
      var normalized = text.toLowerCase();
      if (normalized === 'true' || normalized === '1' || normalized === 'yes') {
        return true;
      }
      if (normalized === 'false' || normalized === '0' || normalized === 'no') {
        return false;
      }
      throw new Error('Default value must be true or false for boolean type.');
    }
    if (valueType === 'json') {
      if (!text) {
        return {};
      }
      return JSON.parse(text);
    }

    return raw || '';
  }

  private inferDesignerValueType(value: any): 'string' | 'number' | 'boolean' | 'json' | 'expression' {
    if (value && typeof value === 'object' && !Array.isArray(value) && typeof value.$expression === 'string') {
      return 'expression';
    }
    if (typeof value === 'number') {
      return 'number';
    }
    if (typeof value === 'boolean') {
      return 'boolean';
    }
    if (value && typeof value === 'object') {
      return 'json';
    }
    return 'string';
  }

  private stringifyDesignerValue(value: any, valueType: 'string' | 'number' | 'boolean' | 'json' | 'expression'): string {
    if (valueType === 'expression') {
      return value && value.$expression ? String(value.$expression) : '';
    }
    if (valueType === 'json') {
      return JSON.stringify(value === undefined ? {} : value, null, 2);
    }
    return value === undefined || value === null ? '' : String(value);
  }

  private handleAddFilterDesignerCondition(): void {
    try {
      var field = String(this.properties.filterDesignerField || '').trim();
      if (!field) {
        this._filterDesignerMessage = 'Select a field before adding a filter condition.';
        this.context.propertyPane.refresh();
        return;
      }

      var condType = (this.properties.filterDesignerType || 'text') as 'text' | 'number' | 'lookup' | 'boolean' | 'datetime';
      var condOperator = (this.properties.filterDesignerOperator || 'eq') as 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le' | 'contains' | 'startswith' | 'endswith';
      var condLogical = (this.properties.filterDesignerLogical || 'and') as 'and' | 'or';
      var valueSource = (this.properties.filterDesignerValueSource || 'dynamic') as 'dynamic' | 'static' | 'expression' | 'fieldValue';
      var conditionValue: any;

      if (valueSource === 'dynamic') {
        conditionValue = 'dynamic';
      } else {
        var rawValue = String(this.properties.filterDesignerValue || '').trim();
        if (!rawValue) {
          this._filterDesignerMessage = valueSource === 'expression' ? strings.PropExpressionInvalid : 'Enter a value for static filter condition.';
          this.context.propertyPane.refresh();
          return;
        }

        if (valueSource === 'expression') {
          if (!this.isSupportedSafeExpression(rawValue)) {
            this._filterDesignerMessage = strings.PropExpressionInvalid;
            this.context.propertyPane.refresh();
            return;
          }
          conditionValue = rawValue;
        } else if (valueSource === 'fieldValue' && this.isMultiFilterField(field)) {
          conditionValue = this.getSelectedFilterFieldValues();
          if (conditionValue.length === 0) { this._filterDesignerMessage = 'Select at least one field value.'; this.context.propertyPane.refresh(); return; }
        } else if (condType === 'number' || condType === 'lookup') {
          var parsedNumber = Number(rawValue);
          conditionValue = isNaN(parsedNumber) ? rawValue : parsedNumber;
        } else if (condType === 'boolean') {
          var boolText = rawValue.toLowerCase();
          conditionValue = boolText === 'true' || boolText === '1' || boolText === 'yes';
        } else {
          conditionValue = rawValue;
        }
      }

      var existing = this.parseFilterJsonArray(this.properties.filterJson);
      var newCondition: any = {
        field: field,
        type: condType,
        operator: condOperator,
        logical: condLogical,
        value: conditionValue
      };
      if (valueSource === 'expression' || valueSource === 'fieldValue') {
        newCondition.valueType = valueSource;
      }
      existing.push(newCondition);

      this.properties.filterJson = JSON.stringify(existing, null, 2);
      this._filterDesignerMessage = 'Condition added to Record filter JSON.';
      this._filterJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && error.message ? error.message : 'Failed to add filter condition.';
    }

    this.context.propertyPane.refresh();
    this.render();
  }

  private handleLoadSelectedFilterCondition(): void {
    try {
      var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
      var existing = this.parseFilterJsonArray(this.properties.filterJson);
      if (isNaN(index) || index < 0 || index >= existing.length) {
        this._filterDesignerMessage = 'Select a valid condition to load.';
        this.context.propertyPane.refresh();
        return;
      }

      var selected = existing[index] || {};
      var selectedValue = selected.value;
      var storedValueType = String(selected.valueType || '').toLowerCase();
      var valueSource: 'dynamic' | 'static' | 'expression' | 'fieldValue' = 'static';
      if (storedValueType === 'expression') { valueSource = 'expression'; }
      else if (storedValueType === 'fieldvalue') { valueSource = 'fieldValue'; }
      else if (selectedValue === 'dynamic') { valueSource = 'dynamic'; }

      this.properties.filterDesignerField = selected.field ? String(selected.field) : '';
      this.properties.filterDesignerType = selected.type || 'text';
      this.properties.filterDesignerOperator = selected.operator || 'eq';
      this.properties.filterDesignerLogical = selected.logical === 'or' ? 'or' : 'and';
      this.properties.filterDesignerValueSource = valueSource;
      this.properties.filterDesignerValue = valueSource !== 'dynamic' ? (Array.isArray(selectedValue) ? JSON.stringify(selectedValue) : String(selectedValue === undefined || selectedValue === null ? '' : selectedValue)) : '';
      this._filterDesignerMessage = 'Condition loaded into designer fields.';
    } catch (error) {
      this._filterDesignerMessage = error && error.message ? error.message : 'Failed to load condition.';
    }

    this.context.propertyPane.refresh();
  }

  private handleUpdateSelectedFilterCondition(): void {
    try {
      var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
      var existing = this.parseFilterJsonArray(this.properties.filterJson);
      if (isNaN(index) || index < 0 || index >= existing.length) {
        this._filterDesignerMessage = 'Select a valid condition to update.';
        this.context.propertyPane.refresh();
        return;
      }

      var field = String(this.properties.filterDesignerField || '').trim();
      if (!field) {
        this._filterDesignerMessage = 'Select a field before updating the condition.';
        this.context.propertyPane.refresh();
        return;
      }

      var condType = (this.properties.filterDesignerType || 'text') as 'text' | 'number' | 'lookup' | 'boolean' | 'datetime';
      var condOperator = (this.properties.filterDesignerOperator || 'eq') as 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le' | 'contains' | 'startswith' | 'endswith';
      var condLogical = (this.properties.filterDesignerLogical || 'and') as 'and' | 'or';
      var valueSource = (this.properties.filterDesignerValueSource || 'dynamic') as 'dynamic' | 'static' | 'expression' | 'fieldValue';
      var conditionValue: any = 'dynamic';

      if (valueSource !== 'dynamic') {
        var rawValue = String(this.properties.filterDesignerValue || '').trim();
        if (!rawValue) {
          this._filterDesignerMessage = valueSource === 'expression' ? strings.PropExpressionInvalid : 'Enter a value for static filter condition.';
          this.context.propertyPane.refresh();
          return;
        }

        if (valueSource === 'expression') {
          if (!this.isSupportedSafeExpression(rawValue)) {
            this._filterDesignerMessage = strings.PropExpressionInvalid;
            this.context.propertyPane.refresh();
            return;
          }
          conditionValue = rawValue;
        } else if (valueSource === 'fieldValue' && this.isMultiFilterField(field)) {
          conditionValue = this.getSelectedFilterFieldValues();
          if (conditionValue.length === 0) { this._filterDesignerMessage = 'Select at least one field value.'; this.context.propertyPane.refresh(); return; }
        } else if (condType === 'number' || condType === 'lookup') {
          var parsedNumber = Number(rawValue);
          conditionValue = isNaN(parsedNumber) ? rawValue : parsedNumber;
        } else if (condType === 'boolean') {
          var boolText = rawValue.toLowerCase();
          conditionValue = boolText === 'true' || boolText === '1' || boolText === 'yes';
        } else {
          conditionValue = rawValue;
        }
      }

      var updatedCondition: any = {
        field: field,
        type: condType,
        operator: condOperator,
        logical: condLogical,
        value: conditionValue
      };
      if (valueSource === 'expression' || valueSource === 'fieldValue') {
        updatedCondition.valueType = valueSource;
      }
      existing[index] = updatedCondition;

      this.properties.filterJson = JSON.stringify(existing, null, 2);
      this._filterDesignerMessage = 'Selected condition updated.';
      this._filterJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && error.message ? error.message : 'Failed to update condition.';
    }

    this.context.propertyPane.refresh();
    this.render();
  }

  private handleRemoveSelectedFilterCondition(): void {
    try {
      var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
      var existing = this.parseFilterJsonArray(this.properties.filterJson);
      if (isNaN(index) || index < 0 || index >= existing.length) {
        this._filterDesignerMessage = 'Select a valid condition to remove.';
        this.context.propertyPane.refresh();
        return;
      }

      existing.splice(index, 1);
      this.properties.filterJson = JSON.stringify(existing, null, 2);
      this.properties.filterDesignerSelectedIndex = '';
      this._filterDesignerMessage = 'Selected condition removed.';
      this._filterJsonValidationMessage = existing.length === 0 ? strings.JsonValidationEmpty : strings.JsonValidationSuccess;
    } catch (error) {
      this._filterDesignerMessage = error && error.message ? error.message : 'Failed to remove condition.';
    }

    this.context.propertyPane.refresh();
    this.render();
  }

  private handleResetFilterJson(): void {
    this.properties.filterJson = '';
    this._filterJsonValidationMessage = strings.JsonValidationEmpty;
    this._filterDesignerMessage = 'Record filter JSON cleared.';
    this.context.propertyPane.refresh();
    this.render();
  }

  private handleAddDefaultDesignerEntry(): void {
    try {
      var field = String(this.properties.defaultDesignerField || '').trim();
      if (!field) {
        this._defaultDesignerMessage = 'Select a field before adding a default value.';
        this.context.propertyPane.refresh();
        return;
      }

      var valueType = (this.properties.defaultDesignerValueType || 'string') as 'string' | 'number' | 'boolean' | 'json' | 'expression';
      var rawValue = String(this.properties.defaultDesignerValue || '');
      var parsedValue = this.parseDesignerValue(rawValue, valueType);

      var defaults = this.parseDefaultValuesObject(this.properties.defaultValuesJson);
      defaults[field] = parsedValue;

      this.properties.defaultValuesJson = JSON.stringify(defaults, null, 2);
      this._defaultDesignerMessage = 'Default value added/updated.';
      this._defaultValuesJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._defaultDesignerMessage = error && error.message ? error.message : 'Failed to add default value.';
    }

    this.context.propertyPane.refresh();
  }

  private handleLoadSelectedDefaultEntry(): void {
    try {
      var fieldKey = String(this.properties.defaultDesignerSelectedField || '').trim();
      var defaults = this.parseDefaultValuesObject(this.properties.defaultValuesJson);
      if (!fieldKey || !Object.prototype.hasOwnProperty.call(defaults, fieldKey)) {
        this._defaultDesignerMessage = 'Select a valid default entry to load.';
        this.context.propertyPane.refresh();
        return;
      }

      var currentValue = defaults[fieldKey];
      var valueType = this.inferDesignerValueType(currentValue);
      this.properties.defaultDesignerField = fieldKey;
      this.properties.defaultDesignerValueType = valueType;
      this.properties.defaultDesignerValue = this.stringifyDesignerValue(currentValue, valueType);
      this._defaultDesignerMessage = 'Default entry loaded into designer fields.';
    } catch (error) {
      this._defaultDesignerMessage = error && error.message ? error.message : 'Failed to load default entry.';
    }

    this.context.propertyPane.refresh();
  }

  private handleRemoveSelectedDefaultEntry(): void {
    try {
      var fieldKey = String(this.properties.defaultDesignerSelectedField || '').trim();
      var defaults = this.parseDefaultValuesObject(this.properties.defaultValuesJson);
      if (!fieldKey || !Object.prototype.hasOwnProperty.call(defaults, fieldKey)) {
        this._defaultDesignerMessage = 'Select a valid default entry to remove.';
        this.context.propertyPane.refresh();
        return;
      }

      delete defaults[fieldKey];
      this.properties.defaultValuesJson = JSON.stringify(defaults, null, 2);
      this.properties.defaultDesignerSelectedField = '';
      this._defaultDesignerMessage = 'Selected default entry removed.';
      this._defaultValuesJsonValidationMessage = Object.keys(defaults).length === 0 ? strings.JsonValidationEmpty : strings.JsonValidationSuccess;
    } catch (error) {
      this._defaultDesignerMessage = error && error.message ? error.message : 'Failed to remove default entry.';
    }

    this.context.propertyPane.refresh();
  }

  private handleResetDefaultValuesJson(): void {
    this.properties.defaultValuesJson = '';
    this._defaultValuesJsonValidationMessage = strings.JsonValidationEmpty;
    this._defaultDesignerMessage = 'Default values JSON cleared.';
    this.context.propertyPane.refresh();
  }

  private parseAdvancedValidationRulesArray(raw: string | undefined): any[] {
    var source = String(raw || '').trim();
    if (!source) {
      return [];
    }

    var parsed = JSON.parse(source);
    if (!Array.isArray(parsed)) {
      throw new Error(strings.JsonValidationExpectedArray);
    }

    for (var i = 0; i < parsed.length; i += 1) {
      var rule = parsed[i];
      if (!rule || typeof rule !== 'object') {
        throw new Error('Validation rule at index ' + i + ' must be an object.');
      }

      var expression = rule.expression === undefined || rule.expression === null ? '' : String(rule.expression).trim();
      var message = rule.message === undefined || rule.message === null ? '' : String(rule.message).trim();
      if (!expression) {
        throw new Error('Validation rule at index ' + i + ' is missing expression.');
      }
      if (!message) {
        throw new Error('Validation rule at index ' + i + ' is missing message.');
      }
    }

    return parsed;
  }

  private handleValidateAdvancedValidationJson(): void {
    try {
      this.parseAdvancedValidationRulesArray(this.properties.advancedValidationJson);
      this._advancedValidationJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._advancedValidationJsonValidationMessage = error && error.message ? error.message : 'Validation rules JSON is invalid.';
    }
    this.context.propertyPane.refresh();
  }

  private appendToValidationExpression(fragment: string): void {
    var nextFragment = String(fragment || '').trim();
    if (!nextFragment) {
      return;
    }

    var currentExpression = String(this.properties.validationDesignerExpression || '').trim();
    this.properties.validationDesignerExpression = currentExpression ? (currentExpression + ' ' + nextFragment) : nextFragment;
    this._validationDesignerMessage = 'Expression updated.';
    this.context.propertyPane.refresh();
  }

  private handleInsertSelectedValidationExpressionField(): void {
    var selectedField = String(this.properties.validationDesignerExpressionField || '').trim();
    if (!selectedField) {
      this._validationDesignerMessage = 'Select a field to insert into the expression.';
      this.context.propertyPane.refresh();
      return;
    }

    var escapedField = selectedField.replace(/"/g, '\\"');
    this.appendToValidationExpression('field("' + escapedField + '")');
  }

  private handleInsertValidationExpressionTemplate(): void {
    var templateKey = String(this.properties.validationDesignerExpressionTemplate || '').trim();
    var templateExpression = '';

    if (templateKey === 'today') {
      templateExpression = 'today()';
    } else if (templateKey === 'now') {
      templateExpression = 'now()';
    } else if (templateKey === 'daysFromToday') {
      templateExpression = 'daysFromToday(7)';
    } else if (templateKey === 'isEmpty') {
      templateExpression = 'isEmpty(field("InternalName"))';
    } else if (templateKey === 'hasValue') {
      templateExpression = 'hasValue(field("InternalName"))';
    } else if (templateKey === 'between') {
      templateExpression = 'between(field("InternalName"), 1, 10)';
    } else if (templateKey === 'eq') {
      templateExpression = '==';
    } else if (templateKey === 'ne') {
      templateExpression = '!=';
    } else if (templateKey === 'gt') {
      templateExpression = '>';
    } else if (templateKey === 'ge') {
      templateExpression = '>=';
    } else if (templateKey === 'lt') {
      templateExpression = '<';
    } else if (templateKey === 'le') {
      templateExpression = '<=';
    } else if (templateKey === 'openParen') {
      templateExpression = '(';
    } else if (templateKey === 'closeParen') {
      templateExpression = ')';
    } else if (templateKey === 'and') {
      templateExpression = '&&';
    } else if (templateKey === 'or') {
      templateExpression = '||';
    } else if (templateKey === 'not') {
      templateExpression = '!';
    }

    if (!templateExpression) {
      this._validationDesignerMessage = 'Select a helper template to insert.';
      this.context.propertyPane.refresh();
      return;
    }

    this.appendToValidationExpression(templateExpression);
  }

  private handleWrapValidationExpressionWithParentheses(): void {
    var currentExpression = String(this.properties.validationDesignerExpression || '').trim();
    if (!currentExpression) {
      this._validationDesignerMessage = 'Enter or build an expression before wrapping with parentheses.';
      this.context.propertyPane.refresh();
      return;
    }

    this.properties.validationDesignerExpression = '(' + currentExpression + ')';
    this._validationDesignerMessage = 'Expression wrapped with parentheses.';
    this.context.propertyPane.refresh();
  }

  private handleClearValidationExpression(): void {
    this.properties.validationDesignerExpression = '';
    this._validationDesignerMessage = 'Expression cleared.';
    this.context.propertyPane.refresh();
  }

  private handleInsertValidationComparisonTemplate(): void {
    var selectedField = String(this.properties.validationDesignerExpressionField || '').trim();
    if (!selectedField) {
      this._validationDesignerMessage = 'Select a field first, then insert a comparison template.';
      this.context.propertyPane.refresh();
      return;
    }

    var escapedField = selectedField.replace(/"/g, '\\"');
    this.appendToValidationExpression('field("' + escapedField + '") == "value"');
  }

  private handleApplyValidationOneClickTemplate(): void {
    var templateKey = String(this.properties.validationDesignerOneClickTemplate || '').trim();
    var expression = '';
    var defaultMessage = '';

    if (templateKey === 'dateRange') {
      expression = 'field("StartDate") <= field("EndDate")';
      defaultMessage = 'Start date must be earlier than or equal to end date.';
    } else if (templateKey === 'requiredPair') {
      expression = 'hasValue(field("FieldA")) && hasValue(field("FieldB"))';
      defaultMessage = 'FieldA and FieldB are both required.';
    } else if (templateKey === 'conditionalRequired') {
      expression = 'field("Type") == "X" && hasValue(field("Details"))';
      defaultMessage = 'Details is required when Type is X.';
    }

    if (!expression) {
      this._validationDesignerMessage = 'Select a one-click template first.';
      this.context.propertyPane.refresh();
      return;
    }

    this.properties.validationDesignerExpression = expression;
    if (!String(this.properties.validationDesignerMessage || '').trim()) {
      this.properties.validationDesignerMessage = defaultMessage;
    }
    this._validationDesignerMessage = 'Template applied. Replace field names/placeholders to match your list internal names.';
    this.context.propertyPane.refresh();
  }

  private handleAddValidationDesignerRule(): void {
    try {
      var expression = String(this.properties.validationDesignerExpression || '').trim();
      var message = String(this.properties.validationDesignerMessage || '').trim();
      var targetField = String(this.properties.validationDesignerTargetField || '').trim();

      if (!expression) {
        this._validationDesignerMessage = 'Enter an expression before adding a validation rule.';
        this.context.propertyPane.refresh();
        return;
      }

      if (!message) {
        this._validationDesignerMessage = 'Enter a message before adding a validation rule.';
        this.context.propertyPane.refresh();
        return;
      }

      var rules = this.parseAdvancedValidationRulesArray(this.properties.advancedValidationJson);
      rules.push({
        expression: expression,
        message: message,
        targetField: targetField || undefined,
      });

      this.properties.advancedValidationJson = JSON.stringify(rules, null, 2);
      this.properties.advancedValidationEnabled = true;
      this._validationDesignerMessage = 'Validation rule added.';
      this._advancedValidationJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._validationDesignerMessage = error && error.message ? error.message : 'Failed to add validation rule.';
    }

    this.context.propertyPane.refresh();
  }

  private handleLoadSelectedValidationRule(): void {
    try {
      var index = parseInt(String(this.properties.validationDesignerSelectedIndex || ''), 10);
      var rules = this.parseAdvancedValidationRulesArray(this.properties.advancedValidationJson);
      if (isNaN(index) || index < 0 || index >= rules.length) {
        this._validationDesignerMessage = 'Select a valid validation rule to load.';
        this.context.propertyPane.refresh();
        return;
      }

      var selected = rules[index] || {};
      this.properties.validationDesignerExpression = selected.expression ? String(selected.expression) : '';
      this.properties.validationDesignerMessage = selected.message ? String(selected.message) : '';
      this.properties.validationDesignerTargetField = selected.targetField ? String(selected.targetField) : '';
      this._validationDesignerMessage = 'Validation rule loaded into designer fields.';
    } catch (error) {
      this._validationDesignerMessage = error && error.message ? error.message : 'Failed to load validation rule.';
    }

    this.context.propertyPane.refresh();
  }

  private handleUpdateSelectedValidationRule(): void {
    try {
      var index = parseInt(String(this.properties.validationDesignerSelectedIndex || ''), 10);
      var rules = this.parseAdvancedValidationRulesArray(this.properties.advancedValidationJson);
      if (isNaN(index) || index < 0 || index >= rules.length) {
        this._validationDesignerMessage = 'Select a valid validation rule to update.';
        this.context.propertyPane.refresh();
        return;
      }

      var expression = String(this.properties.validationDesignerExpression || '').trim();
      var message = String(this.properties.validationDesignerMessage || '').trim();
      var targetField = String(this.properties.validationDesignerTargetField || '').trim();
      if (!expression) {
        this._validationDesignerMessage = 'Enter an expression before updating the validation rule.';
        this.context.propertyPane.refresh();
        return;
      }
      if (!message) {
        this._validationDesignerMessage = 'Enter a message before updating the validation rule.';
        this.context.propertyPane.refresh();
        return;
      }

      rules[index] = {
        expression: expression,
        message: message,
        targetField: targetField || undefined,
      };

      this.properties.advancedValidationJson = JSON.stringify(rules, null, 2);
      this._validationDesignerMessage = 'Selected validation rule updated.';
      this._advancedValidationJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._validationDesignerMessage = error && error.message ? error.message : 'Failed to update validation rule.';
    }

    this.context.propertyPane.refresh();
  }

  private handleRemoveSelectedValidationRule(): void {
    try {
      var index = parseInt(String(this.properties.validationDesignerSelectedIndex || ''), 10);
      var rules = this.parseAdvancedValidationRulesArray(this.properties.advancedValidationJson);
      if (isNaN(index) || index < 0 || index >= rules.length) {
        this._validationDesignerMessage = 'Select a valid validation rule to remove.';
        this.context.propertyPane.refresh();
        return;
      }

      rules.splice(index, 1);
      this.properties.advancedValidationJson = JSON.stringify(rules, null, 2);
      this.properties.validationDesignerSelectedIndex = '';
      this._validationDesignerMessage = 'Selected validation rule removed.';
      this._advancedValidationJsonValidationMessage = rules.length === 0 ? strings.JsonValidationEmpty : strings.JsonValidationSuccess;
    } catch (error) {
      this._validationDesignerMessage = error && error.message ? error.message : 'Failed to remove validation rule.';
    }

    this.context.propertyPane.refresh();
  }

  private handleResetAdvancedValidationJson(): void {
    this.properties.advancedValidationJson = '';
    this.properties.advancedValidationEnabled = false;
    this._advancedValidationJsonValidationMessage = strings.JsonValidationEmpty;
    this._validationDesignerMessage = 'Validation rules JSON cleared.';
    this.context.propertyPane.refresh();
  }

  private parseRequiredRulesArray(raw: string | undefined): any[] {
    var source = String(raw || '').trim();
    if (!source) {
      return [];
    }

    var parsed = JSON.parse(source);
    if (!Array.isArray(parsed)) {
      throw new Error(strings.JsonValidationExpectedArray);
    }

    for (var i = 0; i < parsed.length; i += 1) {
      var rule = parsed[i];
      if (!rule || typeof rule !== 'object') {
        throw new Error('Required rule at index ' + i + ' must be an object.');
      }

      var field = rule.field === undefined || rule.field === null ? '' : String(rule.field).trim();
      if (!field) {
        throw new Error('Required rule at index ' + i + ' is missing field.');
      }
    }

    return parsed;
  }

  private handleValidateRequiredRulesJson(): void {
    try {
      this.parseRequiredRulesArray(this.properties.requiredRulesJson);
      this._requiredRulesJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._requiredRulesJsonValidationMessage = error && error.message ? error.message : 'Required rules JSON is invalid.';
    }
    this.context.propertyPane.refresh();
  }

  private handleAddRequiredDesignerRule(): void {
    try {
      var field = String(this.properties.requiredDesignerField || '').trim();
      var requiredMode = this.properties.requiredDesignerRequired || 'required';
      var message = String(this.properties.requiredDesignerMessage || '').trim();

      if (!field) {
        this._requiredDesignerMessage = 'Select a field before adding a required rule.';
        this.context.propertyPane.refresh();
        return;
      }

      var rules = this.parseRequiredRulesArray(this.properties.requiredRulesJson);
      rules.push({
        field: field,
        required: requiredMode === 'required',
        message: message || undefined,
      });

      this.properties.requiredRulesJson = JSON.stringify(rules, null, 2);
      this._requiredDesignerMessage = 'Required rule added.';
      this._requiredRulesJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._requiredDesignerMessage = error && error.message ? error.message : 'Failed to add required rule.';
    }

    this.context.propertyPane.refresh();
  }

  private handleLoadSelectedRequiredRule(): void {
    try {
      var index = parseInt(String(this.properties.requiredDesignerSelectedIndex || ''), 10);
      var rules = this.parseRequiredRulesArray(this.properties.requiredRulesJson);
      if (isNaN(index) || index < 0 || index >= rules.length) {
        this._requiredDesignerMessage = 'Select a valid required rule to load.';
        this.context.propertyPane.refresh();
        return;
      }

      var selected = rules[index] || {};
      this.properties.requiredDesignerField = selected.field ? String(selected.field) : '';
      this.properties.requiredDesignerRequired = selected.required === false ? 'optional' : 'required';
      this.properties.requiredDesignerMessage = selected.message ? String(selected.message) : '';
      this._requiredDesignerMessage = 'Required rule loaded into designer fields.';
    } catch (error) {
      this._requiredDesignerMessage = error && error.message ? error.message : 'Failed to load required rule.';
    }

    this.context.propertyPane.refresh();
  }

  private handleUpdateSelectedRequiredRule(): void {
    try {
      var index = parseInt(String(this.properties.requiredDesignerSelectedIndex || ''), 10);
      var rules = this.parseRequiredRulesArray(this.properties.requiredRulesJson);
      if (isNaN(index) || index < 0 || index >= rules.length) {
        this._requiredDesignerMessage = 'Select a valid required rule to update.';
        this.context.propertyPane.refresh();
        return;
      }

      var field = String(this.properties.requiredDesignerField || '').trim();
      var requiredMode = this.properties.requiredDesignerRequired || 'required';
      var message = String(this.properties.requiredDesignerMessage || '').trim();
      if (!field) {
        this._requiredDesignerMessage = 'Select a field before updating the required rule.';
        this.context.propertyPane.refresh();
        return;
      }

      rules[index] = {
        field: field,
        required: requiredMode === 'required',
        message: message || undefined,
      };

      this.properties.requiredRulesJson = JSON.stringify(rules, null, 2);
      this._requiredDesignerMessage = 'Selected required rule updated.';
      this._requiredRulesJsonValidationMessage = strings.JsonValidationSuccess;
    } catch (error) {
      this._requiredDesignerMessage = error && error.message ? error.message : 'Failed to update required rule.';
    }

    this.context.propertyPane.refresh();
  }

  private handleRemoveSelectedRequiredRule(): void {
    try {
      var index = parseInt(String(this.properties.requiredDesignerSelectedIndex || ''), 10);
      var rules = this.parseRequiredRulesArray(this.properties.requiredRulesJson);
      if (isNaN(index) || index < 0 || index >= rules.length) {
        this._requiredDesignerMessage = 'Select a valid required rule to remove.';
        this.context.propertyPane.refresh();
        return;
      }

      rules.splice(index, 1);
      this.properties.requiredRulesJson = JSON.stringify(rules, null, 2);
      this.properties.requiredDesignerSelectedIndex = '';
      this._requiredDesignerMessage = 'Selected required rule removed.';
      this._requiredRulesJsonValidationMessage = rules.length === 0 ? strings.JsonValidationEmpty : strings.JsonValidationSuccess;
    } catch (error) {
      this._requiredDesignerMessage = error && error.message ? error.message : 'Failed to remove required rule.';
    }

    this.context.propertyPane.refresh();
  }

  private handleResetRequiredRulesJson(): void {
    this.properties.requiredRulesJson = '';
    this._requiredRulesJsonValidationMessage = strings.JsonValidationEmpty;
    this._requiredDesignerMessage = 'Required rules JSON cleared.';
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.promoteRuntimeDynamicProperties();
    this.pushDynamicDiagnostic('Property pane opened. current itemRef=' + String(this.properties.dynamicItemIdReference || '(empty)') + ', current source=' + String(this.properties.dynamicPreferredSourceInstanceId || '(empty)'));
    var fullVersionLabel = 'Version: ' + this.getWebPartVersion();
    var fontFamilyOptions = this.getFontFamilyOptions();
    var useDynamicItemAsItemId = this.properties.useDynamicItemIdAsItemId !== false;
    var dynamicTargetFieldOptions = this._dynamicTargetLookupFields.length > 0
      ? this._dynamicTargetLookupFields
      : [{ key: '', text: strings.PropDynamicTargetFieldNone }];
    var permissionLookupFieldOptions = this._lookupPermissionFields.length > 0 ? this._lookupPermissionFields : [{ key: '', text: strings.PropPermissionBaseLookupFieldNone }];
    var isButtonCornerRounded = (this.properties.buttonCornerStyle || 'square') === 'rounded';
    var filterConditionOptions: IDropdownOption[] = [{ key: '', text: strings.PropFilterDesignerExistingNone }];
    var defaultEntryOptions: IDropdownOption[] = [{ key: '', text: strings.PropDefaultDesignerExistingNone }];
    var validationRuleOptions: IDropdownOption[] = [{ key: '', text: 'No validation rules in JSON' }];

    try {
      var existingConditions = this.parseFilterJsonArray(this.properties.filterJson);
      if (existingConditions.length > 0) {
        filterConditionOptions = existingConditions.map(function(condition: any, index: number) {
          var field = String((condition && condition.field) || '(field)');
          var type = String((condition && condition.type) || 'text');
          var operator = String((condition && condition.operator) || 'eq').toLowerCase();
          var logical = String((condition && condition.logical) || 'and').toUpperCase();
          var value = condition && condition.value !== undefined ? String(condition.value) : '';
          var logicalPrefix = index === 0 ? '' : logical + ' ';
          return {
            key: String(index),
            text: '#' + String(index + 1) + ': ' + logicalPrefix + field + ' [' + type + '] ' + operator + ' ' + value
          };
        });
      }
    } catch (_parseFilterError) {
      // Keep fallback option when JSON is invalid.
    }

    try {
      var existingDefaults = this.parseDefaultValuesObject(this.properties.defaultValuesJson);
      var defaultKeys = Object.keys(existingDefaults);
      if (defaultKeys.length > 0) {
        defaultEntryOptions = defaultKeys.map(function(key: string) {
          var value = existingDefaults[key];
          var preview = typeof value === 'object' ? JSON.stringify(value) : String(value);
          return {
            key: key,
            text: key + ': ' + preview
          };
        });
      }
    } catch (_parseDefaultsError) {
      // Keep fallback option when JSON is invalid.
    }

    try {
      var existingValidationRules = this.parseAdvancedValidationRulesArray(this.properties.advancedValidationJson);
      if (existingValidationRules.length > 0) {
        validationRuleOptions = existingValidationRules.map(function(rule: any, index: number) {
          var expression = String((rule && rule.expression) || '').trim();
          var target = String((rule && rule.targetField) || '').trim();
          return {
            key: String(index),
            text: '#' + String(index + 1) + ': ' + expression + (target ? ' -> ' + target : '')
          };
        });
      }
    } catch (_parseValidationRulesError) {
      // Keep fallback option when JSON is invalid.
    }

    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: '',
              groupFields: [
                this.createPropertyPanePageTitle('mainSettingsPageTitle', 'Main settings')
              ]
            },
            {
              groupName: fullVersionLabel,
              groupFields: [
                PropertyPaneLabel('propertyPaneVersionInfo', {
                  text: ' '
                })
              ]
            },
            {
              groupName: 'Page Guide',
              groupFields: [
                PropertyPaneLabel('propertyPanePageGuide1', {
                  text: 'Page 1: Configure the list, mode, dynamic data, and permissions.'
                }),
                PropertyPaneLabel('propertyPanePageGuide2', {
                  text: 'Page 2: Configure form buttons, redirects, and button styles.'
                }),
                PropertyPaneLabel('propertyPanePageGuide3', {
                  text: 'Page 3: Build record filters.'
                }),
                PropertyPaneLabel('propertyPanePageGuide4', {
                  text: 'Page 4: Build advanced validation rules.'
                }),
                PropertyPaneLabel('propertyPanePageGuide5', {
                  text: 'Page 5: Configure default field values.'
                })
              ]
            },
            {
              groupName: strings.PropertyGroupBasic,
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: strings.PropListLabel,
                  options: this._lists,
                  selectedKey: this.properties.listName,
                }),
                PropertyPaneDropdown('mode', {
                  label: strings.PropModeLabel,
                  options: [
                    { key: 'new', text: strings.PropModeNew },
                    { key: 'edit', text: strings.PropModeEdit },
                    { key: 'view', text: strings.PropModeView },
                  ],
                }),
                PropertyPaneCheckbox('forceFullWidth', { text: 'Force Full Width' }),
                ...(!this.properties.forceFullWidth ? [
                  PropertyPaneSlider('fixedWidth', {
                    label: 'Fixed width (px)',
                    min: 480,
                    max: 1600,
                    step: 20,
                    value: toPositiveNumber(this.properties.fixedWidth, 800),
                    showValue: true,
                  })
                ] : []),
                PropertyPaneTextField('itemIdQueryParam', {
                  label: strings.PropItemIdQueryParamLabel,
                  placeholder: 'itemid',
                  value: this.properties.itemIdQueryParam || 'itemid',
                }),
                PropertyPaneDropdown('dynamicPreferredSourceInstanceId', {
                  label: strings.PropDynamicItemIdLabel,
                  options: this._listControlSources,
                  selectedKey: this.properties.dynamicPreferredSourceInstanceId || ''
                }),
                PropertyPaneDropdown('permissionBaseLookupField', {
                  label: strings.PropPermissionBaseLookupFieldLabel,
                  options: permissionLookupFieldOptions,
                  selectedKey: this.properties.permissionBaseLookupField || ''
                }),
                ...(this.properties.permissionBaseLookupField ? [
                  PropertyPaneDropdown('permissionScope', {
                    label: strings.PropPermissionScopeLabel,
                    options: [
                      { key: 'list', text: strings.PropPermissionScopeList },
                      { key: 'item', text: strings.PropPermissionScopeItem },
                    ],
                    selectedKey: this.properties.permissionScope || 'list'
                  })
                ] : []),
                ...(this.properties.permissionBaseLookupField && this.properties.permissionScope === 'item' ? [
                  PropertyPaneLabel('permissionItemHelperTitle', {
                    text: 'Add mode has no item yet, so pick the exact item that should determine permissions. This item is used for both Add and Edit.'
                  }),
                  PropertyPaneButton('loadPermissionLookupItems', {
                    text: 'Refresh lookup items',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleLoadPermissionLookupItems.bind(this)
                  }),
                  PropertyPaneDropdown('permissionScopeLookupPick', {
                    label: 'Select item',
                    options: this._permissionLookupItemOptions.length > 0 ? this._permissionLookupItemOptions : [{ key: '', text: 'No items loaded yet' }],
                    selectedKey: this.properties.permissionScopeLookupPick || ''
                  }),
                  PropertyPaneLabel('permissionLookupHelperResult', {
                    text: this._permissionLookupMessage || 'Lookup items load automatically when you select the column.'
                  }),
                ] : []),
                PropertyPaneToggle('useDynamicItemIdAsItemId', {
                  label: strings.PropUseDynamicItemAsItemIdLabel,
                  onText: strings.PropToggleOn,
                  offText: strings.PropToggleOff,
                  checked: this.properties.useDynamicItemIdAsItemId !== false,
                }),
                ...(!useDynamicItemAsItemId ? [
                  PropertyPaneDropdown('dynamicItemTargetField', {
                    label: strings.PropDynamicItemTargetFieldLabel,
                    options: dynamicTargetFieldOptions,
                    selectedKey: this.properties.dynamicItemTargetField || '',
                  })
                ] : []),
                PropertyPaneTextField('itemId', {
                  label: strings.PropItemIdLabel,
                  placeholder: strings.PropItemIdPlaceholder,
                  value: this.properties.itemId ? String(this.properties.itemId) : '',
                  description: strings.PropItemIdHelp,
                }),
                PropertyPaneToggle('showFieldDescription', {
                  label: strings.PropShowFieldDescriptionLabel,
                  onText: strings.PropToggleOn,
                  offText: strings.PropToggleOff,
                  checked: this.properties.showFieldDescription,
                }),
                PropertyPaneToggle('enableDynamicDiagnostics', {
                  label: strings.PropEnableDynamicDiagnosticsLabel,
                  onText: strings.PropToggleOn,
                  offText: strings.PropToggleOff,
                  checked: this.properties.enableDynamicDiagnostics !== false,
                }),
              ]
            },
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
                this.createPropertyPanePageTitle('buttonSettingsPageTitle', 'Button settings')
              ]
            },
            {
              groupName: '',
              groupFields: [
                PropertyPaneTextField('addSubmitButtonLabel', {
                  label: strings.PropAddSubmitButtonLabel,
                  placeholder: strings.PropAddSubmitButtonPlaceholder,
                  value: this.properties.addSubmitButtonLabel,
                }),
                PropertyPaneTextField('editSubmitButtonLabel', {
                  label: strings.PropEditSubmitButtonLabel,
                  placeholder: strings.PropEditSubmitButtonPlaceholder,
                  value: this.properties.editSubmitButtonLabel,
                }),
                PropertyPaneTextField('submitButtonLabel', {
                  label: strings.PropSubmitButtonLabel,
                  placeholder: strings.PropSubmitButtonPlaceholder,
                  value: this.properties.submitButtonLabel,
                }),
                PropertyPaneToggle('showCancelButton', {
                  label: strings.PropShowCancelButtonLabel,
                  onText: strings.PropToggleOn,
                  offText: strings.PropToggleOff,
                  checked: this.properties.showCancelButton,
                }),
                ...(this.properties.showCancelButton ? [
                  PropertyPaneTextField('cancelButtonLabel', {
                    label: strings.PropCancelButtonLabel,
                    placeholder: strings.PropCancelButtonPlaceholder,
                    value: this.properties.cancelButtonLabel,
                  })
                ] : []),
                PropertyPaneTextField('cancelRedirectUrl', {
                  label: strings.PropCancelRedirectUrlLabel,
                  placeholder: strings.PropCancelRedirectUrlPlaceholder,
                  value: this.properties.cancelRedirectUrl,
                }),
                PropertyPaneTextField('submitRedirectUrl', {
                  label: strings.PropSubmitRedirectUrlLabel,
                  placeholder: strings.PropSubmitRedirectUrlPlaceholder,
                  value: this.properties.submitRedirectUrl,
                }),
                PropertyPaneTextField('onSubmitMessage', {
                  label: strings.PropSubmitSuccessMessageLabel,
                  placeholder: strings.PropSubmitSuccessMessagePlaceholder,
                  value: this.properties.onSubmitMessage,
                  multiline: true,
                  resizable: true,
                  rows: 2,
                }),
                PropertyFieldColorPicker('buttonTextColor', {
                  label: strings.PropButtonTextColorLabel,
                  selectedColor: normalizeColorValue(this.properties.buttonTextColor, '#000000'),
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'buttonTextColorField'
                }),
                PropertyFieldColorPicker('buttonBackgroundColor', {
                  label: strings.PropButtonBackgroundColorLabel,
                  selectedColor: normalizeColorValue(this.properties.buttonBackgroundColor, '#f0f0f0'),
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'buttonBackgroundColorField'
                }),
                PropertyFieldColorPicker('buttonBorderColor', {
                  label: strings.PropButtonBorderColorLabel,
                  selectedColor: normalizeColorValue(this.properties.buttonBorderColor, '#f0f0f0'),
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'buttonBorderColorField'
                }),
                PropertyPaneSlider('buttonBorderWidth', {
                  label: strings.PropButtonBorderWidthLabel,
                  min: 0,
                  max: 10,
                  step: 1,
                  value: toPositiveNumber(this.properties.buttonBorderWidth, 1),
                }),
                PropertyPaneDropdown('buttonFontFamily', {
                  label: strings.PropButtonFontFamilyLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.buttonFontFamily || 'inherit',
                }),
                PropertyPaneSlider('buttonFontSize', {
                  label: strings.PropButtonFontSizeLabel,
                  min: 10,
                  max: 40,
                  step: 1,
                  value: toPositiveNumber(this.properties.buttonFontSize, 14),
                }),
                PropertyPaneDropdown('buttonFontStyle', {
                  label: strings.PropButtonFontStyleLabel,
                  options: [
                    { key: 'normal', text: strings.PropFontStyleNormal },
                    { key: 'italic', text: strings.PropFontStyleItalic },
                  ],
                  selectedKey: this.properties.buttonFontStyle || 'normal',
                }),
                PropertyPaneToggle('buttonFontBold', {
                  label: strings.PropButtonFontBoldLabel,
                  onText: strings.PropToggleOn,
                  offText: strings.PropToggleOff,
                  checked: this.properties.buttonFontBold === true,
                }),
                PropertyPaneDropdown('buttonCornerStyle', {
                  label: strings.PropButtonCornerStyleLabel,
                  options: [
                    { key: 'square', text: strings.PropCornerStyleSquare },
                    { key: 'rounded', text: strings.PropCornerStyleRounded }
                  ],
                  selectedKey: this.properties.buttonCornerStyle || 'square',
                }),
                ...(isButtonCornerRounded ? [
                  PropertyPaneSlider('buttonCornerRadius', {
                    label: strings.PropButtonCornerRadiusLabel,
                    min: 0,
                    max: 40,
                    step: 1,
                    value: toPositiveNumber(this.properties.buttonCornerRadius, 4),
                  })
                ] : []),
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
                this.createPropertyPanePageTitle('filterDesignerPageTitle', 'Filter Designer')
              ]
            },
            {
              groupName: '',
              groupFields: [
                PropertyPaneLabel('filterDesignerHelp', {
                  text: 'Supports AND/OR and operators. Example: Status ne "Closed" OR Status ne "Cancelled". Use "Join with previous" to chain conditions.'
                }),
                PropertyPaneDropdown('filterDesignerField', {
                  label: strings.PropFilterDesignerFieldLabel,
                  options: this._listFields,
                  selectedKey: this.properties.filterDesignerField || '',
                }),
                PropertyPaneDropdown('filterDesignerType', {
                  label: strings.PropFilterDesignerTypeLabel,
                  options: [
                    { key: 'text', text: strings.PropertyFieldTypeText },
                    { key: 'number', text: strings.PropertyFieldTypeNumber },
                    { key: 'lookup', text: strings.PropertyFieldTypeLookup },
                    { key: 'boolean', text: strings.PropertyFieldTypeBoolean },
                    { key: 'datetime', text: strings.PropertyFieldTypeDatetime },
                  ],
                  selectedKey: this.properties.filterDesignerType || 'text',
                }),
                PropertyPaneDropdown('filterDesignerOperator', {
                  label: 'Operator',
                  options: [
                    { key: 'eq', text: 'Equals (eq)' },
                    { key: 'ne', text: 'Not equal (ne)' },
                    { key: 'gt', text: 'Greater than (gt)' },
                    { key: 'ge', text: 'Greater or equal (ge)' },
                    { key: 'lt', text: 'Less than (lt)' },
                    { key: 'le', text: 'Less or equal (le)' },
                    { key: 'contains', text: 'Contains (text)' },
                    { key: 'startswith', text: 'Starts with (text)' },
                    { key: 'endswith', text: 'Ends with (text)' },
                  ],
                  selectedKey: this.properties.filterDesignerOperator || 'eq',
                }),
                PropertyPaneDropdown('filterDesignerLogical', {
                  label: 'Join with previous',
                  options: [
                    { key: 'and', text: 'AND' },
                    { key: 'or', text: 'OR' },
                  ],
                  selectedKey: this.properties.filterDesignerLogical || 'and',
                }),
                PropertyPaneDropdown('filterDesignerValueSource', {
                  label: strings.PropFilterDesignerValueSourceLabel,
                  options: [
                    { key: 'dynamic', text: strings.PropFilterDesignerValueSourceDynamic },
                    { key: 'static', text: strings.PropFilterDesignerValueSourceStatic },
                    { key: 'expression', text: strings.PropExpressionValueType },
                    ...(this.supportsFilterFieldValue(this.properties.filterDesignerField || '') ? [{ key: 'fieldValue', text: 'Field value' }] : [])
                  ],
                  selectedKey: this.properties.filterDesignerValueSource || 'dynamic',
                }),
                ...((this.properties.filterDesignerValueSource || 'dynamic') !== 'dynamic' && (this.properties.filterDesignerValueSource || 'dynamic') !== 'fieldValue' ? [
                  PropertyPaneTextField('filterDesignerValue', {
                    label: (this.properties.filterDesignerValueSource || 'dynamic') === 'expression' ? strings.PropExpressionLabel : strings.PropFilterDesignerValueLabel,
                    placeholder: (this.properties.filterDesignerValueSource || 'dynamic') === 'expression' ? strings.PropExpressionPlaceholder : '',
                    value: this.properties.filterDesignerValue || '',
                  }),
                  ...((this.properties.filterDesignerValueSource || 'dynamic') === 'expression' ? [this.createExpressionHelpField('filterExpressionHelp', false)] : [])
                ] : []),
                ...((this.properties.filterDesignerValueSource || 'dynamic') === 'fieldValue' && !this.isMultiFilterField(this.properties.filterDesignerField || '') && (this._fieldChoicesByInternalName[this.properties.filterDesignerField || ''] || []).length > 0 ? [
                  PropertyPaneDropdown('filterDesignerValue', {
                    label: 'Select field value',
                    options: (this._fieldChoicesByInternalName[this.properties.filterDesignerField || ''] || []).map(function(choice: string) { return { key: choice, text: choice }; }),
                    selectedKey: this.properties.filterDesignerValue || ''
                  })
                ] : []),
                ...((this.properties.filterDesignerValueSource || 'dynamic') === 'fieldValue' && this.isMultiFilterField(this.properties.filterDesignerField || '') && (this._fieldChoicesByInternalName[this.properties.filterDesignerField || ''] || []).length > 0 ? [
                  this.createFilterMultiValuePicker((this._fieldChoicesByInternalName[this.properties.filterDesignerField || ''] || []).map(function(choice: string) { return { key: choice, text: choice }; }))
                ] : []),
                ...((this.properties.filterDesignerType || 'text') === 'lookup' && ((this.properties.filterDesignerValueSource || 'dynamic') === 'static' || (this.properties.filterDesignerValueSource || 'dynamic') === 'fieldValue') ? [
                  PropertyPaneLabel('filterLookupHelperTitle', {
                    text: 'Lookup item picker: choose an item from the target list to fill in its ID.'
                  }),
                  PropertyPaneButton('loadFilterLookupItems', {
                    text: 'Refresh lookup items',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleLoadFilterLookupItems.bind(this)
                  }),
                  PropertyPaneDropdown('filterDesignerLookupPick', {
                    label: 'Select item',
                    options: this._filterLookupItemOptions.length > 0 ? this._filterLookupItemOptions : [{ key: '', text: 'No items loaded yet' }],
                    selectedKey: this.properties.filterDesignerLookupPick || '',
                  }),
                  PropertyPaneButton('applyFilterLookupPick', {
                    text: 'Use selected ID',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleApplyFilterLookupPick.bind(this)
                  }),
                  PropertyPaneLabel('filterLookupHelperResult', {
                    text: this._filterLookupMessage || 'Lookup items load automatically when you select the column.'
                  }),
                ] : []),
                PropertyPaneButton('applyFilterDesignerCondition', {
                  text: strings.PropFilterDesignerAddButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleAddFilterDesignerCondition.bind(this)
                }),
                PropertyPaneDropdown('filterDesignerSelectedIndex', {
                  label: strings.PropFilterDesignerExistingLabel,
                  options: filterConditionOptions,
                  selectedKey: this.properties.filterDesignerSelectedIndex || '',
                }),
                PropertyPaneButton('loadSelectedFilterDesignerCondition', {
                  text: strings.PropFilterDesignerLoadButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleLoadSelectedFilterCondition.bind(this)
                }),
                PropertyPaneButton('updateSelectedFilterDesignerCondition', {
                  text: strings.PropFilterDesignerUpdateButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleUpdateSelectedFilterCondition.bind(this)
                }),
                PropertyPaneButton('removeSelectedFilterDesignerCondition', {
                  text: strings.PropFilterDesignerRemoveButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleRemoveSelectedFilterCondition.bind(this)
                }),
                PropertyPaneButton('resetFilterJsonDesigner', {
                  text: strings.PropFilterDesignerResetButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleResetFilterJson.bind(this)
                }),
                PropertyPaneLabel('filterDesignerResult', {
                  text: this._filterDesignerMessage || strings.PropFilterDesignerStatusPlaceholder
                }),
                PropertyPaneTextField('filterJson', {
                  label: strings.PropFilterJsonLabel,
                  placeholder: strings.PropFilterJsonPlaceholder,
                  value: this.properties.filterJson,
                  multiline: true,
                  resizable: true,
                  rows: 4,
                  description: strings.PropFilterJsonHelp,
                }),
                PropertyPaneButton('validateFilterJson', {
                  text: strings.PropValidateFilterJsonButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleValidateFilterJson.bind(this)
                }),
                PropertyPaneLabel('validateFilterJsonResult', {
                  text: this._filterJsonValidationMessage || strings.JsonValidationResultPlaceholder
                }),
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
                this.createPropertyPanePageTitle('validationDesignerPageTitle', 'Validation Designer')
              ]
            },
            {
              groupName: '',
              groupFields: [
                PropertyPaneLabel('validationDesignerHelp', {
                  text: 'Use the builder controls below so you do not need to manually type internal names. Expressions support parentheses and operators: == != > >= < <= && || !'
                }),
                PropertyPaneLabel('validationDesignerHelpFunctions', {
                  text: 'Functions: field("InternalName"), today(), now(), daysFromToday(n), isEmpty(x), hasValue(x), between(x,a,b).'
                }),
                PropertyPaneLabel('validationDesignerHelpDocsLink', {
                  text: 'Operators reference: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Operator_precedence'
                }),
                PropertyPaneToggle('advancedValidationEnabled', {
                  label: 'Enable advanced validation rules',
                  onText: strings.PropToggleOn,
                  offText: strings.PropToggleOff,
                  checked: this.properties.advancedValidationEnabled === true,
                }),
                PropertyPaneDropdown('validationDesignerExpressionField', {
                  label: 'Expression field selector',
                  options: dynamicTargetFieldOptions,
                  selectedKey: this.properties.validationDesignerExpressionField || '',
                }),
                PropertyPaneButton('insertValidationExpressionField', {
                  text: 'Insert selected field token',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleInsertSelectedValidationExpressionField.bind(this)
                }),
                PropertyPaneButton('insertValidationComparisonTemplate', {
                  text: 'Insert selected field == "value"',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleInsertValidationComparisonTemplate.bind(this)
                }),
                PropertyPaneDropdown('validationDesignerExpressionTemplate', {
                  label: 'Expression helper',
                  options: [
                    { key: '', text: 'Select helper to insert' },
                    { key: 'today', text: 'today()' },
                    { key: 'now', text: 'now()' },
                    { key: 'daysFromToday', text: 'daysFromToday(7)' },
                    { key: 'isEmpty', text: 'isEmpty(field("InternalName"))' },
                    { key: 'hasValue', text: 'hasValue(field("InternalName"))' },
                    { key: 'between', text: 'between(field("InternalName"), 1, 10)' },
                    { key: 'eq', text: 'Equals operator (==)' },
                    { key: 'ne', text: 'Not-equals operator (!=)' },
                    { key: 'gt', text: 'Greater than (>)' },
                    { key: 'ge', text: 'Greater or equal (>=)' },
                    { key: 'lt', text: 'Less than (<)' },
                    { key: 'le', text: 'Less or equal (<=)' },
                    { key: 'openParen', text: 'Open parenthesis (()' },
                    { key: 'closeParen', text: 'Close parenthesis ())' },
                    { key: 'and', text: 'AND operator (&&)' },
                    { key: 'or', text: 'OR operator (||)' },
                    { key: 'not', text: 'NOT operator (!)' },
                  ],
                  selectedKey: this.properties.validationDesignerExpressionTemplate || '',
                }),
                PropertyPaneButton('insertValidationExpressionTemplate', {
                  text: 'Insert selected helper',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleInsertValidationExpressionTemplate.bind(this)
                }),
                PropertyPaneDropdown('validationDesignerOneClickTemplate', {
                  label: 'One-click template',
                  options: [
                    { key: '', text: 'Select one-click template' },
                    { key: 'dateRange', text: 'Date range: StartDate <= EndDate' },
                    { key: 'requiredPair', text: 'Required pair: FieldA and FieldB required' },
                    { key: 'conditionalRequired', text: 'Conditional required: if Type==X then Details required' },
                  ],
                  selectedKey: this.properties.validationDesignerOneClickTemplate || '',
                }),
                PropertyPaneButton('applyValidationOneClickTemplate', {
                  text: 'Apply one-click template',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleApplyValidationOneClickTemplate.bind(this)
                }),
                PropertyPaneButton('wrapValidationExpression', {
                  text: 'Wrap expression with (...)',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleWrapValidationExpressionWithParentheses.bind(this)
                }),
                PropertyPaneButton('clearValidationExpression', {
                  text: 'Clear expression',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleClearValidationExpression.bind(this)
                }),
                PropertyPaneDropdown('validationDesignerTargetField', {
                  label: 'Message target field (optional)',
                  options: dynamicTargetFieldOptions,
                  selectedKey: this.properties.validationDesignerTargetField || '',
                }),
                PropertyPaneTextField('validationDesignerExpression', {
                  label: 'Expression',
                  value: this.properties.validationDesignerExpression || '',
                  placeholder: '(field("DueDate") > today()) && hasValue(field("Status"))',
                  description: 'Tip: Build expression using the selectors above to avoid internal-name typos.',
                }),
                PropertyPaneTextField('validationDesignerMessage', {
                  label: 'Validation message',
                  value: this.properties.validationDesignerMessage || '',
                  placeholder: 'Due date must be later than today',
                }),
                PropertyPaneButton('addValidationDesignerRule', {
                  text: 'Add rule to JSON',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleAddValidationDesignerRule.bind(this)
                }),
                PropertyPaneDropdown('validationDesignerSelectedIndex', {
                  label: 'Existing rules',
                  options: validationRuleOptions,
                  selectedKey: this.properties.validationDesignerSelectedIndex || '',
                }),
                PropertyPaneButton('loadValidationDesignerRule', {
                  text: 'Load selected rule',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleLoadSelectedValidationRule.bind(this)
                }),
                PropertyPaneButton('updateValidationDesignerRule', {
                  text: 'Update selected rule',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleUpdateSelectedValidationRule.bind(this)
                }),
                PropertyPaneButton('removeValidationDesignerRule', {
                  text: 'Remove selected rule',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleRemoveSelectedValidationRule.bind(this)
                }),
                PropertyPaneButton('resetValidationDesignerRules', {
                  text: 'Clear validation JSON',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleResetAdvancedValidationJson.bind(this)
                }),
                PropertyPaneLabel('validationDesignerResult', {
                  text: this._validationDesignerMessage || 'Validation designer result appears here.'
                }),
                PropertyPaneTextField('advancedValidationJson', {
                  label: 'Validation rules JSON',
                  value: this.properties.advancedValidationJson,
                  multiline: true,
                  resizable: true,
                  rows: 8,
                  description: 'JSON array format: [{"expression":"...","message":"...","targetField":"InternalName(optional)"}]'
                }),
                PropertyPaneButton('validateAdvancedValidationJson', {
                  text: 'Validate validation JSON',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleValidateAdvancedValidationJson.bind(this)
                }),
                PropertyPaneLabel('validateAdvancedValidationJsonResult', {
                  text: this._advancedValidationJsonValidationMessage || strings.JsonValidationResultPlaceholder
                }),
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
                this.createPropertyPanePageTitle('defaultsDesignerPageTitle', 'Defaults Designer')
              ]
            },
            {
              groupName: '',
              groupFields: [
                PropertyPaneDropdown('defaultDesignerField', {
                  label: strings.PropDefaultDesignerFieldLabel,
                  options: dynamicTargetFieldOptions,
                  selectedKey: this.properties.defaultDesignerField || '',
                }),
                PropertyPaneDropdown('defaultDesignerValueType', {
                  label: strings.PropDefaultDesignerValueTypeLabel,
                  options: [
                    { key: 'string', text: strings.PropDefaultDesignerValueTypeString },
                    { key: 'number', text: strings.PropDefaultDesignerValueTypeNumber },
                    { key: 'boolean', text: strings.PropDefaultDesignerValueTypeBoolean },
                    { key: 'json', text: strings.PropDefaultDesignerValueTypeJson },
                    { key: 'expression', text: strings.PropExpressionValueType },
                  ],
                  selectedKey: this.properties.defaultDesignerValueType || 'string',
                }),
                PropertyPaneTextField('defaultDesignerValue', {
                  label: (this.properties.defaultDesignerValueType || 'string') === 'expression' ? strings.PropExpressionLabel : strings.PropDefaultDesignerValueLabel,
                  placeholder: (this.properties.defaultDesignerValueType || 'string') === 'expression' ? strings.PropExpressionPlaceholder : '',
                  value: this.properties.defaultDesignerValue || '',
                  multiline: (this.properties.defaultDesignerValueType || 'string') === 'json',
                  rows: (this.properties.defaultDesignerValueType || 'string') === 'json' ? 3 : 1,
                }),
                ...((this.properties.defaultDesignerValueType || 'string') === 'expression' ? [this.createExpressionHelpField('defaultExpressionHelp', true)] : []),
                ...(this.isLookupTypeField(this.properties.defaultDesignerField || '') ? [
                  PropertyPaneLabel('defaultLookupHelperTitle', {
                    text: 'Lookup item picker: choose an item from the target list to fill in its ID.'
                  }),
                  PropertyPaneButton('loadDefaultLookupItems', {
                    text: 'Refresh lookup items',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleLoadDefaultLookupItems.bind(this)
                  }),
                  PropertyPaneDropdown('defaultDesignerLookupPick', {
                    label: 'Select item',
                    options: this._defaultLookupItemOptions.length > 0 ? this._defaultLookupItemOptions : [{ key: '', text: 'No items loaded yet' }],
                    selectedKey: this.properties.defaultDesignerLookupPick || '',
                  }),
                  PropertyPaneButton('applyDefaultLookupPick', {
                    text: 'Use selected ID',
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.handleApplyDefaultLookupPick.bind(this)
                  }),
                  PropertyPaneLabel('defaultLookupHelperResult', {
                    text: this._defaultLookupMessage || 'Lookup items load automatically when you select the column.'
                  }),
                ] : []),
                PropertyPaneButton('applyDefaultDesignerEntry', {
                  text: strings.PropDefaultDesignerAddButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleAddDefaultDesignerEntry.bind(this)
                }),
                PropertyPaneDropdown('defaultDesignerSelectedField', {
                  label: strings.PropDefaultDesignerExistingLabel,
                  options: defaultEntryOptions,
                  selectedKey: this.properties.defaultDesignerSelectedField || '',
                }),
                PropertyPaneButton('loadSelectedDefaultDesignerEntry', {
                  text: strings.PropDefaultDesignerLoadButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleLoadSelectedDefaultEntry.bind(this)
                }),
                PropertyPaneButton('removeSelectedDefaultDesignerEntry', {
                  text: strings.PropDefaultDesignerRemoveButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleRemoveSelectedDefaultEntry.bind(this)
                }),
                PropertyPaneButton('resetDefaultValuesDesigner', {
                  text: strings.PropDefaultDesignerResetButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleResetDefaultValuesJson.bind(this)
                }),
                PropertyPaneLabel('defaultDesignerResult', {
                  text: this._defaultDesignerMessage || strings.PropDefaultDesignerStatusPlaceholder
                }),
                PropertyPaneTextField('defaultValuesJson', {
                  label: strings.PropDefaultValuesJsonLabel,
                  placeholder: strings.PropDefaultValuesJsonPlaceholder,
                  value: this.properties.defaultValuesJson,
                  multiline: true,
                  resizable: true,
                  rows: 6,
                  description: strings.PropDefaultValuesJsonHelp,
                }),
                PropertyPaneButton('validateDefaultValuesJson', {
                  text: strings.PropValidateDefaultValuesJsonButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.handleValidateDefaultValuesJson.bind(this)
                }),
                PropertyPaneLabel('validateDefaultValuesJsonResult', {
                  text: this._defaultValuesJsonValidationMessage || strings.JsonValidationResultPlaceholder
                }),
              ]
            },
          ]
        },
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
}
