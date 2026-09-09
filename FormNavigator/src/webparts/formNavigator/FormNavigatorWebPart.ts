import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import * as strings from 'FormNavigatorWebPartStrings';
import './FormNavigator.css';

const packageSolutionConfig: any = require('../../../config/package-solution.json');
const DYNAMIC_FORM_COMPONENT_ID = 'da1a4e74-6fba-498c-8cda-af20071f7ed3';
const REPORT_FORM_COMPONENT_ID = 'e47c4f0e-0f0d-49d5-bcab-7a4d05742037';
const NAVIGATE_EVENT = 'spse:dynamicform-wizard-navigate';
const STATE_EVENT = 'spse:dynamicform-state-changed';

export type NavigationPolicy = 'free' | 'completed' | 'sequential' | 'preview';
export type NavigationLayout = 'vertical' | 'horizontal' | 'tree';

export interface ICategoryConfig { id?: string; title: string; order?: number; collapsed?: boolean; }
export interface IFormConfig { instanceId: string; label: string; categoryId: string; order?: number; required?: boolean; policy?: NavigationPolicy; }
interface IFormState { instanceId: string; instanceName?: string; listName?: string; active?: boolean; itemId?: number; mode?: string; dirty?: boolean; completed?: boolean; isSubmitting?: boolean; hasErrors?: boolean; }
export interface IFormNavigatorWebPartProps {
  title?: string;
  categories?: ICategoryConfig[];
  forms?: IFormConfig[];
  categoriesJson?: string;
  formsJson?: string;
  navigationPolicy?: NavigationPolicy;
  layout?: NavigationLayout;
  showStatus?: boolean;
  showCategoryCounts?: boolean;
  confirmDirtyNavigation?: boolean;
  surfaceBackgroundColor?: string;
  surfaceTextColor?: string;
  categoryBackgroundColor?: string;
  categoryTextColor?: string;
  nodeBackgroundColor?: string;
  nodeTextColor?: string;
  activeBackgroundColor?: string;
  activeTextColor?: string;
  completeColor?: string;
  lockedColor?: string;
  errorColor?: string;
  borderColor?: string;
  borderWidth?: number;
  cornerRadius?: number;
  fontFamily?: string;
  fontSize?: number;
  categoryFontSize?: number;
  fontBold?: boolean;
  categoryFontBold?: boolean;
  nodePadding?: number;
  nodeGap?: number;
  categoryGap?: number;
  indentSize?: number;
  enableDiagnostics?: boolean;
}

export default class FormNavigatorWebPart extends BaseClientSideWebPart<IFormNavigatorWebPartProps> {
  private _states: { [instanceId: string]: IFormState } = {};
  private _collapsed: { [categoryId: string]: boolean } = {};
  private _stateHandler: (event: Event) => void;
  private _discoveredForms: { key: string; text: string }[] = [];
  private _configurationError: string = '';
  private _initializedConfiguration: string = '';
  private _initializationAttempts: number = 0;
  private _initializationTimer: number = 0;

  public constructor() {
    super();
    this._stateHandler = (event: Event) => this.handleStateChanged(event);
  }

  protected onInit(): Promise<void> {
    if (typeof window !== 'undefined') { window.addEventListener(STATE_EVENT, this._stateHandler); }
    this.migrateLegacyConfiguration();
    this.discoverForms();
    return Promise.resolve();
  }

  protected onDispose(): void {
    if (typeof window !== 'undefined') { window.removeEventListener(STATE_EVENT, this._stateHandler); }
    if (this._initializationTimer) { window.clearTimeout(this._initializationTimer); }
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) { return; }
    console.log('[FormNavigatorWebPart] ' + message);
  }

  private getWebPartVersion(): string {
    const solutionVersion = packageSolutionConfig && packageSolutionConfig.solution ? String(packageSolutionConfig.solution.version || '') : '';
    return solutionVersion || (this.context.manifest ? String(this.context.manifest.version || 'Unknown') : 'Unknown');
  }

  private normalizeId(value: string): string { return String(value || '').replace(/[{}]/g, '').toLowerCase(); }

  private migrateLegacyConfiguration(): void {
    if (!Array.isArray(this.properties.categories) && this.properties.categoriesJson) {
      try { this.properties.categories = JSON.parse(this.properties.categoriesJson); } catch (_categoryError) { this.properties.categories = []; }
    }
    if (!Array.isArray(this.properties.forms) && this.properties.formsJson) {
      try { this.properties.forms = JSON.parse(this.properties.formsJson); } catch (_formError) { this.properties.forms = []; }
    }
    if (!Array.isArray(this.properties.categories)) { this.properties.categories = []; }
    if (!Array.isArray(this.properties.forms)) { this.properties.forms = []; }
  }

  private discoverForms(): void {
    var provider: any = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
    var discovered: { key: string; text: string }[] = [];
    if (provider && provider.getAvailableSources) {
      var sources: any[] = provider.getAvailableSources() || [];
      for (var index = 0; index < sources.length; index += 1) {
        var source = sources[index];
        var metadata = source && source.metadata || {};
        var componentId = this.normalizeId(String(metadata.componentId || ''));
        var alias = String(metadata.alias || '').toLowerCase();
        var isDynamicForm = componentId === DYNAMIC_FORM_COMPONENT_ID || alias === 'sharepointdynamicformwebpart';
        var isReportForm = componentId === REPORT_FORM_COMPONENT_ID || alias === 'reportformswebpart';
        if (!isDynamicForm && !isReportForm) { continue; }
        var instanceId = String(metadata.instanceId || source.id || '').replace(/[{}]/g, '');
        var instanceName = '';
        var navigationState: any = undefined;
        try { instanceName = source.getPropertyValue ? String(source.getPropertyValue('instanceName') || '') : ''; } catch (_nameError) { instanceName = ''; }
        try { navigationState = source.getPropertyValue ? source.getPropertyValue('navigationState') : undefined; } catch (_stateError) { navigationState = undefined; }
        if (instanceId) {
          discovered.push({ key: this.normalizeId(instanceId), text: (instanceName || String(metadata.title || (isReportForm ? 'Report Form' : 'Dynamic Form'))) + ' [' + instanceId + ']' });
          if (navigationState) { this._states[this.normalizeId(instanceId)] = navigationState; }
        }
      }
    }
    this._discoveredForms = discovered;
    this.logDiagnostic('Discovered ' + String(discovered.length) + ' form/report source(s).');
  }

  private parseConfiguration(): { categories: ICategoryConfig[]; forms: IFormConfig[] } | undefined {
    try {
      var categories = (this.properties.categories || []).slice(0);
      var forms = (this.properties.forms || []).slice(0);
      if (!Array.isArray(categories)) { throw new Error(strings.InvalidCategories); }
      if (!Array.isArray(forms)) { throw new Error(strings.InvalidForms); }
      var categoryIds: { [id: string]: boolean } = {};
      for (var categoryIndex = 0; categoryIndex < categories.length; categoryIndex += 1) {
        var category = categories[categoryIndex];
        var categoryId = String(category && (category.id || category.title) || '').trim();
        if (!categoryId || !String(category.title || '').trim() || categoryIds[categoryId]) { throw new Error(strings.InvalidCategories); }
        category.id = categoryId;
        category.order = categoryIndex;
        categoryIds[categoryId] = true;
        if (this._collapsed[categoryId] === undefined) { this._collapsed[categoryId] = category.collapsed === true; }
      }
      var formIds: { [id: string]: boolean } = {};
      for (var formIndex = 0; formIndex < forms.length; formIndex += 1) {
        var form = forms[formIndex];
        var formId = this.normalizeId(form && form.instanceId || '');
        if (!formId || !String(form.label || '').trim() || formIds[formId] || !categoryIds[String(form.categoryId || '')]) { throw new Error(strings.InvalidForms); }
        form.instanceId = formId;
        form.order = formIndex;
        formIds[formId] = true;
      }
      this._configurationError = '';
      return { categories: categories, forms: forms };
    } catch (error) {
      this._configurationError = error && error.message ? error.message : strings.InvalidForms;
      return undefined;
    }
  }

  private handleStateChanged(event: Event): void {
    var state = (event as any).detail as IFormState;
    var instanceId = this.normalizeId(state && state.instanceId || '');
    if (!instanceId) { return; }
    this._states[instanceId] = state;
    this.render();
  }

  private getPolicy(form: IFormConfig): NavigationPolicy { return form.policy || this.properties.navigationPolicy || 'sequential'; }

  private isSequentiallyUnlocked(form: IFormConfig, forms: IFormConfig[]): boolean {
    var formIndex = forms.indexOf(form);
    for (var index = 0; index < formIndex; index += 1) {
      var predecessor = forms[index];
      if (predecessor.required === false) { continue; }
      var predecessorState: IFormState = this._states[this.normalizeId(predecessor.instanceId)] || {} as IFormState;
      if (!predecessorState.completed) { return false; }
    }
    return true;
  }

  private isLocked(form: IFormConfig, forms: IFormConfig[]): boolean {
    var state: IFormState = this._states[this.normalizeId(form.instanceId)] || {} as IFormState;
    if (state.active) { return false; }
    var policy = this.getPolicy(form);
    if (policy === 'free') { return false; }
    if (policy === 'completed') { return !state.completed; }
    if (policy === 'preview') { return !this.isSequentiallyUnlocked(form, forms) && !(Number(state.itemId || 0) > 0); }
    return !this.isSequentiallyUnlocked(form, forms);
  }

  private getStatus(form: IFormConfig, forms: IFormConfig[]): { label: string; icon: string; color: string } {
    var state: IFormState = this._states[this.normalizeId(form.instanceId)] || {} as IFormState;
    if (state.hasErrors) { return { label: strings.StatusError, icon: 'ErrorBadge', color: this.properties.errorColor || '#a80000' }; }
    if (state.active) { return { label: strings.StatusActive, icon: 'ChevronRight', color: this.properties.activeTextColor || '#ffffff' }; }
    if (state.completed) { return { label: strings.StatusComplete, icon: 'CompletedSolid', color: this.properties.completeColor || '#107c10' }; }
    if (this.isLocked(form, forms)) { return { label: strings.StatusLocked, icon: 'Lock', color: this.properties.lockedColor || '#8a8886' }; }
    return { label: strings.StatusAvailable, icon: 'CircleRing', color: this.properties.nodeTextColor || '#323130' };
  }

  private navigateTo(form: IFormConfig, forms: IFormConfig[]): void {
    if (this.isLocked(form, forms)) { return; }
    var activeState: IFormState | undefined;
    var stateKey: string;
    for (stateKey in this._states) {
      if (Object.prototype.hasOwnProperty.call(this._states, stateKey) && this._states[stateKey].active) { activeState = this._states[stateKey]; break; }
    }
    if (activeState && activeState.dirty && this.properties.confirmDirtyNavigation !== false && !window.confirm(strings.UnsavedConfirm)) { return; }
    this.activateForm(form, forms, true);
  }

  private activateForm(form: IFormConfig, forms: IFormConfig[], showNotFound: boolean): boolean {
    var targetState: IFormState = this._states[this.normalizeId(form.instanceId)] || {} as IFormState;
    var previewLocked = this.getPolicy(form) === 'preview' && !this.isSequentiallyUnlocked(form, forms);
    var detail: any = {
      targetInstanceId: form.instanceId,
      sourceInstanceId: this.context.instanceId,
      itemId: Number(targetState.itemId || 0),
      openSavedItem: previewLocked || Number(targetState.itemId || 0) > 0,
      navigatorNavigation: true,
      managedInstanceIds: forms.map((configuredForm: IFormConfig) => this.normalizeId(configuredForm.instanceId)),
      handled: false
    };
    var event: any;
    if (typeof CustomEvent === 'function') { event = new CustomEvent(NAVIGATE_EVENT, { detail: detail }); }
    else { event = document.createEvent('CustomEvent'); event.initCustomEvent(NAVIGATE_EVENT, false, false, detail); }
    window.dispatchEvent(event);
    if (!detail.handled && showNotFound) { window.alert(strings.TargetNotFound); }
    return detail.handled === true;
  }

  private initializeConfiguredNavigation(forms: IFormConfig[]): void {
    if (this.displayMode === DisplayMode.Edit || !forms.length) { return; }
    var signature = forms.map((form: IFormConfig) => this.normalizeId(form.instanceId)).join('|');
    if (signature === this._initializedConfiguration) { return; }
    this._initializedConfiguration = signature;
    var initialForm = forms[0];
    for (var index = 0; index < forms.length; index += 1) {
      if ((this._states[this.normalizeId(forms[index].instanceId)] || {} as IFormState).active) {
        initialForm = forms[index];
        break;
      }
    }
    if (this.activateForm(initialForm, forms, false)) {
      this._initializationAttempts = 0;
      return;
    }
    this._initializedConfiguration = '';
    if (this._initializationAttempts < 10) {
      this._initializationAttempts += 1;
      this._initializationTimer = window.setTimeout(() => this.render(), 100);
    }
  }

  private escapeHtml(value: string): string {
    return String(value || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  public render(): void {
    this.discoverForms();
    var config = this.parseConfiguration();
    var rootStyle = 'background:' + (this.properties.surfaceBackgroundColor || '#ffffff') + ';color:' + (this.properties.surfaceTextColor || '#201f1e')
      + ';font-family:' + (this.properties.fontFamily || 'Segoe UI') + ';font-size:' + String(this.properties.fontSize || 14) + 'px;'
      + 'border:' + String(this.properties.borderWidth === undefined ? 1 : this.properties.borderWidth) + 'px solid ' + (this.properties.borderColor || '#d2d0ce') + ';'
      + 'border-radius:' + String(this.properties.cornerRadius || 0) + 'px;padding:12px;';
    if (!config) { this.domElement.innerHTML = '<div class="fn-root fn-error" style="' + rootStyle + '">' + this.escapeHtml(this._configurationError) + '</div>'; return; }
    if (!config.categories.length || !config.forms.length) { this.domElement.innerHTML = '<div class="fn-root" style="' + rootStyle + '"><div class="fn-empty">' + this.escapeHtml(strings.NoConfiguration) + '</div></div>'; return; }

    var categoryHtml: string[] = [];
    for (var categoryIndex = 0; categoryIndex < config.categories.length; categoryIndex += 1) {
      var category = config.categories[categoryIndex];
      var categoryForms = config.forms.filter(function(form) { return form.categoryId === category.id; });
      var completeCount = 0;
      for (var countIndex = 0; countIndex < categoryForms.length; countIndex += 1) {
        var countState: IFormState = this._states[this.normalizeId(categoryForms[countIndex].instanceId)] || {} as IFormState;
        if (countState.completed) { completeCount += 1; }
      }
      var collapsed = this._collapsed[category.id] === true;
      var nodes: string[] = [];
      for (var formIndex = 0; formIndex < categoryForms.length; formIndex += 1) {
        var form = categoryForms[formIndex];
        var state: IFormState = this._states[this.normalizeId(form.instanceId)] || {} as IFormState;
        var locked = this.isLocked(form, config.forms);
        var status = this.getStatus(form, config.forms);
        var nodeStyle = 'background:' + (state.active ? (this.properties.activeBackgroundColor || '#0078d4') : (this.properties.nodeBackgroundColor || '#ffffff'))
          + ';color:' + (state.active ? (this.properties.activeTextColor || '#ffffff') : (this.properties.nodeTextColor || '#323130'))
          + ';border:' + String(this.properties.borderWidth === undefined ? 1 : this.properties.borderWidth) + 'px solid ' + (this.properties.borderColor || '#d2d0ce')
          + ';border-radius:' + String(this.properties.cornerRadius || 0) + 'px;padding:' + String(this.properties.nodePadding || 10) + 'px;'
          + 'font-weight:' + (this.properties.fontBold === true ? '700' : '400') + ';';
        if (this.properties.layout === 'tree') { nodeStyle += 'margin-left:' + String(this.properties.indentSize || 18) + 'px;width:calc(100% - ' + String(this.properties.indentSize || 18) + 'px);'; }
        nodes.push('<button type="button" class="fn-node" data-form-id="' + this.escapeHtml(form.instanceId) + '" style="' + nodeStyle + '"' + (locked ? ' disabled="disabled"' : '') + ' aria-current="' + (state.active ? 'step' : 'false') + '">'
          + (this.properties.showStatus === false ? '' : '<span class="fn-status" style="color:' + status.color + '"><i class="ms-Icon ms-Icon--' + status.icon + '" aria-hidden="true"></i><span class="fn-visually-hidden">' + this.escapeHtml(status.label) + '</span></span>')
          + '<span class="fn-node-label">' + this.escapeHtml(form.label) + '</span></button>');
      }
      var categoryStyle = 'background:' + (this.properties.categoryBackgroundColor || '#f3f2f1') + ';color:' + (this.properties.categoryTextColor || '#323130')
        + ';font-size:' + String(this.properties.categoryFontSize || 14) + 'px;font-weight:' + (this.properties.categoryFontBold === false ? '400' : '700')
        + ';padding:' + String(this.properties.nodePadding || 10) + 'px;border-radius:' + String(this.properties.cornerRadius || 0) + 'px;';
      categoryHtml.push('<section class="fn-category"><button type="button" class="fn-category-button" data-category-id="' + this.escapeHtml(category.id) + '" style="' + categoryStyle + '" aria-expanded="' + (!collapsed) + '">'
        + '<i class="ms-Icon ms-Icon--' + (collapsed ? 'ChevronRight' : 'ChevronDown') + '" aria-hidden="true"></i><span class="fn-category-title">' + this.escapeHtml(category.title) + '</span>'
        + (this.properties.showCategoryCounts === false ? '' : '<span class="fn-category-count">' + String(completeCount) + '/' + String(categoryForms.length) + '</span>') + '</button>'
        + (collapsed ? '' : '<div class="fn-node-list" style="gap:' + String(this.properties.nodeGap || 0) + 'px;margin-top:' + String(this.properties.nodeGap || 0) + 'px">' + (nodes.length ? nodes.join('') : '<div class="fn-empty">' + this.escapeHtml(strings.NoFormsInCategory) + '</div>') + '</div>') + '</section>');
    }
    this.domElement.innerHTML = '<nav class="fn-root fn-layout-' + (this.properties.layout || 'vertical') + '" aria-label="' + this.escapeHtml(this.properties.title || strings.TitleLabel) + '" style="' + rootStyle + '">'
      + (this.properties.title ? '<h2 class="fn-title">' + this.escapeHtml(this.properties.title) + '</h2>' : '')
      + '<div class="fn-categories" style="gap:' + String(this.properties.categoryGap || 0) + 'px">' + categoryHtml.join('') + '</div></nav>';
    var categoryButtons = this.domElement.querySelectorAll('[data-category-id]');
    for (var buttonIndex = 0; buttonIndex < categoryButtons.length; buttonIndex += 1) {
      (categoryButtons[buttonIndex] as HTMLButtonElement).onclick = (ev: MouseEvent) => {
        var categoryId = String((ev.currentTarget as HTMLElement).getAttribute('data-category-id') || '');
        this._collapsed[categoryId] = !this._collapsed[categoryId];
        this.render();
      };
    }
    var formButtons = this.domElement.querySelectorAll('[data-form-id]');
    for (var formButtonIndex = 0; formButtonIndex < formButtons.length; formButtonIndex += 1) {
      (formButtons[formButtonIndex] as HTMLButtonElement).onclick = (ev: MouseEvent) => {
        var formId = this.normalizeId(String((ev.currentTarget as HTMLElement).getAttribute('data-form-id') || ''));
        for (var configuredIndex = 0; configuredIndex < config.forms.length; configuredIndex += 1) {
          if (this.normalizeId(config.forms[configuredIndex].instanceId) === formId) { this.navigateTo(config.forms[configuredIndex], config.forms); break; }
        }
      };
    }
    this.initializeConfiguredNavigation(config.forms);
  }

  protected onPropertyPaneConfigurationStart(): void { this.discoverForms(); this.context.propertyPane.refresh(); }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'categories' || propertyPath === 'forms') {
      this._initializedConfiguration = '';
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  private colorField(targetProperty: string, label: string): any {
    return PropertyFieldColorPicker(targetProperty, { label: label, selectedColor: (this.properties as any)[targetProperty], onPropertyChange: this.onPropertyPaneFieldChanged.bind(this), properties: this.properties, disabled: false, alphaSliderHidden: true, style: PropertyFieldColorPickerStyle.Inline, key: targetProperty });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    var categoryOptions = (this.properties.categories || []).filter(function(category: ICategoryConfig) {
      return !!String(category && category.title || '').trim();
    }).map(function(category: ICategoryConfig) {
      var categoryId = String(category.id || category.title).trim();
      return { key: categoryId, text: String(category.title).trim() };
    });
    return { pages: [{ header: { description: strings.PropertyPaneDescription }, groups: [
      { groupName: strings.VersionGroup + ': ' + this.getWebPartVersion(), groupFields: [PropertyPaneLabel('versionInfo', { text: ' ' })] },
      { groupName: strings.DataGroup, groupFields: [
        PropertyPaneTextField('title', { label: strings.TitleLabel }),
        PropertyFieldCollectionData('categories', {
          key: 'formNavigatorCategories', label: strings.CategoriesLabel, panelHeader: strings.CategoriesHeader,
          panelDescription: strings.CategoriesDescription, manageBtnLabel: strings.ManageCategoriesLabel,
          value: this.properties.categories || [], enableSorting: true,
          fields: [
            { id: 'title', title: strings.CategoryNameLabel, type: CustomCollectionFieldType.string, required: true },
            { id: 'collapsed', title: strings.CategoryCollapsedLabel, type: CustomCollectionFieldType.boolean, defaultValue: false }
          ]
        }),
        PropertyFieldCollectionData('forms', {
          key: 'formNavigatorForms', label: strings.FormsLabel, panelHeader: strings.FormsHeader,
          panelDescription: strings.FormsDescription, manageBtnLabel: strings.ManageFormsLabel,
          value: this.properties.forms || [], enableSorting: true,
          fields: [
            { id: 'categoryId', title: strings.FormCategoryLabel, type: CustomCollectionFieldType.dropdown, required: true, options: categoryOptions },
            { id: 'instanceId', title: strings.FormSelectionLabel, type: CustomCollectionFieldType.dropdown, required: true, options: this._discoveredForms },
            { id: 'label', title: strings.FormLabelLabel, type: CustomCollectionFieldType.string, required: true },
            { id: 'required', title: strings.FormRequiredLabel, type: CustomCollectionFieldType.boolean, defaultValue: true },
            { id: 'policy', title: strings.FormPolicyLabel, type: CustomCollectionFieldType.dropdown, options: [{ key: '', text: strings.PolicyDefault }, { key: 'free', text: strings.PolicyFree }, { key: 'completed', text: strings.PolicyCompleted }, { key: 'sequential', text: strings.PolicySequential }, { key: 'preview', text: strings.PolicyPreview }] }
          ]
        })
      ] },
      { groupName: strings.BehaviorGroup, groupFields: [
        PropertyPaneDropdown('navigationPolicy', { label: strings.NavigationPolicyLabel, options: [{ key: 'free', text: strings.PolicyFree }, { key: 'completed', text: strings.PolicyCompleted }, { key: 'sequential', text: strings.PolicySequential }, { key: 'preview', text: strings.PolicyPreview }] }),
        PropertyPaneDropdown('layout', { label: strings.LayoutLabel, options: [{ key: 'vertical', text: strings.LayoutVertical }, { key: 'horizontal', text: strings.LayoutHorizontal }, { key: 'tree', text: strings.LayoutTree }] }),
        PropertyPaneCheckbox('showStatus', { text: strings.ShowStatusLabel, checked: this.properties.showStatus !== false }),
        PropertyPaneCheckbox('showCategoryCounts', { text: strings.ShowCategoryCountsLabel, checked: this.properties.showCategoryCounts !== false }),
        PropertyPaneCheckbox('confirmDirtyNavigation', { text: strings.ConfirmDirtyLabel, checked: this.properties.confirmDirtyNavigation !== false })
      ] },
      { groupName: strings.AppearanceGroup, groupFields: [
        this.colorField('surfaceBackgroundColor', strings.SurfaceBackgroundLabel), this.colorField('surfaceTextColor', strings.SurfaceTextLabel), this.colorField('categoryBackgroundColor', strings.CategoryBackgroundLabel), this.colorField('categoryTextColor', strings.CategoryTextLabel), this.colorField('nodeBackgroundColor', strings.NodeBackgroundLabel), this.colorField('nodeTextColor', strings.NodeTextLabel),
        PropertyPaneDropdown('fontFamily', { label: strings.FontFamilyLabel, options: [{ key: 'Segoe UI', text: 'Segoe UI' }, { key: 'Arial', text: 'Arial' }, { key: 'Calibri', text: 'Calibri' }, { key: 'Tahoma', text: 'Tahoma' }, { key: 'Verdana', text: 'Verdana' }, { key: 'Georgia', text: 'Georgia' }] }),
        PropertyPaneSlider('fontSize', { label: strings.FontSizeLabel, min: 10, max: 28, step: 1 }), PropertyPaneSlider('categoryFontSize', { label: strings.CategoryFontSizeLabel, min: 10, max: 32, step: 1 }),
        PropertyPaneCheckbox('fontBold', { text: strings.FontBoldLabel, checked: this.properties.fontBold === true }), PropertyPaneCheckbox('categoryFontBold', { text: strings.CategoryFontBoldLabel, checked: this.properties.categoryFontBold !== false })
      ] },
      { groupName: strings.StateAppearanceGroup, groupFields: [this.colorField('activeBackgroundColor', strings.ActiveBackgroundLabel), this.colorField('activeTextColor', strings.ActiveTextLabel), this.colorField('completeColor', strings.CompleteColorLabel), this.colorField('lockedColor', strings.LockedColorLabel), this.colorField('errorColor', strings.ErrorColorLabel)] },
      { groupName: strings.SpacingGroup, groupFields: [this.colorField('borderColor', strings.BorderColorLabel), PropertyPaneSlider('borderWidth', { label: strings.BorderWidthLabel, min: 0, max: 10, step: 1 }), PropertyPaneSlider('cornerRadius', { label: strings.CornerRadiusLabel, min: 0, max: 30, step: 1 }), PropertyPaneSlider('nodePadding', { label: strings.NodePaddingLabel, min: 2, max: 30, step: 1 }), PropertyPaneSlider('nodeGap', { label: strings.NodeGapLabel, min: 0, max: 30, step: 1 }), PropertyPaneSlider('categoryGap', { label: strings.CategoryGapLabel, min: 0, max: 50, step: 1 }), PropertyPaneSlider('indentSize', { label: strings.IndentSizeLabel, min: 0, max: 60, step: 1 })] },
      { groupName: strings.DiagnosticsGroup, groupFields: [PropertyPaneCheckbox('enableDiagnostics', { text: strings.DiagnosticsLabel, checked: this.properties.enableDiagnostics !== false })] }
    ] }] };
  }
}
