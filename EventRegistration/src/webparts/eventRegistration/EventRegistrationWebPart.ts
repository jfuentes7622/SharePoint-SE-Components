import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import EventRegistration from './components/EventRegistration';
import { IEventRegistrationProps } from './components/IEventRegistrationProps';
import { IRegistrationSelection } from './models';
import EventRegistrationService from './services/EventRegistrationService';
import { applyAppearance, getAppearanceMode, IEventRegistrationAppearance, IEventRegistrationTheme } from './helpers/appearance';
import { getDisplayMode } from './helpers/displayMode';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface IEventRegistrationWebPartProps extends IEventRegistrationAppearance {
  title: string;
  description: string;
  displayMode?: string;
  eventId?: string;
  registrationPageUrl?: string;
  excludeFromTabs?: boolean;
  listPrefix: string;
  organizerGroupId: number;
  preferredCalendarInstanceId: string;
  buttonDisplayMode: string;
  enableDiagnostics?: boolean;
}

export interface IDynamicDataPropertyDefinitionCompat {
  id: string;
  title: string;
  description?: string;
}

export default class EventRegistrationWebPart extends BaseClientSideWebPart<IEventRegistrationWebPartProps> {
  private siteTheme: IEventRegistrationTheme;
  private groupOptions: IPropertyPaneDropdownOption[] = [{ key: 0, text: '(Site administrators only)' }];
  private calendarOptions: IPropertyPaneDropdownOption[] = [{ key: '', text: '(No preferred Calendar)' }];
  private dynamicDataProvider: any;
  private dynamicDataSourceManager: any;
  private subscribedCalendarId: string = '';
  private calendarChangedHandler: () => void;
  private availableSourcesChangedHandler: () => void;
  private dynamicEventId: number = 0;
  private dynamicSessionId: number = 0;
  private selectedRegistration: IRegistrationSelection = { id: 0, status: '', confirmation: '' };

  public get id(): string {
    return this.context.instanceId;
  }

  public get metadata(): any {
    return {
      title: 'Event Registration',
      description: 'Selected Event Registration result.',
      alias: this.context.manifest.alias,
      componentId: this.context.manifest.id,
      instanceId: this.context.instanceId
    };
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinitionCompat> {
    return [
      { id: 'selectedRegistrationId', title: 'Selected Registration ID' },
      { id: 'selectedRegistrationStatus', title: 'Selected Registration Status' },
      { id: 'selectedRegistrationConfirmation', title: 'Selected Registration Confirmation' }
    ];
  }

  public getPropertyValue(propertyId: string): any {
    if (propertyId === 'selectedRegistrationId') {
      return this.selectedRegistration.id;
    }
    if (propertyId === 'selectedRegistrationStatus') {
      return this.selectedRegistration.status;
    }
    if (propertyId === 'selectedRegistrationConfirmation') {
      return this.selectedRegistration.confirmation;
    }
    throw new Error('EventRegistrationWebPart: unsupported Dynamic Data property "' + propertyId + '".');
  }

  protected onInit(): Promise<void> {
    this.properties.listPrefix = String(this.properties.listPrefix || 'SPS Event Registration').trim()
      || 'SPS Event Registration';
    this.properties.buttonDisplayMode = this.normalizeButtonMode(this.properties.buttonDisplayMode);
    this.dynamicDataProvider = (this.context as any).dynamicDataProvider || (this.context as any)._dynamicDataProvider;
    this.dynamicDataSourceManager = (this.context as any).dynamicDataSourceManager || (this.context as any)._dynamicDataSourceManager;
    if (this.dynamicDataSourceManager) {
      if (this.dynamicDataSourceManager.initializeSource) {
        this.dynamicDataSourceManager.initializeSource(this);
      } else if (this.dynamicDataSourceManager.registerSource) {
        this.dynamicDataSourceManager.registerSource(this);
      }
    }
    this.registerAvailableSourcesHandler();
    this.refreshCalendarSources();
    this.subscribeCalendar();
    this.logDiagnostic('Initialized with list prefix "' + this.properties.listPrefix + '".');
    return Promise.resolve();
  }

  public render(): void {
    this.applyAppearance();
    this.domElement.setAttribute('data-spse-exclude-from-tabs', String(this.properties.excludeFromTabs === true));
    var query = this.readQuerySelection();
    var element = React.createElement<IEventRegistrationProps>(EventRegistration, {
      title: this.properties.title === undefined ? 'Event Registration' : this.properties.title,
      description: this.properties.description || '',
      displayMode: getDisplayMode(this.properties.displayMode),
      registrationPageUrl: this.properties.registrationPageUrl || '',
      webUrl: this.context.pageContext.web.absoluteUrl,
      spHttpClient: this.context.spHttpClient,
      listPrefix: this.properties.listPrefix,
      organizerGroupId: Number(this.properties.organizerGroupId || 0),
      buttonDisplayMode: this.normalizeButtonMode(this.properties.buttonDisplayMode),
      enableDiagnostics: this.properties.enableDiagnostics !== false,
      initialEventId: query.eventId || this.dynamicEventId || (getDisplayMode(this.properties.displayMode) === 'registration'
        ? this.positiveNumber(String(this.properties.eventId || '')) : 0),
      initialSessionId: query.sessionId || this.dynamicSessionId,
      sourceUrl: query.sourceUrl,
      onRegistrationSelected: this.handleRegistrationSelected.bind(this)
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    this.unsubscribeCalendar();
    if (this.dynamicDataProvider && this.availableSourcesChangedHandler
      && this.dynamicDataProvider.unregisterAvailableSourcesChanged) {
      this.dynamicDataProvider.unregisterAvailableSourcesChanged(this.availableSourcesChangedHandler);
    }
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.loadGroupOptions();
    this.refreshCalendarSources();
    this.context.propertyPane.refresh();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'preferredCalendarInstanceId') {
      this.subscribeCalendar();
    }
    if (propertyPath === 'appearanceMode' || propertyPath === 'displayMode') {
      this.context.propertyPane.refresh();
    }
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    var strictTheme = this.getAppearanceMode() === 'siteTheme';
    return {
      pages: [{
        header: { description: '' },
        groups: [
          {
            groupName: 'Version: ' + this.getWebPartVersion(),
            groupFields: [PropertyPaneLabel('propertyPaneVersionInfo', { text: ' ' })]
          },
          {
            groupName: 'Configuration',
            groupFields: [
              PropertyPaneTextField('title', { label: 'Title' }),
              PropertyPaneTextField('description', { label: 'Description', multiline: true }),
              PropertyPaneDropdown('displayMode', {
                label: 'Display mode',
                selectedKey: getDisplayMode(this.properties.displayMode),
                options: [
                  { key: 'combined', text: 'Combined (existing experience)' },
                  { key: 'catalog', text: 'Event catalog' },
                  { key: 'registration', text: 'Single event registration' },
                  { key: 'organizer', text: 'Organizer management' }
                ]
              }),
              PropertyPaneTextField('eventId', {
                label: 'Default event ID',
                description: 'Single-event page fallback. URL and Calendar selections take precedence.',
                disabled: getDisplayMode(this.properties.displayMode) !== 'registration',
                onGetErrorMessage: (value: string): string => {
                  return !value || /^[1-9][0-9]*$/.test(value) ? '' : 'Enter a positive whole-number event ID.';
                }
              }),
              PropertyPaneTextField('registrationPageUrl', {
                label: 'Registration page URL (optional)',
                description: 'Catalog opens this same-origin page with eventid. Leave blank for inline event details.',
                disabled: getDisplayMode(this.properties.displayMode) !== 'catalog'
              }),
              PropertyPaneTextField('listPrefix', {
                label: 'List name prefix',
                description: 'Default: SPS Event Registration'
              }),
              PropertyPaneDropdown('organizerGroupId', {
                label: 'Organizer SharePoint group',
                options: this.groupOptions,
                selectedKey: Number(this.properties.organizerGroupId || 0)
              }),
              PropertyPaneDropdown('preferredCalendarInstanceId', {
                label: 'Preferred Calendar instance',
                options: this.calendarOptions,
                selectedKey: this.properties.preferredCalendarInstanceId || ''
              })
            ]
          },
          {
            groupName: 'Buttons',
            groupFields: [
              PropertyPaneDropdown('buttonDisplayMode', {
                label: 'Button display',
                options: [
                  { key: 'icon', text: 'Icons' },
                  { key: 'iconText', text: 'Icons and names' },
                  { key: 'text', text: 'Names' }
                ],
                selectedKey: this.normalizeButtonMode(this.properties.buttonDisplayMode)
              })
            ]
          },
          {
            groupName: 'Appearance',
            groupFields: [
              PropertyPaneDropdown('appearanceMode', {
                label: 'Appearance mode',
                options: [
                  { key: 'manual', text: 'Manual styling' },
                  { key: 'siteTheme', text: 'Site theme only' },
                  { key: 'siteThemeOverrides', text: 'Site theme with overrides' }
                ],
                selectedKey: this.getAppearanceMode()
              }),
              PropertyPaneLabel('siteThemeAppearanceHelp', {
                text: strictTheme
                  ? 'Theme-controlled colors and font family are disabled while Site theme only is selected.'
                  : 'Choose how configurable colors and font family use the SharePoint site theme.'
              }),
              PropertyFieldColorPicker('surfaceBackgroundColor', {
                label: 'Background color',
                selectedColor: this.properties.surfaceBackgroundColor || '#ffffff',
                onPropertyChange: this.handleColorPropertyChange.bind(this),
                properties: this.properties,
                style: PropertyFieldColorPickerStyle.Inline,
                key: 'surfaceBackgroundColorField',
                disabled: strictTheme
              }),
              PropertyFieldColorPicker('surfaceTextColor', {
                label: 'Text color',
                selectedColor: this.properties.surfaceTextColor || '#323130',
                onPropertyChange: this.handleColorPropertyChange.bind(this),
                properties: this.properties,
                style: PropertyFieldColorPickerStyle.Inline,
                key: 'surfaceTextColorField',
                disabled: strictTheme
              }),
              PropertyFieldColorPicker('accentColor', {
                label: 'Accent color',
                selectedColor: this.properties.accentColor || '#0078d4',
                onPropertyChange: this.handleColorPropertyChange.bind(this),
                properties: this.properties,
                style: PropertyFieldColorPickerStyle.Inline,
                key: 'accentColorField',
                disabled: strictTheme
              }),
              PropertyFieldColorPicker('buttonTextColor', {
                label: 'Button text color',
                selectedColor: this.properties.buttonTextColor || '#ffffff',
                onPropertyChange: this.handleColorPropertyChange.bind(this),
                properties: this.properties,
                style: PropertyFieldColorPickerStyle.Inline,
                key: 'buttonTextColorField',
                disabled: strictTheme
              }),
              PropertyFieldColorPicker('activeButtonColor', {
                label: 'Active button background color',
                selectedColor: this.properties.activeButtonColor || '#004578',
                onPropertyChange: this.handleColorPropertyChange.bind(this),
                properties: this.properties,
                style: PropertyFieldColorPickerStyle.Inline,
                key: 'activeButtonColorField',
                disabled: strictTheme
              }),
              PropertyFieldColorPicker('borderColor', {
                label: 'Border color',
                selectedColor: this.properties.borderColor || '#c8c6c4',
                onPropertyChange: this.handleColorPropertyChange.bind(this),
                properties: this.properties,
                style: PropertyFieldColorPickerStyle.Inline,
                key: 'borderColorField',
                disabled: strictTheme
              }),
              PropertyPaneSlider('borderWidth', {
                label: 'Border width', min: 0, max: 8, step: 1,
                value: typeof this.properties.borderWidth === 'number' ? this.properties.borderWidth : 0
              }),
              PropertyPaneSlider('cornerRadius', {
                label: 'Corner radius', min: 0, max: 24, step: 1,
                value: typeof this.properties.cornerRadius === 'number' ? this.properties.cornerRadius : 0
              }),
              PropertyPaneTextField('fontFamily', { label: 'Font family', disabled: strictTheme }),
              PropertyPaneSlider('fontSize', {
                label: 'Font size', min: 10, max: 30, step: 1,
                value: typeof this.properties.fontSize === 'number' ? this.properties.fontSize : 14
              })
            ]
          },
          {
            groupName: 'Advanced',
            groupFields: [
              PropertyPaneCheckbox('excludeFromTabs', {
                text: 'Exclude this web part from SPS Tabs',
                checked: this.properties.excludeFromTabs === true
              })
            ]
          },
          {
            groupName: 'Diagnostics',
            groupFields: [
              PropertyPaneCheckbox('enableDiagnostics', {
                text: 'Enable diagnostics logging',
                checked: this.properties.enableDiagnostics !== false
              })
            ]
          }
        ]
      }]
    };
  }

  private loadGroupOptions(): void {
    var service = new EventRegistrationService(this.context.pageContext.web.absoluteUrl, this.context.spHttpClient,
      this.properties.listPrefix, this.properties.enableDiagnostics !== false);
    service.getSiteGroups().then((options: Array<{ key: number; text: string }>): void => {
      this.groupOptions = [{ key: 0, text: '(Site administrators only)' }].concat(options);
      this.context.propertyPane.refresh();
    }).catch((error: any): void => {
      console.error('[EventRegistrationWebPart] Unable to load SharePoint groups: ' + this.errorMessage(error));
    });
  }

  private refreshCalendarSources(): void {
    this.calendarOptions = [{ key: '', text: '(No preferred Calendar)' }];
    if (!this.dynamicDataProvider || !this.dynamicDataProvider.getAvailableSources) {
      return;
    }
    var sources = this.dynamicDataProvider.getAvailableSources() || [];
    for (var index = 0; index < sources.length; index += 1) {
      var source = sources[index];
      if (source && source.metadata && String(source.metadata.alias || '') === 'CalendarWebPart') {
        this.calendarOptions.push({
          key: String(source.id),
          text: String(source.metadata.title || 'Calendar') + ' (' + String(source.id) + ')'
        });
      }
    }
  }

  private subscribeCalendar(): void {
    this.unsubscribeCalendar();
    var source = this.getPreferredCalendar();
    if (!source || !this.dynamicDataProvider || !this.dynamicDataProvider.registerSourceChanged) {
      this.dynamicEventId = 0;
      this.dynamicSessionId = 0;
      return;
    }
    this.readCalendarSelection(source);
    this.calendarChangedHandler = (): void => {
      this.readCalendarSelection(source);
      this.render();
    };
    this.dynamicDataProvider.registerSourceChanged(source.id, this.calendarChangedHandler);
    this.subscribedCalendarId = String(source.id);
    this.logDiagnostic('Subscribed to Calendar Dynamic Data source ' + this.subscribedCalendarId + '.');
  }

  private unsubscribeCalendar(): void {
    if (this.subscribedCalendarId && this.calendarChangedHandler && this.dynamicDataProvider
      && this.dynamicDataProvider.unregisterSourceChanged) {
      this.dynamicDataProvider.unregisterSourceChanged(this.subscribedCalendarId, this.calendarChangedHandler);
    }
    this.subscribedCalendarId = '';
    this.calendarChangedHandler = undefined;
  }

  private getPreferredCalendar(): any {
    if (!this.dynamicDataProvider || !this.dynamicDataProvider.getAvailableSources) {
      return undefined;
    }
    var sources = this.dynamicDataProvider.getAvailableSources() || [];
    var preferred = String(this.properties.preferredCalendarInstanceId || '');
    for (var index = 0; index < sources.length; index += 1) {
      var source = sources[index];
      if (source && String(source.id) === preferred && source.metadata
        && String(source.metadata.alias || '') === 'CalendarWebPart') {
        return source;
      }
    }
    return undefined;
  }

  private readCalendarSelection(source: any): void {
    this.dynamicSessionId = this.safeDynamicNumber(source, 'selectedSessionId');
    this.dynamicEventId = this.safeDynamicNumber(source, 'selectedItemId');
    this.logDiagnostic('Calendar selection: event=' + String(this.dynamicEventId)
      + ', session=' + String(this.dynamicSessionId) + '.');
  }

  private safeDynamicNumber(source: any, propertyId: string): number {
    try {
      var value = Number(source.getPropertyValue(propertyId) || 0);
      return isNaN(value) || value < 1 ? 0 : value;
    } catch (error) {
      this.logDiagnostic('Calendar does not expose ' + propertyId + '.');
      return 0;
    }
  }

  private registerAvailableSourcesHandler(): void {
    if (!this.dynamicDataProvider || !this.dynamicDataProvider.registerAvailableSourcesChanged) {
      return;
    }
    this.availableSourcesChangedHandler = (): void => {
      this.refreshCalendarSources();
      this.subscribeCalendar();
      if (this.context.propertyPane) {
        this.context.propertyPane.refresh();
      }
      this.render();
    };
    this.dynamicDataProvider.registerAvailableSourcesChanged(this.availableSourcesChangedHandler);
  }

  private handleRegistrationSelected(selection: IRegistrationSelection): void {
    this.selectedRegistration = selection;
    this.notifyDynamicData('selectedRegistrationId');
    this.notifyDynamicData('selectedRegistrationStatus');
    this.notifyDynamicData('selectedRegistrationConfirmation');
  }

  private notifyDynamicData(propertyId: string): void {
    if (!this.dynamicDataSourceManager) {
      return;
    }
    if (this.dynamicDataSourceManager.notifyPropertyChanged) {
      this.dynamicDataSourceManager.notifyPropertyChanged(propertyId);
    }
    if (this.dynamicDataSourceManager.notifySourceChanged) {
      this.dynamicDataSourceManager.notifySourceChanged();
    } else if (this.dynamicDataSourceManager.notifyDataChanged) {
      this.dynamicDataSourceManager.notifyDataChanged();
    }
  }

  private readQuerySelection(): { eventId: number; sessionId: number; sourceUrl: string } {
    var values: { [key: string]: string } = {};
    if (typeof window !== 'undefined') {
      var query = String(window.location.search || '').replace(/^\?/, '').split('&');
      query.forEach(function(part: string): void {
        var pair = part.split('=');
        if (pair[0]) {
          values[decodeURIComponent(pair[0]).toLowerCase()] = decodeURIComponent((pair.slice(1).join('=') || '').replace(/\+/g, ' '));
        }
      });
    }
    return {
      eventId: this.positiveNumber(values.eventid),
      sessionId: this.positiveNumber(values.sessionid),
      sourceUrl: values.source || values.returnurl || ''
    };
  }

  private positiveNumber(value: string): number {
    var number = Number(value || 0);
    return isNaN(number) || number < 1 ? 0 : Math.floor(number);
  }

  private normalizeButtonMode(value: string): string {
    return value === 'icon' || value === 'iconText' ? value : 'text';
  }

  protected onThemeChanged(currentTheme: any): void {
    this.siteTheme = currentTheme;
    // SharePoint can supply the initial theme before properties and onInit are ready.
    if (!this.renderedOnce) {
      return;
    }
    this.render();
    if (this.context && this.context.propertyPane) {
      this.context.propertyPane.refresh();
    }
  }

  private getAppearanceMode(): string {
    return getAppearanceMode(this.properties.appearanceMode);
  }

  private applyAppearance(): void {
    applyAppearance(this.domElement.style, this.properties, this.siteTheme);
  }

  private handleColorPropertyChange(propertyPath: string, oldValue: string, newValue: string): void {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private getWebPartVersion(): string {
    var solution = packageSolutionConfig && packageSolutionConfig.solution;
    var version = solution ? String(solution.version || '') : '';
    return version || String(this.context.manifest.version || 'Unknown');
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }
    console.log('[EventRegistrationWebPart] ' + message);
  }

  private errorMessage(error: any): string {
    return error && error.message ? String(error.message) : String(error);
  }
}
