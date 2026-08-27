/*
  Sharepoint will never render list-items if the webpart has a 'display: none' CSS style at load-time.  To get around this,
  while the component is loading, we identify our webpart-zones and set an attribute on each one which allows the DISPLAY property
  to be other than NONE and set the VISIBILITY property to hidden.  Once the page has loaded, we have to remove the attribute
  and then the normal 'display: none' can take over... it's a necessary & ugly hack, but it works reliably
*/

import * as React from 'react';
import * as ReactDom from 'react-dom';

import { DisplayMode,Version } from '@microsoft/sp-core-library';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneLabel
} from '@microsoft/sp-webpart-base';

import { BaseClientSideWebPart,  WebPartContext } from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import * as strings from 'TabComponentWebPartStrings';
import { releaseOptionalFullWidth, updateResponsiveOptionalFullWidth } from '../shared/deterministicFullWidth';
import TabComponent from './components/TabComponent';
import { ITabComponentProps, ITabVisualSettings } from './components/ITabComponentProps';
import ErrorComponent, { IErrorComponentProps } from './components/ErrorComponent';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

require('../tabComponent/assets/TabStyles-round.css');

const packageSolutionConfig: any = require('../../../config/package-solution.json');
const picturePickerModule: any = require('sp-client-custom-fields/lib/PropertyFieldPicturePickerHost');
const PropertyFieldPicturePickerHost: any = picturePickerModule.default || picturePickerModule;

export interface IConfigListData {
  collectionData: any[];
}

export interface ITabComponentWebPartProps {
  forceFullWidth?: boolean;
  description: string;
  disableColor: string;
  overrideCSS: string;
  useGlobalCSS: boolean;
  selectedColor: string;
  TabType: string;
  tabShape: 'rounded' | 'square';
  fontFamily: 'Segoe UI' | 'Arial' | 'Calibri' | 'Verdana' | 'Tahoma' | 'Trebuchet MS' | 'Georgia' | 'Times New Roman' | 'Courier New';
  fontStyle: 'normal' | 'italic' | 'oblique';
  isBold: boolean;
  fontSize: number | string;
  tabHeight: number | string;
  tabWidth: number | string;
  autoTabWidth: boolean;
  tabTextColor: string;
  tabBorderColor: string;
  tabBorderWidth: number | string;
  tabBorderStyle: 'none' | 'solid' | 'dashed' | 'dotted' | 'double';
  tabCornerRadius: number | string;
  tabLineColor: string;
  tabLineWidth: number | string;
  webPartBorderColor: string;
  webPartBorderWidth: number | string;
  webPartBorderStyle: 'none' | 'solid' | 'dashed' | 'dotted' | 'double';
  webPartCornerStyle: 'square' | 'rounded';
  webPartCornerRadius: number | string;
  inactiveImageFade: number | string;
  contentPaddingTop: number | string;
  contentPaddingBottom: number | string;
  collectionData: ITabCollectionItem[];
  tabs: any[];
  enableDiagnostics?: boolean;
}

export interface ITabCollectionItem {
  Title?: string;
  textPosition?: 'left' | 'center' | 'right';
  imageUrl?: string;
  imagePosition?: 'left' | 'right';
  onlyImage?: boolean;
}

export default class TabComponentWebPart extends BaseClientSideWebPart<ITabComponentWebPartProps> {

  private found: boolean = false;
  private foundAnotherTab: boolean = false;
  private tabsProcessedCount = 0;
  private ContentArea: HTMLElement;


  public constructor(context?: WebPartContext) {
    super();
    // Hack: to invoke correctly the onPropertyChange function outside this class
    // we need to bind this object on it first
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
    window.addEventListener('load', this.windowLoaded.bind(this), false);
  }

  private logDiagnostic(message: string): void {
    if (this.properties && this.properties.enableDiagnostics === false) {
      return;
    }

    console.log('[TabComponentWebPart] ' + message);
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

  private windowLoaded(): void {
    this.logDiagnostic('Window load event fired; removing part-Loading attribute from tracked zones.');
    if (!this.ContentArea) {
      return;
    }
    // Remove the attribute from the webparts so the normal CSS rules apply
    Array.from(this.ContentArea.querySelectorAll('[part-Loading]')).forEach(d => {
      d.removeAttribute('part-Loading');
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
    if (propertyPath === 'tabs') {
      this.properties.collectionData = [];
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private handleColorPropertyChange(propertyPath: string, oldValue: any, newValue: any): void {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.context.propertyPane.refresh();
    this.render();
  }

  private renderTabImagePicker(field: any, value: any,
    onUpdate: (fieldId: string, fieldValue: any) => void): React.ReactElement<any> {
    const properties: any = {};
    properties[field.id] = String(value || '');
    return React.createElement(PropertyFieldPicturePickerHost, {
      key: 'tabImagePicker-' + String(value || ''),
      targetProperty: field.id,
      label: '',
      initialValue: String(value || ''),
      context: this.context,
      previewImage: true,
      allowedFileExtensions: '.gif,.jpg,.jpeg,.bmp,.png,.svg,.webp',
      readOnly: false,
      disabled: false,
      properties: properties,
      disableReactivePropertyChanges: true,
      deferredValidationTime: 0,
      onGetErrorMessage: undefined,
      onPropertyChange: (propertyPath: string, oldValue: any, newValue: any): void => {
        onUpdate(field.id, String(newValue || ''));
      },
      render: (): void => undefined,
      onRender: undefined,
      onDispose: undefined
    });
  }

  // Searches through the ancestors of childElem and returns true if targetElem is found
  private findInAncestors(targetElem: Element, childElem: Element): boolean {
    const tabtype = (this.properties.TabType) ? this.properties.TabType : '1';

    if (tabtype === '2') {
      const tabResult: XPathResult = new XPathEvaluator()
        .evaluate('ancestor-or-self::' + targetElem.nodeName, childElem, undefined, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, undefined);

      for (let i: number = 0; i < tabResult.snapshotLength; i++) {
        if (tabResult.snapshotItem(i) === targetElem) { return true; }
      }
      return false;
    }


    const result: XPathResult = new XPathEvaluator()
      .evaluate('ancestor-or-self::' + targetElem.nodeName, childElem, undefined, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, undefined);

    for (let i: number = 0; i < result.snapshotLength; i++) {
      if (result.snapshotItem(i) === targetElem) {
        this.found = true;
        this.tabsProcessedCount += 1;
        return true;
      }
    }

    //return true until found, only sections after being found are counted.
    if (this.found === false) {

      return true;
    } else {
      this.tabsProcessedCount += 1;
    }

    // Only another section-mode Tab ends this section-mode control's range.
    // WebPart-mode Tabs can be nested inside a section controlled by this instance.
    if (targetElem.querySelector('[data-spse-tab-type="1"]') !== null) {
      this.foundAnotherTab = true;
      return true;
    }

    return this.foundAnotherTab;
  }

  public render(): void {
    updateResponsiveOptionalFullWidth(this.domElement, this.context.instanceId, this.properties.forceFullWidth === true);
    // Need to send a message if webpart is used on a classic page
    if (Environment.type === EnvironmentType.ClassicSharePoint) {
      this.logDiagnostic('Classic SharePoint environment detected; rendering unsupported-environment message instead of tabs.');
      const errElem: React.ReactElement<IErrorComponentProps> = React.createElement(
        ErrorComponent, { ErrorStr: strings.ErrorClassicSharePoint } as IErrorComponentProps
      );
      ReactDom.render(errElem, this.domElement);
    }

    let overrideCSS: string="";
    if (this.properties.useGlobalCSS === true) {
      overrideCSS = this.context.pageContext.site.absoluteUrl.replace(this.context.pageContext.site.serverRelativeUrl,"") + "/tabs/css/tabsoverride.css";
    } else if (this.properties.overrideCSS) {
      overrideCSS = this.properties.overrideCSS;
    }

    if (overrideCSS !== "") {
      this.logDiagnostic('Injecting override stylesheet: ' + overrideCSS);
      // inject the EUCOM master style sheet
      const head: HTMLElement = document.getElementsByTagName('head')[0] || document.documentElement;
      const customStyle: HTMLLinkElement = document.createElement('link');
      customStyle.href = overrideCSS;
      customStyle.rel = 'stylesheet';
      customStyle.type = 'text/css';
      head.insertAdjacentElement('beforeend', customStyle);
    }

    this.domElement.style.setProperty('--disableColor', this.properties.disableColor);
    this.domElement.style.setProperty('--selectedColor', this.properties.selectedColor);
    const tabtype = (this.properties.TabType) ? this.properties.TabType : '1';
    this.domElement.setAttribute('data-spse-tab-type', tabtype);
    // Find the root section by traversing the DOM tree up until it finds an element with the CanvasSection class
    this.ContentArea = this.domElement;

      if (tabtype === '1') {
        const canvasRoot = this.domElement.ownerDocument.querySelector('.Canvas, .grid') as HTMLElement;
        this.ContentArea = canvasRoot && canvasRoot.parentElement
          ? canvasRoot.parentElement
          : (this.domElement.parentElement || this.domElement);
      } else {
        let sectionRoot: HTMLElement | null = this.domElement;
        while (sectionRoot && !sectionRoot.classList.contains('CanvasSection')) {
          sectionRoot = sectionRoot.parentElement;
        }
        this.ContentArea = sectionRoot || this.domElement.parentElement || this.domElement;
      }
    //}

    this.found = false;
    this.foundAnotherTab = false;
    this.tabsProcessedCount = 0;
    let ctrlZones: Array<Element>;
    // Get a collection of the controlzones in this section to build tabs for...except this control's zone
    if (tabtype === '1') {
      ctrlZones = Array.from(this.ContentArea.querySelectorAll('.CanvasZone'))
      //CanvasZoneContainer no longer suported, only shows in workbench now
      //ctrlZones = Array.from(this.ContentArea.querySelectorAll('.CanvasZoneContainer'))
        .filter(d => { return !this.findInAncestors(d, this.domElement); });
    } else {
      ctrlZones = Array.from(this.ContentArea.querySelectorAll('.ControlZone'))
        .filter(d => { return !this.findInAncestors(d, this.domElement); });
    }

    this.logDiagnostic('Discovered ' + String(ctrlZones.length) + ' tab zone(s) for TabType=' + tabtype + ' (processed=' + String(this.tabsProcessedCount) + ', foundAnotherTab=' + String(this.foundAnotherTab) + ').');

    // Add an attribute to each webpart so that the "loading" CSS rules apply, but only do it the first time the
    //   page loads and not if we're in edit mode
    if (this.displayMode !== DisplayMode.Edit && document.readyState !== 'complete') {
      ctrlZones.forEach(d => {
        d.setAttribute('part-Loading', 'true');
        d.classList.add('sectionHidden');
      });
    } else if (this.displayMode === DisplayMode.Edit) {
      ctrlZones.forEach(d => {
        d.removeAttribute('part-Loading');
        d.classList.remove('sectionHidden');
      });
    }

    const configuredTabs = this.ensureConfiguredTabs(ctrlZones.length);
    if (configuredTabs.length !== ctrlZones.length) {
      console.warn('[TabComponentWebPart] Configured tab entries (' + String(configuredTabs.length) + ') do not match discovered zone count (' + String(ctrlZones.length) + '). Extra zones will use default "Tab N" labels; extra configured entries are ignored.');
    }

    // Titles for the tabs from the webpart configuration or a default
    const tabHeadings: Array<string> = ctrlZones.map((d, i, e) => {
      const tab = configuredTabs[i] as ITabCollectionItem | undefined;
      return (tab && tab.Title) ? tab.Title : `Tab ${i + 1}`;
    });

    const tabConfigs: ITabVisualSettings[] = ctrlZones.map((d, i, e) => {
      const tab = configuredTabs[i] as ITabCollectionItem | undefined;
      return {
        textPosition: (tab && tab.textPosition) ? tab.textPosition : 'left',
        imageUrl: tab ? tab.imageUrl : undefined,
        imagePosition: (tab && tab.imagePosition) ? tab.imagePosition : 'left',
        onlyImage: tab ? (tab.onlyImage === true) : false
      };
    });

    const globalFontSize = this.parseNumberSetting(this.properties.fontSize, 14, 10, 36);
    const globalTabHeight = this.parseNumberSetting(this.properties.tabHeight, 60, 36, 160);
    const globalTabWidth = this.parseNumberSetting(this.properties.tabWidth, 120, 60, 400);

    const element: React.ReactElement<ITabComponentProps> = React.createElement(
      TabComponent,
      {
        TabHeaders: tabHeadings,
        ControlZones: ctrlZones,
        PageInEditMode: this.displayMode === DisplayMode.Edit,
        TabConfigs: tabConfigs,
        EnableDiagnostics: this.properties.enableDiagnostics !== false,
        GlobalFontSettings: {
          fontFamily: this.properties.fontFamily || 'Segoe UI',
          fontStyle: this.properties.fontStyle || 'normal',
          isBold: this.properties.isBold || false,
          fontSize: globalFontSize,
          tabHeight: globalTabHeight,
          tabWidth: globalTabWidth,
          autoTabWidth: this.properties.autoTabWidth || false,
          tabShape: this.properties.tabShape || 'rounded',
          tabTextColor: this.properties.tabTextColor || '#ffffff',
          tabBorderColor: this.properties.tabBorderColor || '#000000',
          tabBorderWidth: this.parseNumberSetting(this.properties.tabBorderWidth, 0, 0, 20),
          tabBorderStyle: this.properties.tabBorderStyle || 'solid',
          tabCornerRadius: this.parseNumberSetting(this.properties.tabCornerRadius, 10, 0, 40),
          tabLineColor: this.properties.tabLineColor || '#8A1717',
          tabLineWidth: this.parseNumberSetting(this.properties.tabLineWidth, 3, 0, 20),
          webPartBorderColor: this.properties.webPartBorderColor || '#cccccc',
          webPartBorderWidth: this.parseNumberSetting(this.properties.webPartBorderWidth, 0, 0, 20),
          webPartBorderStyle: this.properties.webPartBorderStyle || 'solid',
          webPartCornerStyle: this.properties.webPartCornerStyle || 'square',
          webPartCornerRadius: this.parseNumberSetting(this.properties.webPartCornerRadius, 8, 0, 40),
          inactiveImageFade: this.parseNumberSetting(this.properties.inactiveImageFade, 45, 0, 100),
          contentPaddingTop: this.parseNumberSetting(this.properties.contentPaddingTop, 0, 0, 100),
          contentPaddingBottom: this.parseNumberSetting(this.properties.contentPaddingBottom, 0, 0, 100)
        }
      } as ITabComponentProps
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    this.logDiagnostic('onDispose invoked; unmounting React tree.');
    releaseOptionalFullWidth(this.domElement.ownerDocument, this.context.instanceId);
    this.domElement.removeAttribute('data-spse-tab-type');
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private bindPropVals(value: string): string {
    //Here is where you could do some validation
    return '';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fullVersionLabel = 'Version: ' + this.getWebPartVersion();
    return {
      pages: [
        {
          header: {
            description: ''
          },
          displayGroupsAsAccordion: false,
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldCollectionData('tabs', {
                  key: 'tabsListField',
                  label: strings.Tabs,
                  value: this.getConfiguredTabs(),
                  panelHeader: strings.ManageTabs,
                  manageBtnLabel: strings.ManageTabs,
                  enableSorting: true,
                  fields: [
                    { id: 'Title', title: 'Title', required: true, type: CustomCollectionFieldType.string },
                    {
                      id: 'textPosition',
                      title: 'Text Position',
                      required: false,
                      type: CustomCollectionFieldType.dropdown,
                      defaultValue: 'left',
                      options: [
                        { key: 'left', text: 'Left of Tab' },
                        { key: 'center', text: 'Center of Tab' },
                        { key: 'right', text: 'Right of Tab' }
                      ]
                    },
                    {
                      id: 'imageUrl',
                      title: 'Image URL',
                      required: false,
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: this.renderTabImagePicker.bind(this)
                    },
                    {
                      id: 'imagePosition',
                      title: 'Image Position',
                      required: false,
                      type: CustomCollectionFieldType.dropdown,
                      defaultValue: 'left',
                      options: [
                        { key: 'left', text: 'Image Left' },
                        { key: 'right', text: 'Image Right' }
                      ]
                    },
                    { id: 'onlyImage', title: 'Only Image', required: false, type: CustomCollectionFieldType.boolean, defaultValue: false }
                  ]
                }),
                PropertyPaneCheckbox('forceFullWidth', { text: 'Force Full Width' }),
                PropertyPaneDropdown('TabType', {
                  label: strings.TabType,
                  options: [
                    { key: '1', text: 'Sections' },
                    { key: '2', text: 'WebParts' }
                  ],
                  selectedKey: this.properties.TabType || '1'
                }),
                PropertyPaneDropdown('fontFamily', {
                  label: 'Font',
                  options: [
                    { key: 'Segoe UI', text: 'Segoe UI' },
                    { key: 'Arial', text: 'Arial' },
                    { key: 'Calibri', text: 'Calibri' },
                    { key: 'Verdana', text: 'Verdana' },
                    { key: 'Tahoma', text: 'Tahoma' },
                    { key: 'Trebuchet MS', text: 'Trebuchet MS' },
                    { key: 'Georgia', text: 'Georgia' },
                    { key: 'Times New Roman', text: 'Times New Roman' },
                    { key: 'Courier New', text: 'Courier New' }
                  ],
                  selectedKey: this.properties.fontFamily || 'Segoe UI'
                }),
                PropertyPaneDropdown('fontStyle', {
                  label: 'Font Style',
                  options: [
                    { key: 'normal', text: 'Normal' },
                    { key: 'italic', text: 'Italic' },
                    { key: 'oblique', text: 'Oblique' }
                  ],
                  selectedKey: this.properties.fontStyle || 'normal'
                }),
                PropertyPaneCheckbox('isBold', {
                  text: 'Bold',
                  checked: this.properties.isBold || false
                }),
                PropertyPaneSlider('fontSize', {
                  label: 'Font Size (px)',
                  min: 10,
                  max: 36,
                  step: 1,
                  value: this.parseNumberSetting(this.properties.fontSize, 14, 10, 36)
                }),
                PropertyPaneSlider('tabHeight', {
                  label: 'Tab Height (px)',
                  min: 36,
                  max: 160,
                  step: 1,
                  value: this.parseNumberSetting(this.properties.tabHeight, 60, 36, 160)
                }),
                PropertyPaneCheckbox('autoTabWidth', {
                  text: 'Auto Size Width',
                  checked: this.properties.autoTabWidth || false
                }),
                PropertyPaneSlider('tabWidth', {
                  label: 'Tab Width (px)',
                  min: 60,
                  max: 400,
                  step: 1,
                  value: this.parseNumberSetting(this.properties.tabWidth, 120, 60, 400),
                  disabled: this.properties.autoTabWidth === true
                }),
                PropertyPaneSlider('inactiveImageFade', {
                  label: 'Inactive Image Fade (%)',
                  min: 0,
                  max: 100,
                  step: 5,
                  value: this.parseNumberSetting(this.properties.inactiveImageFade, 45, 0, 100)
                }),
                PropertyFieldColorPicker('selectedColor', {
                  label: strings.SelectedColor,
                  selectedColor: this.properties.selectedColor || '#8A1717',
                  properties: this.properties,
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'tabsSelectedColorField'
                }),
                PropertyFieldColorPicker('disableColor', {
                  label: strings.DisableColor,
                  selectedColor: this.properties.disableColor || '#393939',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  style: PropertyFieldColorPickerStyle.Inline,
                  properties: this.properties,
                  key: 'tabsDisableColorField'
                }),
                PropertyFieldColorPicker('tabTextColor', {
                  label: 'Tab Foreground Color',
                  selectedColor: this.properties.tabTextColor || '#ffffff',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  style: PropertyFieldColorPickerStyle.Inline,
                  properties: this.properties,
                  key: 'tabsTextColorField'
                }),
                PropertyFieldColorPicker('tabBorderColor', {
                  label: 'Tab Border Color',
                  selectedColor: this.properties.tabBorderColor || '#000000',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  style: PropertyFieldColorPickerStyle.Inline,
                  properties: this.properties,
                  key: 'tabsBorderColorField'
                }),
                PropertyPaneSlider('tabBorderWidth', {
                  label: 'Tab Border Width (px)',
                  min: 0,
                  max: 20,
                  step: 1,
                  value: this.parseNumberSetting(this.properties.tabBorderWidth, 0, 0, 20)
                }),
                PropertyPaneDropdown('tabBorderStyle', {
                  label: 'Tab Border Type',
                  options: [
                    { key: 'none', text: 'None' },
                    { key: 'solid', text: 'Solid' },
                    { key: 'dashed', text: 'Dashed' },
                    { key: 'dotted', text: 'Dotted' },
                    { key: 'double', text: 'Double' }
                  ],
                  selectedKey: this.properties.tabBorderStyle || 'solid'
                }),
                PropertyPaneDropdown('tabShape', {
                  label: 'Tab Corner Style',
                  options: [
                    { key: 'square', text: 'Square' },
                    { key: 'rounded', text: 'Rounded' }
                  ],
                  selectedKey: this.properties.tabShape || 'rounded'
                }),
                ...((this.properties.tabShape || 'rounded') === 'rounded' ? [
                  PropertyPaneSlider('tabCornerRadius', {
                    label: 'Tab Corner Radius (px)',
                    min: 0,
                    max: 40,
                    step: 1,
                    value: this.parseNumberSetting(this.properties.tabCornerRadius, 10, 0, 40)
                  })
                ] : []),
                PropertyFieldColorPicker('tabLineColor', {
                  label: 'Line Below Tabs Color',
                  selectedColor: this.properties.tabLineColor || '#8A1717',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  style: PropertyFieldColorPickerStyle.Inline,
                  properties: this.properties,
                  key: 'tabsLineColorField'
                }),
                PropertyPaneSlider('tabLineWidth', {
                  label: 'Line Below Tabs Size (px)',
                  min: 0,
                  max: 20,
                  step: 1,
                  value: this.parseNumberSetting(this.properties.tabLineWidth, 3, 0, 20)
                }),
                PropertyFieldColorPicker('webPartBorderColor', {
                  label: 'Web Part Border Color',
                  selectedColor: this.properties.webPartBorderColor || '#cccccc',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  style: PropertyFieldColorPickerStyle.Inline,
                  properties: this.properties,
                  key: 'tabsWebPartBorderColorField'
                }),
                PropertyPaneSlider('webPartBorderWidth', {
                  label: 'Web Part Border Width (px)',
                  min: 0,
                  max: 20,
                  step: 1,
                  value: this.parseNumberSetting(this.properties.webPartBorderWidth, 0, 0, 20)
                }),
                PropertyPaneDropdown('webPartBorderStyle', {
                  label: 'Web Part Border Type',
                  options: [
                    { key: 'none', text: 'None' },
                    { key: 'solid', text: 'Solid' },
                    { key: 'dashed', text: 'Dashed' },
                    { key: 'dotted', text: 'Dotted' },
                    { key: 'double', text: 'Double' }
                  ],
                  selectedKey: this.properties.webPartBorderStyle || 'solid'
                }),
                PropertyPaneDropdown('webPartCornerStyle', {
                  label: 'Web Part Corner Style',
                  options: [
                    { key: 'square', text: 'Square' },
                    { key: 'rounded', text: 'Rounded' }
                  ],
                  selectedKey: this.properties.webPartCornerStyle || 'square'
                }),
                ...((this.properties.webPartCornerStyle || 'square') === 'rounded' ? [
                  PropertyPaneSlider('webPartCornerRadius', {
                    label: 'Web Part Corner Radius (px)',
                    min: 0,
                    max: 40,
                    step: 1,
                    value: this.parseNumberSetting(this.properties.webPartCornerRadius, 8, 0, 40)
                  })
                ] : []),
                PropertyPaneSlider('contentPaddingTop', {
                  label: 'Content Padding Top (px)',
                  min: 0,
                  max: 100,
                  step: 1,
                  value: this.parseNumberSetting(this.properties.contentPaddingTop, 0, 0, 100)
                }),
                PropertyPaneSlider('contentPaddingBottom', {
                  label: 'Content Padding Bottom (px)',
                  min: 0,
                  max: 100,
                  step: 1,
                  value: this.parseNumberSetting(this.properties.contentPaddingBottom, 0, 0, 100)
                }),
                PropertyPaneTextField('overrideCSS', {
                  label: strings.OverrideCSS,
                  validateOnFocusOut: true,
                  onGetErrorMessage: this.bindPropVals.bind(this)
                }),
                PropertyPaneCheckbox('useGlobalCSS', {
                  text: strings.useGlobalCSS,
                  disabled: false
                }),
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
        }
      ]
    };
  }

  private parseNumberSetting(value: number | string | undefined, fallback: number, min: number, max: number): number {
    const parsedValue = typeof value === 'number' ? value : Number(value);
    if (isNaN(parsedValue)) {
      return fallback;
    }

    if (parsedValue < min) {
      return min;
    }

    if (parsedValue > max) {
      return max;
    }

    return parsedValue;
  }

  private getConfiguredTabs(): ITabCollectionItem[] {
    if (Array.isArray(this.properties.collectionData) && this.properties.collectionData.length > 0) {
      return this.properties.collectionData as ITabCollectionItem[];
    }

    return Array.isArray(this.properties.tabs) ? this.properties.tabs as ITabCollectionItem[] : [];
  }

  private ensureConfiguredTabs(zoneCount: number): ITabCollectionItem[] {
    const configuredTabs = this.getConfiguredTabs().slice();
    if (zoneCount <= configuredTabs.length) {
      return configuredTabs;
    }

    for (let index = configuredTabs.length; index < zoneCount; index += 1) {
      configuredTabs.push({
        Title: `Tab ${index + 1}`,
        textPosition: 'left',
        imagePosition: 'left',
        onlyImage: false
      });
    }
    this.properties.tabs = configuredTabs;
    this.properties.collectionData = [];
    this.logDiagnostic('Created missing configured tab entries. Count=' + String(zoneCount) + '.');
    return configuredTabs;
  }
}
