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
  PropertyPaneLabel
} from '@microsoft/sp-webpart-base';

import { BaseClientSideWebPart,  WebPartContext } from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import * as strings from 'TabComponentWebPartStrings';
import TabComponent from './components/TabComponent';
import { ITabComponentProps, ITabVisualSettings } from './components/ITabComponentProps';
import ErrorComponent, { IErrorComponentProps } from './components/ErrorComponent';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldCustomList, CustomListFieldType } from 'sp-client-custom-fields/lib/PropertyFieldCustomList';

require('../tabComponent/assets/TabStyles-round.css');

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface IConfigListData {
  collectionData: any[];
}

export interface ITabComponentWebPartProps {
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

    this.logDiagnostic('Constructor invoked. TabType=' + String(this.properties && this.properties.TabType || '(default)'));
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
    // Remove the attribute from the webparts so the normal CSS rules apply
    Array.from(this.ContentArea.querySelectorAll('[part-Loading]')).forEach(d => {
      d.removeAttribute('part-Loading');
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
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

    //found another tab control, but might not begood
    if (targetElem.querySelector("div[id^='TabControl']") !== null) {
      this.foundAnotherTab = true;
      return true;
    }

    if (this.tabsProcessedCount > (this.properties.tabs.length + 1) && this.found === true) { return true; }

    return this.foundAnotherTab;
  }

  public render(): void {
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
    // Find the root section by traversing the DOM tree up until it finds an element with the CanvasSection class
    this.ContentArea = this.domElement;

      if (tabtype === '1') {
        this.ContentArea = this.domElement.ownerDocument.querySelectorAll(".Canvas, .grid")[0].parentElement!;
      } else {
         while (!this.ContentArea.classList.contains('CanvasSection')) { this.ContentArea = this.ContentArea.parentElement!; }
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

    const configuredTabs = this.getConfiguredTabs();
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
          tabShape: this.properties.tabShape || 'rounded'
        }
      } as ITabComponentProps
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    this.logDiagnostic('onDispose invoked; unmounting React tree.');
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
                PropertyFieldCustomList('collectionData', {
                  key: 'tabsCollectionData',
                  label: strings.Tabs,
                  headerText: 'Tabs',
                  value: this.properties.collectionData || this.properties.tabs || [],
                  context: this.context,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  fields: [
                    { id: 'Title', title: 'Title', required: true, type: CustomListFieldType.string },
                    { id: 'textPosition', title: 'Text Position (left|center|right)', required: false, type: CustomListFieldType.string },
                    { id: 'imageUrl', title: 'Image URL', required: false, type: CustomListFieldType.picture },
                    { id: 'imagePosition', title: 'Image Position (left|right)', required: false, type: CustomListFieldType.string },
                    { id: 'onlyImage', title: 'Only Image', required: false, type: CustomListFieldType.boolean }
                  ]
                }),
                PropertyPaneDropdown('TabType', {
                  label: strings.TabType,
                  options: [
                    { key: '1', text: 'Sections' },
                    { key: '2', text: 'WebParts' }
                  ],
                  selectedKey: this.properties.TabType || '1'
                }),
                PropertyPaneDropdown('tabShape', {
                  label: 'Tab Shape',
                  options: [
                    { key: 'rounded', text: 'Rounded Top Corners' },
                    { key: 'square', text: 'Squared' }
                  ],
                  selectedKey: this.properties.tabShape || 'rounded'
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
                PropertyPaneTextField('fontSize', {
                  label: 'Font Size (px)',
                  value: `${this.parseNumberSetting(this.properties.fontSize, 14, 10, 36)}`,
                  description: 'Enter a number between 10 and 36'
                }),
                PropertyPaneTextField('tabHeight', {
                  label: 'Tab Height (px)',
                  value: `${this.parseNumberSetting(this.properties.tabHeight, 60, 36, 160)}`,
                  description: 'Enter a number between 36 and 160'
                }),
                PropertyPaneCheckbox('autoTabWidth', {
                  text: 'Auto Size Width',
                  checked: this.properties.autoTabWidth || false
                }),
                PropertyPaneTextField('tabWidth', {
                  label: 'Tab Width (px)',
                  value: `${this.parseNumberSetting(this.properties.tabWidth, 120, 60, 400)}`,
                  description: 'Enter a number between 60 and 400',
                  disabled: this.properties.autoTabWidth === true
                }),
                PropertyFieldColorPicker('selectedColor', {
                  label: strings.SelectedColor,
                  selectedColor:this.properties.selectedColor,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'SelectColor',
                  key: 'tabsSelectedColorField'
                }),
                PropertyFieldColorPicker('disableColor', {
                  label: strings.DisableColor,
                  selectedColor: this.properties.disableColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'InactiveColor',
                  properties: this.properties,
                  key: 'tabsDisableColorField'
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
}
