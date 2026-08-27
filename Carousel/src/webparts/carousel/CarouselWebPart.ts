import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';

import {
  WebPartContext,
} from '@microsoft/sp-webpart-base';

import {
  BaseClientSideWebPart,
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField, 
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'CarouselWebPartStrings';
import Carousel from './components/Carousel';
import { ICarouselProps } from './components/ICarouselProps';
import RecordSvc from '../../services/RecordService';
import { SPContexts } from './components/SPContexts';
import { ConfigData } from './components/ConfigData';
import { releaseOptionalFullWidth, updateResponsiveOptionalFullWidth } from '../shared/deterministicFullWidth';

import SlideItemModel from './components/SlideItem';
import {SPHttpClient,
  SPHttpClientResponse,
 }  from '@microsoft/sp-http';

import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface ICarouselWebPartProps {
  forceFullWidth?: boolean;
  carouselWidth: number;
  carouselHeight: number;
  carouselTransitionInterval: string;
  carouselSlideInterval: string;
  carouselSlideItems: Array<SlideItemModel>;
  carouselSlideLibrary: string;
  slideTitleField: string;
  slideDescriptionField: string;
  slideLinkField: string;
  imageIsCircle: boolean;
  webPartBackgroundColor: string;
  webPartBorderColor: string;
  webPartBorderWidth: number;
  webPartBorderStyle: string;
  webPartCornerRadius: number;
  webPartPadding: number;
  showTitle: boolean;
  title: string;
  titleTextColor: string;
  titleFontFamily: string;
  titleFontSize: number;
  titleFontStyle: string;
  titleFontBold: boolean;
  titleAlignment: string;
  titleBottomSpacing: number;
  showSlideTitle: boolean;
  slideTitleColor: string;
  slideTitleFontFamily: string;
  slideTitleFontSize: number;
  slideTitleFontStyle: string;
  slideTitleFontBold: boolean;
  slideTitleAlignment: string;
  showSlideDescription: boolean;
  slideDescriptionColor: string;
  slideDescriptionFontFamily: string;
  slideDescriptionFontSize: number;
  slideDescriptionFontStyle: string;
  slideDescriptionFontBold: boolean;
  slideDescriptionAlignment: string;
  showSlideLink: boolean;
  slideLinkColor: string;
  slideLinkFontFamily: string;
  slideLinkFontSize: number;
  slideLinkFontStyle: string;
  slideLinkFontBold: boolean;
  slideLinkAlignment: string;
  slideCaptionBackgroundColor: string;
  slideCaptionPadding: number;
  enableDiagnostics: boolean;
  recSvc: RecordSvc;
  spfxContext: WebPartContext;
}

export default class CarouselWebPart extends BaseClientSideWebPart<ICarouselWebPartProps> {

  private _recordSvc: RecordSvc;
  private _siteLists: string[] = [];
  private _columnOptions: IPropertyPaneDropdownOption[] = [];

  public constructor() {
    super();
    // Hack: to invoke correctly the onPropertyChange function outside this class
    // we need to bind this object on it first
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }
    console.log('[CarouselWebPart] ' + message);
  }

  public render(): void {
    updateResponsiveOptionalFullWidth(this.domElement, this.context.instanceId, this.properties.forceFullWidth === true);

    this.logDiagnostic('render() called');
    const webPartBackgroundColor = this.getVisibleBackgroundColor(this.properties.webPartBackgroundColor, '#ffffff');
    this.domElement.style.backgroundColor = 'transparent';
    this._recordSvc = this.GetNewRecSvc();

     //workaround sharepoint's cache clobbering issue with uri fragments using history manipulation
     window.addEventListener('beforeunload', () => {
      //only affect the history if the current url has a fragment
      if (window.location.hash) {
        //generate unique id of form ####-####
        const fragment_noonce_value: string = `${Math.random().toString().substr(2, 4)}-${Math.random().toString().substr(2, 4)}`;
        //set fragment-noonce query param, and keep other existing query parameters
        const searchParams: URLSearchParams = new URLSearchParams(window.location.search);
        searchParams.set('fragment-noonce', fragment_noonce_value);
        //replace history, preventing the browser from pulling the cached json result that clobbered the original html result
        window.history.replaceState({}, "", `?${searchParams.toString()}${window.location.hash}`);
      }
    });

    const normalizedWidth = Number(this.properties.carouselWidth) > 0 ? Number(this.properties.carouselWidth) : 500;
    const normalizedHeight = Number(this.properties.carouselHeight);

    this.domElement.style.setProperty('--cWidth', `${normalizedWidth}px`);
    this.domElement.style.setProperty('--cHeight', (isFinite(normalizedHeight) && normalizedHeight > 0) ? `${normalizedHeight}px` : 'auto');

    const element: React.ReactElement<ICarouselProps> = React.createElement(
      Carousel,
      {
        carouselSlideInterval: this.properties.carouselSlideInterval,
        carouselSlideItems: this.properties.carouselSlideItems,
        carouselTransitionInterval: this.properties.carouselTransitionInterval,
        carouselWidth: this.properties.carouselWidth,
        carouselHeight: this.properties.carouselHeight,
        carouselSlideLibrary: this.properties.carouselSlideLibrary,
        slideTitleField: this.properties.slideTitleField || 'Title',
        slideDescriptionField: this.properties.slideDescriptionField || '',
        slideLinkField: this.properties.slideLinkField || '',
        imageIsCircle: this.properties.imageIsCircle,
        webPartBackgroundColor: webPartBackgroundColor,
        webPartBorderColor: this.properties.webPartBorderColor || '#d2d0ce',
        webPartBorderWidth: this.properties.webPartBorderWidth || 0,
        webPartBorderStyle: this.properties.webPartBorderStyle || 'solid',
        webPartCornerRadius: this.properties.webPartCornerRadius || 0,
        webPartPadding: this.properties.webPartPadding || 0,
        showTitle: this.properties.showTitle === true,
        title: this.properties.title || '',
        titleTextColor: this.properties.titleTextColor || '#323130',
        titleFontFamily: this.properties.titleFontFamily || 'inherit',
        titleFontSize: this.properties.titleFontSize || 20,
        titleFontStyle: this.properties.titleFontStyle || 'normal',
        titleFontBold: this.properties.titleFontBold !== false,
        titleAlignment: this.properties.titleAlignment || 'center',
        titleBottomSpacing: this.properties.titleBottomSpacing || 0,
        showSlideTitle: this.properties.showSlideTitle !== false,
        slideTitleColor: this.properties.slideTitleColor || '#ffffff',
        slideTitleFontFamily: this.properties.slideTitleFontFamily || 'inherit',
        slideTitleFontSize: this.properties.slideTitleFontSize || 22,
        slideTitleFontStyle: this.properties.slideTitleFontStyle || 'normal',
        slideTitleFontBold: this.properties.slideTitleFontBold !== false,
        slideTitleAlignment: this.properties.slideTitleAlignment || 'left',
        showSlideDescription: this.properties.showSlideDescription !== false,
        slideDescriptionColor: this.properties.slideDescriptionColor || '#ffffff',
        slideDescriptionFontFamily: this.properties.slideDescriptionFontFamily || 'inherit',
        slideDescriptionFontSize: this.properties.slideDescriptionFontSize || 14,
        slideDescriptionFontStyle: this.properties.slideDescriptionFontStyle || 'normal',
        slideDescriptionFontBold: this.properties.slideDescriptionFontBold === true,
        slideDescriptionAlignment: this.properties.slideDescriptionAlignment || 'left',
        showSlideLink: this.properties.showSlideLink !== false,
        slideLinkColor: this.properties.slideLinkColor || '#ffffff',
        slideLinkFontFamily: this.properties.slideLinkFontFamily || 'inherit',
        slideLinkFontSize: this.properties.slideLinkFontSize || 14,
        slideLinkFontStyle: this.properties.slideLinkFontStyle || 'normal',
        slideLinkFontBold: this.properties.slideLinkFontBold !== false,
        slideLinkAlignment: this.properties.slideLinkAlignment || 'left',
        slideCaptionBackgroundColor: this.properties.slideCaptionBackgroundColor || 'rgba(0,0,0,0.65)',
        slideCaptionPadding: this.properties.slideCaptionPadding || 12,
        enableDiagnostics: this.properties.enableDiagnostics !== false,
        spfxContext:this.context,
        recSvc: this._recordSvc
      }
    );

     ReactDom.render(element, this.domElement);
        
  }

  private async _getSiteLists(): Promise<string[]> {
    this.logDiagnostic('_getSiteLists() fetching document libraries');
    const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title&$filter=Hidden eq false and BaseTemplate eq 101&$orderby=Title`;
    const rawResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1);
  
    const lists: string[] = (await rawResponse.json()).value.map(
      (list: {Title: string}) => {
        return list.Title;
      }
    );
    this.logDiagnostic('_getSiteLists() found ' + lists.length + ' library(ies)');
    return lists;
  }

  private async loadLibraryColumns(libraryTitle: string): Promise<IPropertyPaneDropdownOption[]> {
    if (!libraryTitle) {
      return [];
    }

    const escapedTitle = libraryTitle.replace(/'/g, "''");
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${escapedTitle}')/fields?$select=Title,InternalName,Hidden&$filter=Hidden eq false`;
    const response = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    if (!response.ok) {
      this.logDiagnostic('Unable to load library columns: HTTP ' + response.status + ' ' + response.statusText);
      return [];
    }

    const json: any = await response.json();
    const options: IPropertyPaneDropdownOption[] = [{ key: '', text: '(None)' }];
    (json.value || []).forEach((field: any): void => {
      if (field.InternalName) {
        options.push({
          key: field.InternalName,
          text: String(field.Title || field.InternalName) + ' (' + String(field.InternalName) + ')'
        });
      }
    });
    options.splice(1, options.length - 1, ...options.slice(1).sort((left, right): number => String(left.text).localeCompare(String(right.text))));
    return options;
  }

  private resolveColumnMappings(): void {
    const keys = this._columnOptions.map((option: IPropertyPaneDropdownOption): string => String(option.key));
    const hasField = (fieldName: string): boolean => keys.indexOf(fieldName) !== -1;

    if (!hasField(this.properties.slideTitleField)) {
      this.properties.slideTitleField = hasField('Title') ? 'Title' : '';
    }
    if (!hasField(this.properties.slideDescriptionField)) {
      this.properties.slideDescriptionField = hasField('Description') ? 'Description' : '';
    }
    if (!hasField(this.properties.slideLinkField)) {
      this.properties.slideLinkField = hasField('LinkTarget') ? 'LinkTarget' : (hasField('ClickLink') ? 'ClickLink' : '');
    }
  }

   private GetNewRecSvc(): RecordSvc {
    return new RecordSvc(
      <ConfigData>{        
        slideListName: this.properties.carouselSlideLibrary,
        slideTitleField: this.properties.slideTitleField || 'Title',
        slideDescriptionField: this.properties.slideDescriptionField || '',
        slideLinkField: this.properties.slideLinkField || '',
        siteUrl: decodeURI(this.context.pageContext.web.absoluteUrl),
        enableDiagnostics: this.properties.enableDiagnostics !== false
      },
      <SPContexts>{
        spHttpClient: this.context.spHttpClient,
        absUrl: decodeURI(this.context.pageContext.web.absoluteUrl),
        aUrl: window.location.protocol +  "//" + window.location.host,
        relUrl: decodeURI(this.context.pageContext.web.serverRelativeUrl)
      }
    );
  }  

  protected onDispose(): void {
    this.logDiagnostic('onDispose() called');
    releaseOptionalFullWidth(this.domElement.ownerDocument, this.context.instanceId);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

/*   protected get dataVersion(): Version {
    return Version.parse('1.0');
  } */

  protected async onInit(): Promise<void> {
    this.logDiagnostic('onInit() called');
    this._siteLists = await this._getSiteLists();
    this._columnOptions = await this.loadLibraryColumns(this.properties.carouselSlideLibrary);
    this.resolveColumnMappings();
    return super.onInit();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    this.logDiagnostic('onPropertyPaneFieldChanged: ' + propertyPath);
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    (this.properties as any)[propertyPath] = newValue;
    if (propertyPath === 'carouselSlideLibrary') {
      this._columnOptions = await this.loadLibraryColumns(String(newValue || ''));
      this.resolveColumnMappings();
      this.context.propertyPane.refresh();
    }
    this.render();
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

  private getVisibleBackgroundColor(value: string, fallback: string): string {
    const color = String(value || '').replace(/\s/g, '').toLowerCase();
    return !color || color === 'transparent' || /rgba\([^,]+,[^,]+,[^,]+,0(?:\.0+)?\)/.test(color)
      ? fallback
      : value;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fullVersionLabel = 'Version: ' + this.getWebPartVersion();
    const fontOptions: IPropertyPaneDropdownOption[] = [
      { key: 'inherit', text: 'Page default' },
      { key: 'Segoe UI', text: 'Segoe UI' },
      { key: 'Arial', text: 'Arial' },
      { key: 'Georgia', text: 'Georgia' },
      { key: 'Tahoma', text: 'Tahoma' },
      { key: 'Trebuchet MS', text: 'Trebuchet MS' },
      { key: 'Verdana', text: 'Verdana' }
    ];

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
              groupName: 'Carousel Library',
              groupFields: [
               PropertyPaneDropdown('carouselSlideLibrary', {  
                label: 'Select Carousel Library',
                options: (this._siteLists || []).map((list: string) => {
                  return <IPropertyPaneDropdownOption>{
                    key: list, text: list
                  };
                 }),
                }),
                PropertyPaneDropdown('slideTitleField', {
                  label: 'Slide Title Column',
                  options: this._columnOptions,
                  selectedKey: this.properties.slideTitleField || 'Title'
                }),
                PropertyPaneDropdown('slideDescriptionField', {
                  label: 'Slide Description Column',
                  options: this._columnOptions,
                  selectedKey: this.properties.slideDescriptionField || ''
                }),
                PropertyPaneDropdown('slideLinkField', {
                  label: 'Slide Link Column',
                  options: this._columnOptions,
                  selectedKey: this.properties.slideLinkField || ''
                }),
                PropertyPaneCheckbox('forceFullWidth', { text: 'Force Full Width' })
              ]
             },
             {
              groupName: 'Image Carousel Properties',
              groupFields: [
                PropertyPaneTextField('carouselWidth', {
                  label: 'Carousel Width (ex:500, no px/%)'
                }),
                PropertyPaneTextField('carouselHeight', {
                  label: 'Carousel Height (ex:500, no px/%)'
                }),
                PropertyPaneTextField('carouselSlideInterval', {
                  label: 'Slide Interval (pause duration for each slide in mSeconds)'
                }),
                PropertyPaneTextField('carouselTransitionInterval', {
                  label: 'Transition Interval (duration of slide transition in mSeconds)'
                }),
                PropertyPaneCheckbox('imageIsCircle', {
                  text: 'Display images in a circle'
                })
              ]
            },
            {
              groupName: 'Web Part Appearance',
              groupFields: [
                PropertyFieldColorPicker('webPartBackgroundColor', {
                  label: 'Web Part Background Color',
                  selectedColor: this.getVisibleBackgroundColor(this.properties.webPartBackgroundColor, '#ffffff'),
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Color',
                  key: 'webPartBackgroundColorField'
                }),
                PropertyFieldColorPicker('webPartBorderColor', {
                  label: 'Web Part Border Color',
                  selectedColor: this.properties.webPartBorderColor || '#d2d0ce',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Color',
                  key: 'webPartBorderColorField'
                }),
                PropertyPaneSlider('webPartBorderWidth', {
                  label: 'Border Width', min: 0, max: 12, step: 1,
                  value: this.properties.webPartBorderWidth || 0
                }),
                PropertyPaneDropdown('webPartBorderStyle', {
                  label: 'Border Style',
                  selectedKey: this.properties.webPartBorderStyle || 'solid',
                  options: [
                    { key: 'solid', text: 'Solid' },
                    { key: 'dashed', text: 'Dashed' },
                    { key: 'dotted', text: 'Dotted' },
                    { key: 'double', text: 'Double' }
                  ]
                }),
                PropertyPaneSlider('webPartCornerRadius', {
                  label: 'Corner Radius', min: 0, max: 50, step: 1,
                  value: this.properties.webPartCornerRadius || 0
                }),
                PropertyPaneSlider('webPartPadding', {
                  label: 'Padding', min: 0, max: 50, step: 1,
                  value: this.properties.webPartPadding || 0
                })
              ]
            },
            {
              groupName: 'Title',
              groupFields: [
                PropertyPaneCheckbox('showTitle', {
                  text: 'Show Title', checked: this.properties.showTitle === true
                }),
                PropertyPaneTextField('title', { label: 'Title Text' }),
                PropertyFieldColorPicker('titleTextColor', {
                  label: 'Title Color',
                  selectedColor: this.properties.titleTextColor || '#323130',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'FontColor',
                  key: 'titleTextColorField'
                }),
                PropertyPaneDropdown('titleFontFamily', {
                  label: 'Title Font', options: fontOptions,
                  selectedKey: this.properties.titleFontFamily || 'inherit'
                }),
                PropertyPaneSlider('titleFontSize', {
                  label: 'Title Font Size', min: 10, max: 72, step: 1,
                  value: this.properties.titleFontSize || 20
                }),
                PropertyPaneDropdown('titleFontStyle', {
                  label: 'Title Font Style',
                  selectedKey: this.properties.titleFontStyle || 'normal',
                  options: [
                    { key: 'normal', text: 'Normal' },
                    { key: 'italic', text: 'Italic' },
                    { key: 'oblique', text: 'Oblique' }
                  ]
                }),
                PropertyPaneCheckbox('titleFontBold', {
                  text: 'Bold Title', checked: this.properties.titleFontBold !== false
                }),
                PropertyPaneDropdown('titleAlignment', {
                  label: 'Title Alignment',
                  selectedKey: this.properties.titleAlignment || 'center',
                  options: [
                    { key: 'left', text: 'Left' },
                    { key: 'center', text: 'Center' },
                    { key: 'right', text: 'Right' }
                  ]
                }),
                PropertyPaneSlider('titleBottomSpacing', {
                  label: 'Space Below Title', min: 0, max: 50, step: 1,
                  value: this.properties.titleBottomSpacing || 0
                })
              ]
            },
            {
              groupName: 'Slide Caption',
              groupFields: [
                PropertyFieldColorPicker('slideCaptionBackgroundColor', {
                  label: 'Caption Background Color',
                  selectedColor: this.properties.slideCaptionBackgroundColor || 'rgba(0,0,0,0.65)',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Color',
                  key: 'slideCaptionBackgroundColorField'
                }),
                PropertyPaneSlider('slideCaptionPadding', {
                  label: 'Caption Padding', min: 0, max: 40, step: 1,
                  value: typeof this.properties.slideCaptionPadding === 'number' ? this.properties.slideCaptionPadding : 12
                })
              ]
            },
            {
              groupName: 'Slide Title',
              groupFields: [
                PropertyPaneCheckbox('showSlideTitle', {
                  text: 'Show Library Title', checked: this.properties.showSlideTitle !== false
                }),
                PropertyFieldColorPicker('slideTitleColor', {
                  label: 'Slide Title Color', selectedColor: this.properties.slideTitleColor || '#ffffff',
                  onPropertyChange: this.onPropertyPaneFieldChanged, properties: this.properties,
                  disabled: false, alphaSliderHidden: true, style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'FontColor', key: 'slideTitleColorField'
                }),
                PropertyPaneDropdown('slideTitleFontFamily', {
                  label: 'Slide Title Font', options: fontOptions,
                  selectedKey: this.properties.slideTitleFontFamily || 'inherit'
                }),
                PropertyPaneSlider('slideTitleFontSize', {
                  label: 'Slide Title Font Size', min: 10, max: 72, step: 1,
                  value: this.properties.slideTitleFontSize || 22
                }),
                PropertyPaneDropdown('slideTitleFontStyle', {
                  label: 'Slide Title Font Style', selectedKey: this.properties.slideTitleFontStyle || 'normal',
                  options: [{ key: 'normal', text: 'Normal' }, { key: 'italic', text: 'Italic' }, { key: 'oblique', text: 'Oblique' }]
                }),
                PropertyPaneCheckbox('slideTitleFontBold', {
                  text: 'Bold Slide Title', checked: this.properties.slideTitleFontBold !== false
                }),
                PropertyPaneDropdown('slideTitleAlignment', {
                  label: 'Slide Title Alignment', selectedKey: this.properties.slideTitleAlignment || 'left',
                  options: [{ key: 'left', text: 'Left' }, { key: 'center', text: 'Center' }, { key: 'right', text: 'Right' }]
                })
              ]
            },
            {
              groupName: 'Slide Description',
              groupFields: [
                PropertyPaneCheckbox('showSlideDescription', {
                  text: 'Show Library Description', checked: this.properties.showSlideDescription !== false
                }),
                PropertyFieldColorPicker('slideDescriptionColor', {
                  label: 'Description Color', selectedColor: this.properties.slideDescriptionColor || '#ffffff',
                  onPropertyChange: this.onPropertyPaneFieldChanged, properties: this.properties,
                  disabled: false, alphaSliderHidden: true, style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'FontColor', key: 'slideDescriptionColorField'
                }),
                PropertyPaneDropdown('slideDescriptionFontFamily', {
                  label: 'Description Font', options: fontOptions,
                  selectedKey: this.properties.slideDescriptionFontFamily || 'inherit'
                }),
                PropertyPaneSlider('slideDescriptionFontSize', {
                  label: 'Description Font Size', min: 10, max: 48, step: 1,
                  value: this.properties.slideDescriptionFontSize || 14
                }),
                PropertyPaneDropdown('slideDescriptionFontStyle', {
                  label: 'Description Font Style', selectedKey: this.properties.slideDescriptionFontStyle || 'normal',
                  options: [{ key: 'normal', text: 'Normal' }, { key: 'italic', text: 'Italic' }, { key: 'oblique', text: 'Oblique' }]
                }),
                PropertyPaneCheckbox('slideDescriptionFontBold', {
                  text: 'Bold Description', checked: this.properties.slideDescriptionFontBold === true
                }),
                PropertyPaneDropdown('slideDescriptionAlignment', {
                  label: 'Description Alignment', selectedKey: this.properties.slideDescriptionAlignment || 'left',
                  options: [{ key: 'left', text: 'Left' }, { key: 'center', text: 'Center' }, { key: 'right', text: 'Right' }]
                })
              ]
            },
            {
              groupName: 'Slide Link',
              groupFields: [
                PropertyPaneCheckbox('showSlideLink', {
                  text: 'Show Library Link', checked: this.properties.showSlideLink !== false
                }),
                PropertyFieldColorPicker('slideLinkColor', {
                  label: 'Link Color', selectedColor: this.properties.slideLinkColor || '#ffffff',
                  onPropertyChange: this.onPropertyPaneFieldChanged, properties: this.properties,
                  disabled: false, alphaSliderHidden: true, style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'FontColor', key: 'slideLinkColorField'
                }),
                PropertyPaneDropdown('slideLinkFontFamily', {
                  label: 'Link Font', options: fontOptions,
                  selectedKey: this.properties.slideLinkFontFamily || 'inherit'
                }),
                PropertyPaneSlider('slideLinkFontSize', {
                  label: 'Link Font Size', min: 10, max: 48, step: 1,
                  value: this.properties.slideLinkFontSize || 14
                }),
                PropertyPaneDropdown('slideLinkFontStyle', {
                  label: 'Link Font Style', selectedKey: this.properties.slideLinkFontStyle || 'normal',
                  options: [{ key: 'normal', text: 'Normal' }, { key: 'italic', text: 'Italic' }, { key: 'oblique', text: 'Oblique' }]
                }),
                PropertyPaneCheckbox('slideLinkFontBold', {
                  text: 'Bold Link', checked: this.properties.slideLinkFontBold !== false
                }),
                PropertyPaneDropdown('slideLinkAlignment', {
                  label: 'Link Alignment', selectedKey: this.properties.slideLinkAlignment || 'left',
                  options: [{ key: 'left', text: 'Left' }, { key: 'center', text: 'Center' }, { key: 'right', text: 'Right' }]
                })
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
            }

          ]
        }
      ]
    };
  }

}
