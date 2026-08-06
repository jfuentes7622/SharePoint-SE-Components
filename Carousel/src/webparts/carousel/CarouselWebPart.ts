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
  PropertyPaneLabel
} from '@microsoft/sp-webpart-base';

import * as strings from 'CarouselWebPartStrings';
import Carousel from './components/Carousel';
import { ICarouselProps } from './components/ICarouselProps';
import RecordSvc from '../../services/RecordService';
import { SPContexts } from './components/SPContexts';
import { ConfigData } from './components/ConfigData';

import SlideItemModel from './components/SlideItem';
import {SPHttpClient,
  SPHttpClientResponse,
 }  from '@microsoft/sp-http';

import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface ICarouselWebPartProps {
  carouselWidth: number;
  carouselHeight: number;
  carouselBackgroundColor: string;
  carouselTransitionInterval: string;
  carouselSlideInterval: string;
  carouselSlideItems: Array<SlideItemModel>;
  carouselSlideLibrary: string;
  imageIsCircle: boolean;
  enableDiagnostics: boolean;
  recSvc: RecordSvc;
  spfxContext: WebPartContext;
}

export default class CarouselWebPart extends BaseClientSideWebPart<ICarouselWebPartProps> {

  private _recordSvc: RecordSvc;
  private _siteLists: string[];

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

    this.logDiagnostic('render() called');
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
        carouselBackgroundColor: this.properties.carouselBackgroundColor,
        carouselSlideInterval: this.properties.carouselSlideInterval,
        carouselSlideItems: this.properties.carouselSlideItems,
        carouselTransitionInterval: this.properties.carouselTransitionInterval,
        carouselWidth: this.properties.carouselWidth,
        carouselHeight: this.properties.carouselHeight,
        carouselSlideLibrary: this.properties.carouselSlideLibrary,
        imageIsCircle: this.properties.imageIsCircle,
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

   private GetNewRecSvc(): RecordSvc {
    return new RecordSvc(
      <ConfigData>{        
        slideListName: this.properties.carouselSlideLibrary,
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
    ReactDom.unmountComponentAtNode(this.domElement);
  }

/*   protected get dataVersion(): Version {
    return Version.parse('1.0');
  } */

  protected async onInit(): Promise<void> {
    this.logDiagnostic('onInit() called');
    this._siteLists = await this._getSiteLists();
    return super.onInit();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.logDiagnostic('onPropertyPaneFieldChanged: ' + propertyPath);
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fullVersionLabel = 'Version: ' + this.getWebPartVersion();

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
                options: this._siteLists.map((list: string) => {
                  return <IPropertyPaneDropdownOption>{
                    key: list, text: list
                  };
                 }),
                })
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
                PropertyFieldColorPicker('carouselBackgroundColor', {
                  label: 'Carousel Background Color',
                  selectedColor: this.properties.carouselBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'carouselBackgroundColorField'
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
