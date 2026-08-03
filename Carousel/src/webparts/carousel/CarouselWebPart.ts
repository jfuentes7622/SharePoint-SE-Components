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
  IPropertyPaneDropdownOption
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

export interface ICarouselWebPartProps {
  carouselWidth: number;
  carouselHeight: number;
  carouselBackgroundColor: string;
  carouselTransitionInterval: string;
  carouselSlideInterval: string;
  carouselSlideItems: Array<SlideItemModel>;
  carouselSlideLibrary: string;
  recSvc: RecordSvc;
  spfxContext: WebPartContext;
}

export default class CarouselWebPart extends BaseClientSideWebPart<ICarouselWebPartProps> {

  private _recordSvc: RecordSvc;
  private _siteLists: string[];

  public render(): void {

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
    
    this.domElement.style.setProperty('--cWidth', (this.properties.carouselWidth !== undefined?this.properties.carouselWidth.toString():"500px"));
    this.domElement.style.setProperty('--cHeight', (this.properties.carouselHeight !== undefined? this.properties.carouselHeight.toString():"200px"));
   
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
        spfxContext:this.context,
        recSvc: this._recordSvc
      }
    );

     ReactDom.render(element, this.domElement);
        
  }

  private async _getSiteLists(): Promise<string[]> {
    const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title&$filter=Hidden eq false&$orderby=Title&$BaseTemplate=109`;
    const rawResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1);
  
    return (await rawResponse.json()).value.map(
      (list: {Title: string}) => {
        return list.Title;
      }
    );
  }

   private GetNewRecSvc(): RecordSvc {
    return new RecordSvc(
      <ConfigData>{        
        slideListName: this.properties.carouselSlideLibrary,
        siteUrl: decodeURI(this.context.pageContext.web.absoluteUrl)
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
    ReactDom.unmountComponentAtNode(this.domElement);
  }

/*   protected get dataVersion(): Version {
    return Version.parse('1.0');
  } */

  protected async onInit(): Promise<void> {
    this._siteLists = await this._getSiteLists();
    return super.onInit();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription + ` v${this.context.manifest.version}`,
          },
          groups: [
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
                })      
              ]
            } 

          ]
        }
      ]
    };
  }

}
