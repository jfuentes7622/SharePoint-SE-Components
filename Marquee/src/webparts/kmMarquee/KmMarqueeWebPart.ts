import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode } from '@microsoft/sp-core-library';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import {
  BaseClientSideWebPart,
  WebPartContext,
} from '@microsoft/sp-webpart-base';

import {
IPropertyPaneConfiguration,
PropertyPaneTextField ,
PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'KmMarqueeWebPartStrings';
import KmMarquee from './components/KmMarquee';
import { IKmMarqueeProps } from './components/IKmMarqueeProps';

export interface IKmMarqueeWebPartProps {
  description: string;
  dMode: DisplayMode;
  spfxContext: WebPartContext;
  marqueeBackColor: string;
  marqueeTextColor: string;
  marqueeActive:boolean;
}

export default class KmMarqueeWebPart extends BaseClientSideWebPart<IKmMarqueeWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IKmMarqueeProps> = React.createElement(
      KmMarquee,
      {
        description: this.properties.description,
        dMode: this.displayMode,
        spfxContext:this.context,
        marqueeBackColor: this.properties.marqueeBackColor,
        marqueeTextColor: this.properties.marqueeTextColor,
        marqueeActive: this.properties.marqueeActive
      }
    );

    ReactDom.render(element, this.domElement);
    this.domElement.style.setProperty('--marqueeBackColor', (this.properties.marqueeBackColor !== undefined?this.properties.marqueeBackColor:"#333333"));
    this.domElement.style.setProperty('--marqueeTextColor', (this.properties.marqueeTextColor !== undefined?this.properties.marqueeTextColor:"#FAFAFA"));
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  private _validateText(value: string): string {
    // If validation is not successful, return a string with error message.   
    if (/<(br|basefont|form|div|script|action|src|img|hr|input|source|iframe|param|area|meta|!--|col|link|option|base|img|wbr|!DOCTYPE).*?>|<(a|abbr|acronym|address|applet|article|aside|audio|b|bdi|bdo|big|blockquote|body|button|canvas|caption|center|cite|code|colgroup|command|datalist|dd|del|details|dfn|dialog|dir|div|dl|dt|em|embed|fieldset|figcaption|figure|font|footer|form|frameset|head|header|hgroup|h1|h2|h3|h4|h5|h6|html|i|iframe|ins|kbd|keygen|label|legend|li|map|mark|menu|meter|nav|noframes|noscript|object|ol|optgroup|output|p|pre|progress|q|rp|rt|ruby|s|samp|script|section|select|small|span|strike|strong|style|sub|summary|sup|table|tbody|td|textarea|tfoot|th|thead|time|title|tr|track|tt|u|ul|var|video).*?<\/\2>/i.test(value)) { 
      return "Text Required, please try again!";
    }
    else if (value === null || value.trim().length === 0) {
      return 'Provide Description';
    }

  else if (value.length > 250) {
      return 'Description should not be longer than 250 characters';
    }
   else {
    // If validation is successful, return an empty string.
    return "";}
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
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  placeholder: "This is the scrolling text",
                  onGetErrorMessage: this._validateText
                }),
                PropertyFieldColorPicker('marqueeBackColor', {
                  label: "Announcements Background Color",
                  selectedColor: this.properties.marqueeBackColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  // debounce: 1000,
                  // isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'marqueeBackColor'                 
                }),
                PropertyFieldColorPicker('marqueeTextColor', {
                  label: "Announcements Text Color",
                  selectedColor: this.properties.marqueeTextColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  // debounce: 1000,
                  // isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'marqueeTextColor'
                }),
                PropertyPaneToggle('marqueeActive', {
                  label: 'Active',
                  checked: true
                })
          
              ]
            }
          ]
        }
      ]
    };
  }
}
