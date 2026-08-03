//import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField, 
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import styles from './SpoLinkbuttonWebPart.module.scss';
import * as strings from 'SpoLinkbuttonWebPartStrings';

export interface ILinkButtonWebPartProps {
  description: string;
  Link: string;
  Align: string;
}

require ('./button.css');

export default class LinkButtonWebPart extends BaseClientSideWebPart<ILinkButtonWebPartProps> {
 
  private _validateText(value: string): string {
    // If validation is not successful, return a string with error message.   
    if (/<(br|basefont|form|div|script|action|src|img|hr|input|source|iframe|param|area|meta|!--|col|link|option|base|img|wbr|!DOCTYPE).*?>|<(a|abbr|acronym|address|applet|article|aside|audio|b|bdi|bdo|big|blockquote|body|button|canvas|caption|center|cite|code|colgroup|command|datalist|dd|del|details|dfn|dialog|dir|div|dl|dt|em|embed|fieldset|figcaption|figure|font|footer|form|frameset|head|header|hgroup|h1|h2|h3|h4|h5|h6|html|i|iframe|ins|kbd|keygen|label|legend|li|map|mark|menu|meter|nav|noframes|noscript|object|ol|optgroup|output|p|pre|progress|q|rp|rt|ruby|s|samp|script|section|select|small|span|strike|strong|style|sub|summary|sup|table|tbody|td|textarea|tfoot|th|thead|time|title|tr|track|tt|u|ul|var|video).*?<\/\2>/i.test(value)) { 
      return "Text Required, please try again!";
    }
    else if (value === null || value.trim().length === 0) {
      return 'Provide Description';
    }

  else if (value.length > 50) {
      return 'Description should not be longer than 50 characters';
    }
   else {
    // If validation is successful, return an empty string.
    //alert("regEx False");
    return "";}
  }

  private _validateUrl(value: string): string {
    // If validation is not successful, return a string with error message. 
   
  if (/<(br|basefont|form|div|script|action|src|img|hr|input|source|iframe|param|area|meta|!--|col|link|option|base|img|wbr|!DOCTYPE).*?>|<(a|abbr|acronym|address|applet|article|aside|audio|b|bdi|bdo|big|blockquote|body|button|canvas|caption|center|cite|code|colgroup|command|datalist|dd|del|details|dfn|dialog|dir|div|dl|dt|em|embed|fieldset|figcaption|figure|font|footer|form|frameset|head|header|hgroup|h1|h2|h3|h4|h5|h6|html|i|iframe|ins|kbd|keygen|label|legend|li|map|mark|menu|meter|nav|noframes|noscript|object|ol|optgroup|output|p|pre|progress|q|rp|rt|ruby|s|samp|script|section|select|small|span|strike|strong|style|sub|summary|sup|table|tbody|td|textarea|tfoot|th|thead|time|title|tr|track|tt|u|ul|var|video).*?<\/\2>/i.test(value)) { 
      return "Valid URl required, please try again!";
    }
   else if (value === null || value.trim().length === 0) {
        return 'Provide URL';
      }
    else {
      // If validation is successful, return an empty string.
      //alert("regEx False");
      return ""; }
    }

  public render(): void {
    if (this.properties.Align==='3') {
      this.domElement.innerHTML = `
      <div class="${ styles.linkButton }">
        <div class="${ styles.container }">
          <a class="button" style="margin-left:auto;" href=${this.properties.Link}>${this.properties.description.replace(/</g, "")}</a>
         </div>
      </div>`;      
    } else if (this.properties.Align==='2') {
      this.domElement.innerHTML = `
      <div class="${ styles.linkButton }">
        <div class="${ styles.container }">
          <a class="button" style="margin:auto;" href=${this.properties.Link}>${this.properties.description.replace(/</g, "")}</a>
         </div>
      </div>`;      
    }
      else {
      this.domElement.innerHTML = `
      <div class="${ styles.linkButton }">
        <div class="${ styles.container }">
          <a class="button" href=${this.properties.Link}>${this.properties.description.replace(/</g, "")}</a>
         </div>
      </div>`;
    }
    
    
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription + ` v${this.context.manifest.version}`,
          },
          groups: [
            {
              //groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Button Title",
                  placeholder: "Go to...",
                  onGetErrorMessage: this._validateText
                }),
                PropertyPaneTextField('Link', {
                   label:"Link",
                   placeholder: "https://",
                   onGetErrorMessage: this._validateUrl
                  }),
                PropertyPaneDropdown ('Align', {
                  label: 'Align Button',
                  options :[
                    {key: '1', text:'left'},
                    {key: '2', text:'center'},
                    {key: '3', text:'right'}
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
