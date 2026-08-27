import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,IWebPartContext,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneSlider,
  WebPartContext
} from '@microsoft/sp-webpart-base';

import {SPHttpClient,
       SPHttpClientResponse,
      }  from '@microsoft/sp-http';

import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { PropertyFieldDropdownWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldDropdownWithCallout';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';

require('./Accordion.css');

import * as strings from 'AccordionWebPartStrings';
import Accordion from './components/Accordion';
import { IAccordionProps } from './components/IAccordionProps';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface IAccordionWebPartProps {
    listName: string;
  itemName: string;
  itemContent: string;
  optionChoice: string;
  spfxContext: WebPartContext;
  headerBackgroundColor: string;
  headerTextColor: string;
  headerFontBold: boolean;
  headerFontSize: number;
  contentBackgroundColor: string;
  contentTextColor: string;
  contentFontBold: boolean;
  contentFontSize: number;
  cornerStyle: string;
  cornerRadius: number;
  showCloseBar: boolean;
  closeBarText: string;
  closeBarBackgroundColor: string;
  closeBarTextColor: string;
  closeBarFontSize: number;
  closeBarFontBold: boolean;
  closeBarAlignment: string;
  fontFamily: string;
  fontStyle: string;
  overrideCssUrl: string;
  enableDiagnostics: boolean;
}

const LOG_SOURCE: string = 'Accordion - ';

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {
private lists: IPropertyPaneDropdownOption[];
private listsDropDownDisabled: boolean=true;
private items: IPropertyPaneDropdownOption[];
private itemsDropDownDisabled: boolean=true;
private contents: IPropertyPaneDropdownOption[];
private contentsDropDownDisabled: boolean=true;
private options: IPropertyPaneDropdownOption[];
private optionsDropDownDisabled: boolean=true;

  public constructor() {
    super();
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }

  public render(): void { 

    this.logDiagnostic('render() called. optionChoice=' + String(this.properties.optionChoice) + ', listName=' + String(this.properties.listName || '(none)'));

    this._applyOverrideStylesheet();

    const element: React.ReactElement<IAccordionProps > = React.createElement(
      Accordion,
      {
        listName: this.properties.listName,
        itemName: this.properties.itemName,
        itemContent: this.properties.itemContent,
        optionChoice: this.properties.optionChoice,
        spfxContext:this.context,
        headerBackgroundColor: this.properties.headerBackgroundColor || '#f0f0f0',
        headerTextColor: this.properties.headerTextColor || '#000000',
        headerFontBold: this.properties.headerFontBold || false,
        headerFontSize: Math.max(10, Number(this.properties.headerFontSize) || 16),
        contentBackgroundColor: this.properties.contentBackgroundColor || '#ffffff',
        contentTextColor: this.properties.contentTextColor || '#000000',
        contentFontBold: this.properties.contentFontBold || false,
        contentFontSize: Math.max(10, Number(this.properties.contentFontSize) || 16),
        cornerStyle: this.properties.cornerStyle || 'square',
        cornerRadius: Math.max(0, Number(this.properties.cornerRadius) || 0),
        showCloseBar: this.properties.showCloseBar !== false,
        closeBarText: this.properties.closeBarText || 'Close others \u00d7',
        closeBarBackgroundColor: this.properties.closeBarBackgroundColor || '#333333',
        closeBarTextColor: this.properties.closeBarTextColor || '#ffffff',
        closeBarFontSize: Math.max(10, Number(this.properties.closeBarFontSize) || 12),
        closeBarFontBold: this.properties.closeBarFontBold === true,
        closeBarAlignment: this.properties.closeBarAlignment || 'right',
        fontFamily: this.properties.fontFamily || "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
        fontStyle: this.properties.fontStyle || 'normal',
        overrideCssUrl: this.properties.overrideCssUrl || '',
        enableDiagnostics: this.properties.enableDiagnostics !== false
         }
    );

    ReactDom.render(element, this.domElement);
  }

  private _applyOverrideStylesheet(): void {
    if (this.properties.overrideCssUrl) {
      this.logDiagnostic('Applying override stylesheet: ' + this.properties.overrideCssUrl);
      const existingLink: HTMLElement = document.getElementById('accordion-override-css');
      if (existingLink) {
        existingLink.remove();
      }
      const link: HTMLLinkElement = document.createElement('link');
      link.id = 'accordion-override-css';
      link.rel = 'stylesheet';
      link.href = this.properties.overrideCssUrl;
      document.head.appendChild(link);
    }
  }

  protected onPropertyPaneConfigurationStart():void {
    this.listsDropDownDisabled=!this.lists;
    this.itemsDropDownDisabled=!this.properties.listName || !this.items;
    this.logDiagnostic('onPropertyPaneConfigurationStart called. List Name: ' + this.properties.listName);
    this.logDiagnostic('Cached items: ' + this.items);
    if (this.lists) {
      return;
    }
    this.loadLists()
     .then((listOptions:IPropertyPaneDropdownOption[]): Promise <IPropertyPaneDropdownOption[]>=> {
       this.lists=listOptions;
       this.listsDropDownDisabled=false;
       this.context.propertyPane.refresh();
       return this.loadItems(this.properties.listName);
     })
     .then((itemOptions:IPropertyPaneDropdownOption[]): void=> {
      this.items=itemOptions;
      this.itemsDropDownDisabled=false;
      this.context.propertyPane.refresh();
      this.render();
     });
     this.loadLists()
     .then((listOptions:IPropertyPaneDropdownOption[]): Promise <IPropertyPaneDropdownOption[]>=> {
       this.lists=listOptions;
       this.listsDropDownDisabled=false;
       this.context.propertyPane.refresh();
       return this.loadContents(this.properties.listName); 
     })
    .then((contentOptions:IPropertyPaneDropdownOption[]): void=> {
     this.contents=contentOptions;
     this.contentsDropDownDisabled=false;
      this.context.propertyPane.refresh();
     this.render();
    }); 
    this.loadOptions()
     .then((optionChoice:IPropertyPaneDropdownOption[]): Promise <IPropertyPaneDropdownOption[]>=> {
       this.options=optionChoice;
       this.optionsDropDownDisabled=false;
       this.context.propertyPane.refresh();
       return this.loadOptions();
     });
  }

  protected onPropertyPaneFieldChanged(propertyPath:string, oldValue: any, newValue:any): void {
    this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
    (this.properties as any)[propertyPath] = newValue;
    if(propertyPath==='listName' && newValue) {
      //push new list value
      super.onPropertyPaneFieldChanged(propertyPath,oldValue,newValue);
      const previousItem: string=this.properties.itemName;
      const previousContent: string=this.properties.itemContent;
      //reset selected item
     
      this.properties.itemName='';
      this.properties.itemContent='';
      
      this.itemsDropDownDisabled=true;
      this.contentsDropDownDisabled=true;  
      
      //push new item value
      super.onPropertyPaneFieldChanged('itemName',previousItem, this.properties.itemName);
      super.onPropertyPaneFieldChanged('itemContent',previousContent, this.properties.itemContent);
     
           
      this.context.propertyPane.refresh();

      this.loadItems(newValue)
        .then((itemOptions:IPropertyPaneDropdownOption[]): void=> {
          //store items
          this.items=itemOptions;
          this.itemsDropDownDisabled=false;
          //re-render the web-part
          this.render();
          this.context.propertyPane.refresh();
        });
       this.loadContents(newValue)
        .then((contentOptions:IPropertyPaneDropdownOption[]): void=> {
          //store items
          this.contents=contentOptions;
          this.contentsDropDownDisabled=false;
          //re-render the web-part
          this.render();
          this.context.propertyPane.refresh();
        });
        this.loadOptions()
        .then((optionChoice:IPropertyPaneDropdownOption[]): void=> {
         //store items
         this.options=optionChoice;
          this.optionsDropDownDisabled=false;
          //re-render the web-part
          this.render();
          this.context.propertyPane.refresh();
        });
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.render();
    }
  }
   private async fetchLists(url:string): Promise<any> {
       this.logDiagnostic('fetchLists() requesting url: ' + url);
       return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
       .then ((response:SPHttpClientResponse) => {
         if (response.ok) {
           return response.json();
         }
         else {
           console.error(LOG_SOURCE + "Failed to get url:" + url + ". Error=" + response.statusText );
           return null;
         } }
       );
   }

/*    private async loadLists(): Promise<IPropertyPaneDropdownOption[]>{
     let currentWebUrl = this.context.pageContext.web.absoluteUrl;
     let url = currentWebUrl + '/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100';
     console.info('LOG_SOURCE' + 'LoadList url:' + url);
     return this.fetchLists(url).then((response) => {
       let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
       response.value.map((list:IODataList) => {
         console.info( LOG_SOURCE+ "Found list with title:" + list.Title);
         options.push({key: list.Id, text: list.Title});
     });
     return options;
     });
    } */

    private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {
  const currentWebUrl = this.context.pageContext.web.absoluteUrl;
  const url = `${currentWebUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100`;
  this.logDiagnostic('loadLists() url: ' + url);

  return this.fetchLists(url).then((response) => {
    const options: IPropertyPaneDropdownOption[] = [];

    // Define a minimal interface for the list object
    interface ICustomList {
      Id: string;
      Title: string;
    }

    (response.value as ICustomList[]).map((list) => {
      this.logDiagnostic('loadLists() found list with title: ' + list.Title);
      options.push({ key: list.Id, text: list.Title });
    });

    this.logDiagnostic('loadLists() completed. Count=' + String(options.length));
    return options;
  });
}


    private loadItems(listTitle: string): Promise<IPropertyPaneDropdownOption[]> {
    // giving only single line option for header
     let url=this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetById('"+ listTitle +"')/fields?$select=Title, InternalName&$filter=(((FieldTypeKind eq 2 and FromBaseType eq false) and (Hidden eq false)) or StaticName eq 'Title')";
     this.logDiagnostic('loadItems() url: ' + url);
     return this.fetchLists(url).then((response) => {
      let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        response.value.map((items:any) => {
        this.logDiagnostic('loadItems() found item title: ' + items.Title + ', internal name: ' + items.InternalName);
        options.push({key: items.InternalName, text: items.Title}); 
    });
    this.logDiagnostic('loadItems() completed. Count=' + String(options.length));
    return options;
    });
   }

   private loadContents(listTitle: string): Promise<IPropertyPaneDropdownOption[]> {
     //can be a single line or multiline column 
    let url=this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetById('"+ listTitle +"')/fields?$filter=(((FieldTypeKind eq 3 or FieldTypeKind eq 2) and (FromBaseType eq false) and (Hidden eq false)) or StaticName eq 'Title')";
    this.logDiagnostic('loadContents() url: ' + url);
    return this.fetchLists(url).then((response) => {
     let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      response.value.map((contents:any) => {
      this.logDiagnostic('loadContents() found item title: ' + contents.Title + ', internal name: ' + contents.InternalName);
      options.push({key: contents.InternalName, text: contents.Title});
      
   });
   this.logDiagnostic('loadContents() completed. Count=' + String(options.length));
   return options;
   }); 
  }

  private loadOptions(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options:
      IPropertyPaneDropdownOption[])=>void, reject: (error:any)=>void)=> {
        setTimeout(():void=> {
          resolve([{
            key: 'single',
            text: 'Single'
          },
          {
            key: 'multiple',
            text: 'Multiple'
          }]);
        }, 1000);
      });
    }
  
  

protected async onInit(): Promise<void> {
  this.logDiagnostic('onInit() started. listName=' + String(this.properties.listName || '(none)'));

  this.properties.listName = this.properties.listName && this.properties.listName.trim() !== '' 
    ? this.properties.listName 
    : '';

  this.properties.itemName = this.properties.itemName && this.properties.itemName.trim() !== '' 
    ? this.properties.itemName 
    : '';

  this.properties.itemContent = this.properties.itemContent && this.properties.itemContent.trim() !== '' 
    ? this.properties.itemContent 
    : '';

  this.properties.optionChoice = this.properties.optionChoice && this.properties.optionChoice.trim() !== '' 
    ? this.properties.optionChoice 
       : 'single';

  this.logDiagnostic('onInit() completed. optionChoice=' + this.properties.optionChoice);
  return super.onInit();
}
  
  protected onDispose(): void {
    this.logDiagnostic('onDispose() called. Unmounting component.');
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  
public serialize(): any {
  return {
    ...this.properties,
    searchablePlainTexts: [
      this.properties.listName || '',
      this.properties.itemName || '',
      this.properties.itemContent || '',
      this.properties.optionChoice || ''
    ]
  };
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
              groupName: "Configuration",
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropDownDisabled
                }),
                PropertyPaneDropdown('itemName', {
                  label: 'Header',
                  options: this.items,
                  disabled: this.itemsDropDownDisabled
                }),
                PropertyPaneDropdown('itemContent', {
                  label: 'Content',
                  options: this.contents,
                  disabled: this.contentsDropDownDisabled
                }),
                PropertyPaneDropdown('optionChoice', {
                  label: 'Accordion Style',
                  options: this.options,
                  disabled: this.optionsDropDownDisabled
                })
              ]
            },
            {
              groupName: "Header Styling",
              groupFields: [
                PropertyFieldColorPicker('headerBackgroundColor', {
                  label: 'Header Background Color',
                  selectedColor: this.properties.headerBackgroundColor || '#f0f0f0',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'headerBackgroundColor'
                }),
                PropertyFieldColorPicker('headerTextColor', {
                  label: 'Header Text Color',
                  selectedColor: this.properties.headerTextColor || '#000000',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'headerTextColor'
                }),
                PropertyFieldToggleWithCallout('headerFontBold', {
                  key: 'headerFontBold',
                  label: 'Header Bold Text',
                  checked: !!this.properties.headerFontBold,
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneSlider('headerFontSize', {
                  label: 'Header Font Size',
                  min: 10,
                  max: 72,
                  step: 1,
                  value: Number(this.properties.headerFontSize) || 16,
                  showValue: true
                })
              ]
            },
            {
              groupName: "Content Styling",
              groupFields: [
                PropertyFieldColorPicker('contentBackgroundColor', {
                  label: 'Content Background Color',
                  selectedColor: this.properties.contentBackgroundColor || '#ffffff',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'contentBackgroundColor'
                }),
                PropertyFieldColorPicker('contentTextColor', {
                  label: 'Content Text Color',
                  selectedColor: this.properties.contentTextColor || '#000000',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'contentTextColor'
                }),
                PropertyFieldToggleWithCallout('contentFontBold', {
                  key: 'contentFontBold',
                  label: 'Content Bold Text',
                  checked: !!this.properties.contentFontBold,
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneSlider('contentFontSize', {
                  label: 'Content Font Size',
                  min: 10,
                  max: 72,
                  step: 1,
                  value: Number(this.properties.contentFontSize) || 16,
                  showValue: true
                })
              ]
            },
            {
              groupName: 'Corners',
              groupFields: [
                PropertyPaneDropdown('cornerStyle', {
                  label: 'Corner Style',
                  selectedKey: this.properties.cornerStyle || 'square',
                  options: [
                    { key: 'square', text: 'Square' },
                    { key: 'rounded', text: 'Rounded' }
                  ]
                }),
                PropertyPaneSlider('cornerRadius', {
                  label: 'Rounded Corner Radius',
                  min: 0,
                  max: 40,
                  step: 1,
                  value: Number(this.properties.cornerRadius) || 8,
                  showValue: true,
                  disabled: (this.properties.cornerStyle || 'square') !== 'rounded'
                })
              ]
            },
            {
              groupName: 'Close Bar',
              groupFields: [
                PropertyPaneCheckbox('showCloseBar', {
                  text: 'Show close bar in Single mode',
                  checked: this.properties.showCloseBar !== false
                }),
                PropertyPaneTextField('closeBarText', {
                  label: 'Close Bar Text',
                  value: this.properties.closeBarText || 'Close others \u00d7'
                }),
                PropertyFieldColorPicker('closeBarBackgroundColor', {
                  label: 'Close Bar Background Color',
                  selectedColor: this.properties.closeBarBackgroundColor || '#333333',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'closeBarBackgroundColor'
                }),
                PropertyFieldColorPicker('closeBarTextColor', {
                  label: 'Close Bar Text Color',
                  selectedColor: this.properties.closeBarTextColor || '#ffffff',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'closeBarTextColor'
                }),
                PropertyPaneSlider('closeBarFontSize', {
                  label: 'Close Bar Font Size',
                  min: 10,
                  max: 72,
                  step: 1,
                  value: Number(this.properties.closeBarFontSize) || 12,
                  showValue: true
                }),
                PropertyPaneCheckbox('closeBarFontBold', {
                  text: 'Bold close bar text',
                  checked: this.properties.closeBarFontBold === true
                }),
                PropertyPaneDropdown('closeBarAlignment', {
                  label: 'Close Bar Alignment',
                  selectedKey: this.properties.closeBarAlignment || 'right',
                  options: [
                    { key: 'left', text: 'Left' },
                    { key: 'center', text: 'Center' },
                    { key: 'right', text: 'Right' }
                  ]
                })
              ]
            },
            {
              groupName: "Font Settings",
              groupFields: [
                PropertyFieldDropdownWithCallout('fontFamily', {
                  key: 'fontFamily',
                  label: 'Font Family',
                  selectedKey: this.properties.fontFamily,
                  options: [
                    { key: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif", text: 'Segoe UI' },
                    { key: 'Arial, Helvetica, sans-serif', text: 'Arial' },
                    { key: "'Trebuchet MS', Helvetica, sans-serif", text: 'Trebuchet MS' },
                    { key: "'Georgia', serif", text: 'Georgia' },
                    { key: "'Times New Roman', Times, serif", text: 'Times New Roman' },
                    { key: "'Courier New', Courier, monospace", text: 'Courier New' }
                  ]
                }),
                PropertyFieldDropdownWithCallout('fontStyle', {
                  key: 'fontStyle',
                  label: 'Font Style',
                  selectedKey: this.properties.fontStyle,
                  options: [
                    { key: 'normal', text: 'Normal' },
                    { key: 'italic', text: 'Italic' },
                    { key: 'oblique', text: 'Oblique' }
                  ]
                })
              ]
            },
            {
              groupName: "Advanced:",
              groupFields: [
                PropertyFieldTextWithCallout('overrideCssUrl', {
                  key: 'overrideCssUrl',
                  label: 'Override CSS URL',
                  value: this.properties.overrideCssUrl,
                  placeholder: 'https://contoso.com/styles/accordion-overrides.css',
                  calloutContent: 'Optional stylesheet URL loaded after web part CSS to override component styles.'
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

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }

    console.log('[AccordionWebPart] ' + message);
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
}
