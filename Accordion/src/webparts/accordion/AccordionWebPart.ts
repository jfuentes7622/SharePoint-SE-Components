import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,IWebPartContext,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  WebPartContext
} from '@microsoft/sp-webpart-base';

import {SPHttpClient,
       SPHttpClientResponse,
      }  from '@microsoft/sp-http';

require('./Accordion.css');

import * as strings from 'AccordionWebPartStrings';
import Accordion from './components/Accordion';
import { IAccordionProps } from './components/IAccordionProps';

export interface IAccordionWebPartProps {
    listName: string;
  itemName: string;
  itemContent: string;
  optionChoice: string;
  spfxContext: WebPartContext;
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

  public render(): void { 
    
    const element: React.ReactElement<IAccordionProps > = React.createElement(
      Accordion,
      {
        listName: this.properties.listName,
        itemName: this.properties.itemName,
        itemContent: this.properties.itemContent,
        optionChoice: this.properties.optionChoice,
        spfxContext:this.context
         }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onPropertyPaneConfigurationStart():void {
    this.listsDropDownDisabled=!this.lists;
    this.itemsDropDownDisabled=!this.properties.listName || !this.items;
    console.info(LOG_SOURCE+ 'List Name:' + this.properties.listName);
    console.info(LOG_SOURCE+'items:' + this.items);
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
  console.info('LOG_SOURCE LoadList url: ' + url);

  return this.fetchLists(url).then((response) => {
    const options: IPropertyPaneDropdownOption[] = [];

    // Define a minimal interface for the list object
    interface ICustomList {
      Id: string;
      Title: string;
    }

    (response.value as ICustomList[]).map((list) => {
      console.info('LOG_SOURCE Found list with title: ' + list.Title);
      options.push({ key: list.Id, text: list.Title });
    });

    return options;
  });
}


    private loadItems(listTitle: string): Promise<IPropertyPaneDropdownOption[]> {
    // giving only single line option for header
     let url=this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetById('"+ listTitle +"')/fields?$select=Title, InternalName&$filter=(((FieldTypeKind eq 2 and FromBaseType eq false) and (Hidden eq false)) or StaticName eq 'Title')";
     console.info (LOG_SOURCE + "Items url: " + url);
     return this.fetchLists(url).then((response) => {
      let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        response.value.map((items:any) => {
        console.info(LOG_SOURCE + "Found Item Title:" + items.Title);
        console.info(LOG_SOURCE+ "Found Internal ColumnName:" + items.InternalName);
        options.push({key: items.InternalName, text: items.Title}); 
    });
    return options;
    });
   }

   private loadContents(listTitle: string): Promise<IPropertyPaneDropdownOption[]> {
     //can be a single line or multiline column 
    let url=this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetById('"+ listTitle +"')/fields?$filter=(((FieldTypeKind eq 3 or FieldTypeKind eq 2) and (FromBaseType eq false) and (Hidden eq false)) or StaticName eq 'Title')";
    console.info (LOG_SOURCE + "contents url: " + url);
    return this.fetchLists(url).then((response) => {
     let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      response.value.map((contents:any) => {
      console.info (LOG_SOURCE + "Found Item Title:" + contents.Title);
      console.info(LOG_SOURCE+"Found Column InternalName:" + contents.InternalName);
      options.push({key: contents.InternalName, text: contents.Title});
      
   });
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

  return super.onInit();
}
  
  protected onDispose(): void {
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
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription + ` v${this.context.manifest.version}`,
          },
          groups: [
            {
              groupName: "Configuration:",
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
            }
          ]
        }
      ]
    };
  }
}
