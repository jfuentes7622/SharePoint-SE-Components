import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import {
  SPHttpClient,
  // SPHttpClientResponse
} from '@microsoft/sp-http';

//import {sp} from'@pnp/sp';
//import { IField } from "@pnp/sp/fields/types";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists"
// import "@pnp/sp/fields";

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneCheckbox,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneLabel,
  IPropertyPaneGroup,
  BaseClientSideWebPart
  //PropertyPaneLink,
} from '@microsoft/sp-webpart-base';

import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

//import { CustomFilePicker } from './components/CustomFilePicker';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
//(SPO) import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
//(spo) import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';
//(SPO)import { IColumnReturnProperty, PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';


//import { PropertyFieldColorPickerMini } from 'sp-client-custom-fields/lib/PropertyFieldColorPickerMini';

//import { PropertyFieldSPListQuery,PropertyFieldSPListQueryOrderBy,IPropertyFieldSPListQueryProps} from 'sp-client-custom-fields/lib/PropertyFieldSPListQuery'

import * as strings from 'ContactsWebPartStrings';

import { IContactsProps, IRecord } from './components/IContactsProps';

import ContactImages from './components/ContactImages';

import RecordService from './components/shared/RecordService';
import { SPContexts } from './components/shared/SPContexts';
import { ConfigData } from './components/shared/ConfigData';


// import { sp } from '@pnp/sp';
// import '@pnp/sp/webs';
// import '@pnp/sp/lists';
// import "@pnp/sp/items";
//import '@pnp/sp/fields';


import pnp from 'sp-pnp-js';

const packageSolutionConfig: any = require('../../../config/package-solution.json');
const documentPickerModule: any = require('sp-client-custom-fields/lib/PropertyFieldDocumentPickerHost');
const PropertyFieldDocumentPickerHost: any = documentPickerModule.default || documentPickerModule;
const picturePickerModule: any = require('sp-client-custom-fields/lib/PropertyFieldPicturePickerHost');
const PropertyFieldPicturePickerHost: any = picturePickerModule.default || picturePickerModule;

export interface IContactsWebPartProps {
  overridecss: string;
  directorate: string;
  division: string;
  directorateField: string;
  divisionField: string;
  groupField: string;
  propertyCheckbox: boolean;
  recSvc: RecordService;
  customList: IRecord[];
  spfxContext: WebPartContext;
  PersonnelPanelHeaderBackColor: string;
  PersonnelPanelHeaderTextColor: string;
  ContactsListId: string;
  listSourceMode: string;
  provisionListName: string;
  imageWidth: number;
  personnelDefaultImgUrl: string;
  personnelShowDefaultImg: boolean;
  description: string;
  toggleInfoHeaderValue: boolean;
  idField: string;
  titleField: string;
  branchField: string;
  groupHeiarchyField: string;
  nameField: string;
  eMailField: string;
  phoneNumberField: string;
  bioLinkField: string;
  displayPhotoField: string;
  imageLinkField: string;
  sVoipField: string;
  imageShape: string;
  headerFontFamily: string;
  headerFontStyle: string;
  headerFontBold: boolean;
  headerAlignment: string;
  headerTopCorners: string;
  tileTitleColor: string;
  tileInfoBackgroundColor: string;
  titleFontFamily: string;
  titleFontStyle: string;
  titleFontBold: boolean;
  enableDiagnostics: boolean;
}

const LOG_SOURCE: string = '[ContactsWebPart] ';

require('../contacts/assets/Contacts.css');

export default class ContactsImagesWebPart extends BaseClientSideWebPart<IContactsWebPartProps> {

  private dirs: IPropertyPaneDropdownOption[];
  private dirsDropdownDisabled: boolean = true;
  private divs: IPropertyPaneDropdownOption[];
  private divsDropdownDisabled: boolean = true;
  private provisionListStatus: string = '';
  private isProvisioningList: boolean = false;

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }
    console.log(LOG_SOURCE + message);
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


  private siteUrl() {
    const getUrl = window.location;
    const aUrl = getUrl.protocol + "//" + getUrl.host + "/";
    const rel = decodeURI(this.context.pageContext.web.serverRelativeUrl);
    return aUrl + rel;
  }

  
// private async loadListColumns(listId: string): Promise<IPropertyPaneDropdownOption[]> {
//   const fields = await pnp.sp.web.lists.getById(listId).fields
//     .filter("Hidden eq false and ReadOnlyField eq false")
//     .get();

//   return fields.map(f => ({
//     key: f.InternalName,
//     text: f.Title || f.InternalName
//   }));
// }


  public render(): void {
    this.logDiagnostic('render() called');

    this.domElement.style.setProperty('--PersonnelPanelHeaderBackColor', (this.properties.PersonnelPanelHeaderBackColor !== undefined ? this.properties.PersonnelPanelHeaderBackColor : "#8A1717"));
    this.domElement.style.setProperty('--PersonnelPanelHeaderTextColor', (this.properties.PersonnelPanelHeaderTextColor !== undefined ? this.properties.PersonnelPanelHeaderTextColor : "#FAFAFA"));
    this.domElement.style.setProperty('--PersonnelPanelHeaderFontFamily', (this.properties.headerFontFamily || 'Segoe UI'));
    this.domElement.style.setProperty('--PersonnelPanelHeaderFontStyle', (this.properties.headerFontStyle || 'normal'));
    this.domElement.style.setProperty('--PersonnelPanelHeaderFontWeight', (this.properties.headerFontBold ? 'bold' : 'normal'));
    this.domElement.style.setProperty('--PersonnelPanelHeaderTextAlign', (this.properties.headerAlignment || 'left'));
    this.domElement.style.setProperty('--PersonnelPanelHeaderTopRadius', (this.properties.headerTopCorners === 'squared' ? '0' : '8px'));
    this.domElement.style.setProperty('--PersonnelTileTitleColor', (this.properties.tileTitleColor || 'rgb(51, 51, 51)'));
    this.domElement.style.setProperty('--PersonnelTileInfoBackgroundColor', (this.properties.tileInfoBackgroundColor || 'transparent'));
    this.domElement.style.setProperty('--PersonnelTileTitleFontFamily', (this.properties.titleFontFamily || 'Segoe UI'));
    this.domElement.style.setProperty('--PersonnelTileTitleFontStyle', (this.properties.titleFontStyle || 'normal'));
    this.domElement.style.setProperty('--PersonnelTileTitleFontWeight', (this.properties.titleFontBold === false ? 'normal' : 'bold'));

    if (this.properties.overridecss) {
      // inject the  master style sheet
      const head: HTMLElement = document.getElementsByTagName('head')[0] || document.documentElement;
      const customStyle: HTMLLinkElement = document.createElement('link');
      customStyle.href = this.properties.overridecss;
      customStyle.rel = 'stylesheet';
      customStyle.type = 'text/css';
      head.insertAdjacentElement('beforeend', customStyle);
    }
    const element: React.ReactElement<IContactsProps> = React.createElement(

      ContactImages,
      {
        directorate: this.properties.directorate,
        division: this.properties.division,
        propertyCheckbox: this.properties.propertyCheckbox,
        recSvc: this.GetNewRecSvc(),
        customList: this.getNormalizedCustomList(),
        spfxContext: this.context,
        imageWidth: this.properties.imageWidth,
        ContactsListId: this.properties.ContactsListId,
        directorateField: this.properties.directorateField,
        divisionField: this.properties.divisionField,
        groupField: this.properties.groupField,
        branchField: this.properties.branchField,
        groupHeiarchyField: this.properties.groupHeiarchyField,
        nameField: this.properties.nameField,
        eMailField: this.properties.eMailField,
        phoneNumberField: this.properties.phoneNumberField,
        bioLinkField: this.properties.bioLinkField,
        displayPhotoField: this.properties.displayPhotoField,
        imageLinkField: this.properties.imageLinkField,
        sVoipField: this.properties.sVoipField,
        enableDiagnostics: this.properties.enableDiagnostics
      }

    );
    ReactDom.render(element, this.domElement);
  }

  private getNormalizedCustomList(): IRecord[] {
    const contacts: any[] = this.properties.customList || [];

    return contacts.map((contact: any) => {
      const displayPhoto = contact.displayPhoto === true || contact.displayPhoto === 'true' || contact.displayPhoto === 1 || contact.displayPhoto === '1';
      const displayOrder = Number(contact.groupHeiarchyValue);

      return {
        ...contact,
        displayPhoto: displayPhoto,
        groupHeiarchyValue: isNaN(displayOrder) ? 0 : displayOrder
      };
    }) as IRecord[];
  }

  private renderBiographyPicker(field: any, value: any,
    onUpdate: (fieldId: string, fieldValue: any) => void): React.ReactElement<any> {
    return this.renderCollectionFilePicker(
      PropertyFieldDocumentPickerHost,
      field,
      value,
      onUpdate,
      '.doc,.docx,.ppt,.pptx,.xls,.xlsx,.pdf,.txt',
      false
    );
  }

  private renderImagePicker(field: any, value: any,
    onUpdate: (fieldId: string, fieldValue: any) => void): React.ReactElement<any> {
    return this.renderCollectionFilePicker(
      PropertyFieldPicturePickerHost,
      field,
      value,
      onUpdate,
      '.gif,.jpg,.jpeg,.bmp,.png,.svg,.webp',
      true
    );
  }

  private renderCollectionFilePicker(picker: any, field: any, value: any,
    onUpdate: (fieldId: string, fieldValue: any) => void,
    allowedFileExtensions: string, previewImage: boolean): React.ReactElement<any> {
    const properties: any = {};
    properties[field.id] = String(value || '');

    return React.createElement(picker, {
      key: 'contactFilePicker-' + field.id + '-' + String(value || ''),
      targetProperty: field.id,
      label: '',
      initialValue: String(value || ''),
      context: this.context,
      previewImage: previewImage,
      previewDocument: !previewImage,
      allowedFileExtensions: allowedFileExtensions,
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


  protected onPropertyPaneConfigurationStart(): void {
    this.dirsDropdownDisabled = !this.dirs;
    this.divsDropdownDisabled = !this.properties.directorate || !this.divs;

    if (this.properties.propertyCheckbox && this.properties.ContactsListId && this.columnOptions.length === 0) {
      this.loadListColumns(this.properties.ContactsListId).then((options: IPropertyPaneDropdownOption[]): void => {
        this.columnOptions = options;
        this.context.propertyPane.refresh();
      });
    }

    if (this.dirs || !this.properties.propertyCheckbox || !this.properties.ContactsListId ||
      !this.properties.directorateField || !this.properties.divisionField) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'dirs');

    this.loadDirectorates()
      .then((dirOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
        this.dirs = dirOptions;
        this.dirsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        return this.loadDivisions(this.properties.directorate);
      })
      .then((divOptions: IPropertyPaneDropdownOption[]): void => {
        this.divs = divOptions;
        if (divOptions.length > 1) { this.divsDropdownDisabled = false; } else { this.divsDropdownDisabled = true; }

        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {

    if (propertyPath === 'customList' && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.logDiagnostic('onPropertyPaneFieldChanged customList: ' + this.properties.customList);
      this.context.propertyPane.refresh();
      this.render();
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }

    if (propertyPath === 'propertyCheckbox') {
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'listSourceMode' && oldValue !== newValue) {
      this.properties.ContactsListId = '';
      this.columnOptions = [];
      this.dirs = undefined;
      this.divs = undefined;
      this.properties.directorate = '';
      this.properties.division = '';
      this.provisionListStatus = '';
      this.context.propertyPane.refresh();
      this.render();
    }

    if (propertyPath === 'ContactsListId') {
      this.columnOptions = newValue ? await this.loadListColumns(String(newValue)) : [];
      this.dirs = undefined;
      this.divs = undefined;
      this.properties.directorate = '';
      this.properties.division = '';
      this.context.propertyPane.refresh();
      this.render();
    }

    //if (propertyPath === 'ContactsListId' && newValue) {
    // this.properties.directorateField = (await this.GetInternalName("directorate"))!.toString();
    // this.properties.divisionField = (await this.GetInternalName("division"))!.toString();
    // this.properties.groupField = (await this.GetInternalName("Group"))!.toString();
    // this.properties.bioLinkField = (await this.GetInternalName("bioLink"))!.toString();
    // this.properties.branchField = (await this.GetInternalName("branch"))!.toString();
    // this.properties.groupHeiarchyField = (await this.GetInternalName("groupHeiarchyValue"))!.toString();
    // this.properties.nameField = (await this.GetInternalName("name"))!.toString();
    // this.properties.eMailField = (await this.GetInternalName("eMail"))!.toString();
    // this.properties.phoneNumberField = (await this.GetInternalName("phoneNumber"))!.toString();
    // this.properties.displayPhotoField = (await this.GetInternalName("displayPhoto"))!.toString();
    // this.properties.imageLinkField = (await this.GetInternalName("imageLink"))!.toString();
    // this.properties.sVoipField = (await this.GetInternalName("sVoip"))!.toString();


    //if(oldValue !== newValue){

    if (propertyPath === 'directorateField' && newValue || propertyPath === 'divisionField' && newValue) {
      this.loadDirectorates()
        .then((dirOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
          this.dirs = dirOptions;
          this.dirsDropdownDisabled = false;
          this.context.propertyPane.refresh();
          return this.loadDivisions(this.properties.directorate);
        })
        .then((divOptions: IPropertyPaneDropdownOption[]): void => {
          this.divs = divOptions;
          this.divsDropdownDisabled = true;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        });
    }
    //}

    //uncomment for color changes on the back panel
    /* if (propertyPath === 'PersonnelPanelHeaderBackColor' && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);        
      this.render();        
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
          
    if (propertyPath === 'PersonnelPanelHeaderTextColor' && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.render();     
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
*/
    if (propertyPath === 'directorate' && newValue) {

      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // get previously selected item
      const previousItem: string = this.properties.division;
      // reset selected item
      this.properties.division = '';
      // push new item value
      this.onPropertyPaneFieldChanged('division', previousItem, this.properties.division);
      // disable item selector until new items are loaded
      this.divsDropdownDisabled = false;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'divs');

      this.loadDivisions(this.properties.directorate)
        .then((divOptions: IPropertyPaneDropdownOption[]): void => {
          // store items
          this.divs = divOptions;
          // enable item selector
          this.divsDropdownDisabled = false;
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        });
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  //   private GetInternalName(displayName:string)  {
  //     // return  List.GetById(this.properties.ContactsListId).fields.getByTitle(displayName)().then(function (field:any) {
  //     //   if (field) { return field.InternalName; } else return "";
  //     // });
  //       const url=  this.siteUrl() + "/_api/web/lists/GetById('"+ this.properties.ContactsListId +"')/fields?$select=internalname,title";
  //     if (this.properties.ContactsListId !== undefined)
  //     {
  //      Logger.write("fields url: " + url,LogLevel.Info);
  //      return this.fetchLists(url).then((response) => {
  //         let fieldname:string = "";
  //          response.value.forEach(function (field:any){
  //             if (field.Title.toLowerCase()=== displayName.toLocaleLowerCase())
  //               {
  //                 Logger.write("Found inernal name for :" +displayName+ " as: "+field.InternalName,LogLevel.Info);
  //                 fieldname =  field.InternalName;
  //               }
  //         });
  //         if(fieldname) {return fieldname}else return "";
  //         });
  //   };
  // }

  private isKeyInDropdownOptionsAsync(
  options: IPropertyPaneDropdownOption[],
  targetKey: string
): boolean {
  for (let i = 0; i < options.length; i++) {
    if (options[i].key === targetKey) {
      return true;
    }
  }
  return false;
}

  private async loadDirectorates(): Promise<IPropertyPaneDropdownOption[]> {
    const options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();

    if (this.properties.ContactsListId !== undefined) {
      const url = "/_api/web/lists/GetById('" + this.properties.ContactsListId + "')/items?$top=1000&$select=" + (this.properties.directorateField) + "&$Orderby=" + (this.properties.directorateField) + " asc'";

      this.logDiagnostic("Items url: " + url);
      return this.fetchLists(url).then((response) => {
        if (response) {
          response.value.map((direct: any) => {
            this.logDiagnostic("Found Dir:" + direct[this.properties.directorateField]);
            if (direct[this.properties.directorateField]) {
              
              const targetKey = direct[this.properties.directorateField];
              const exists = this.isKeyInDropdownOptionsAsync(options, targetKey);

              //if (options.findIndex(key => key.key === direct[this.properties.directorateField]) === -1) {
              if (!exists) {
                options.push({ key: direct[this.properties.directorateField], text: direct[this.properties.directorateField] });
              }
            }
          });
        }
        if (options.length === 0) { options.push({ key: 0, text: "Selected List Not Configured for Directorates" }); }
        return options;
      });
    }
    if (options.length === 0) { options.push({ key: 0, text: "Selected List Not Configured for Directorates" }); }
    return options;
  }


  private async loadDivisions(dir: string): Promise<IPropertyPaneDropdownOption[]> {
    const options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
    if (this.properties.ContactsListId !== undefined) {
      const url = "/_api/web/lists/GetById('" + this.properties.ContactsListId + "')/items?$top=1000&$select=" + (this.properties.divisionField) + "&$filter=" + (this.properties.directorateField) + " eq'" + dir + "'&$Orderby=" + (this.properties.divisionField) + " asc'";
      this.logDiagnostic("Items url: " + url);
      const response = await this.fetchLists(url);
      if (response) {
        response.value.map((direct: any) => {
          this.logDiagnostic("Found div:" + direct[this.properties.divisionField]);
          if (direct[this.properties.divisionField]) {
            const targetKey = direct[this.properties.divisionField];
            const exists = this.isKeyInDropdownOptionsAsync(options, targetKey);
            if (!exists) {
            //if (options.findIndex(key => key.key === direct[this.properties.divisionField]) === -1) {
              options.push({ key: direct[this.properties.divisionField], text: direct[this.properties.divisionField] });
            }
          }
        });
      }
      if (options.length === 0) { options.push({ key: 0, text: "Selected List Not Configured for Divisions" }); }
      return options;
    }
    if (options.length === 0) { options.push({ key: 0, text: "Selected List Not Configured for Divisions" }); }
    return options;
  }


  private async fetchLists(url: string): Promise<any> {
    const response = await this.context.spHttpClient.get((this.siteUrl() + url), SPHttpClient.configurations.v1);
    if (response.ok) {
      return response.json();
    }
    else {
      console.error(LOG_SOURCE + "Failed to get url:" + url + ". Error=" + response.statusText);
      return undefined;
    }
  }

  private async loadListColumns(listId: string): Promise<IPropertyPaneDropdownOption[]> {
    if (!listId) {
      return [];
    }
    const url = "/_api/web/lists/GetById('" + listId + "')/fields?$select=Title,InternalName,Hidden,ReadOnlyField&$filter=Hidden eq false";
    const response = await this.fetchLists(url);
    const options: IPropertyPaneDropdownOption[] = [];
    if (response && response.value) {
      response.value.forEach((field: any): void => {
        if (field.InternalName && field.ReadOnlyField !== true) {
          options.push({
            key: field.InternalName,
            text: String(field.Title || field.InternalName) + ' (' + String(field.InternalName) + ')'
          });
        }
      });
    }
    options.sort((left: IPropertyPaneDropdownOption, right: IPropertyPaneDropdownOption): number => {
      return String(left.text).localeCompare(String(right.text));
    });
    return options;
  }

  private GetNewRecSvc(): RecordService {
    // const getUrl = window.location;
    // const aUrl = getUrl.protocol + "//" + getUrl.host + "/" ;
    return new RecordService(
      <ConfigData>{
        personnelListName: this.properties.ContactsListId,
        siteUrl: decodeURI(this.context.pageContext.web.absoluteUrl),
        personnelImgWidth: this.properties.imageWidth,
        personnelDefaultImgUrl: this.properties.personnelDefaultImgUrl,
        personnelShowDefaultImg: this.properties.personnelShowDefaultImg,
        directorateField: this.properties.directorateField,
        divisionField: this.properties.divisionField,
        titleField: this.properties.titleField,
        groupField: this.properties.groupField,
        branchField: this.properties.branchField,
        groupHeiarchyField: this.properties.groupHeiarchyField,
        nameField: this.properties.nameField,
        eMailField: this.properties.eMailField,
        phoneNumberField: this.properties.phoneNumberField,
        bioLinkField: this.properties.bioLinkField,
        displayPhotoField: this.properties.displayPhotoField,
        imageLinkField: this.properties.imageLinkField,
        sVoipField: this.properties.sVoipField,
        imageShape: this.properties.imageShape,
        enableDiagnostics: this.properties.enableDiagnostics

      },
      <SPContexts>{
        spHttpClient: this.context.spHttpClient,
        //absUrl: decodeURI(this.context.pageContext.web.absoluteUrl),
        //absUrl: aUrl,
        absUrl: this.siteUrl(),
        //absUrl: this.context.pageContext.web.absoluteUrl.replace(this.context.pageContext.web.serverRelativeUrl, ""),
        relUrl: decodeURI(this.context.pageContext.web.serverRelativeUrl)
      }
    );
  }

  private async provisionContactList(): Promise<void> {
    const listTitle = String(this.properties.provisionListName || '').trim();
    if (!listTitle) {
      this.provisionListStatus = 'Enter a list name before creating the list.';
      this.context.propertyPane.refresh();
      return;
    }

    this.isProvisioningList = true;
    this.provisionListStatus = 'Creating list and contact fields...';
    this.context.propertyPane.refresh();

    try {
      const result: any = await pnp.sp.web.lists.ensure(
        listTitle,
        'Contact directory data for the SPS Contacts web part.',
        100,
        false
      );
      const list: any = result.list;
      const listData: any = await list.select('Id').get();
      const listId = String(listData.Id || listData.ID || listData.id || '').replace(/[{}]/g, '');

      await list.fields.getByInternalNameOrTitle('Title').update({ Title: 'Job Title' });
      await this.ensureContactField(list, 'directorate', '<Field Type="Text" Name="directorate" StaticName="directorate" DisplayName="Directorate" MaxLength="255" Group="SPS Contacts" />');
      await this.ensureContactField(list, 'division', '<Field Type="Text" Name="division" StaticName="division" DisplayName="Division" MaxLength="255" Group="SPS Contacts" />');
      await this.ensureContactField(list, 'branch', '<Field Type="Text" Name="branch" StaticName="branch" DisplayName="Branch" MaxLength="255" Group="SPS Contacts" />');
      await this.ensureContactField(list, 'Group', '<Field Type="Choice" Name="Group" StaticName="Group" DisplayName="Group" FillInChoice="TRUE" Format="Dropdown" Group="SPS Contacts"><CHOICES><CHOICE>Group 1</CHOICE><CHOICE>Group 2</CHOICE><CHOICE>Group 3</CHOICE></CHOICES></Field>');
      await this.ensureContactField(list, 'groupHeiarchyValue', '<Field Type="Number" Name="groupHeiarchyValue" StaticName="groupHeiarchyValue" DisplayName="Display Order" Min="0" Decimals="0" Group="SPS Contacts" />');
      await this.ensureContactField(list, 'name', '<Field Type="Text" Name="name" StaticName="name" DisplayName="Person Name" MaxLength="255" Group="SPS Contacts" />');
      await this.ensureContactField(list, 'eMail', '<Field Type="Text" Name="eMail" StaticName="eMail" DisplayName="Email" MaxLength="255" Group="SPS Contacts" />');
      await this.ensureContactField(list, 'phoneNumber', '<Field Type="Text" Name="phoneNumber" StaticName="phoneNumber" DisplayName="Phone Number" MaxLength="255" Group="SPS Contacts" />');
      await this.ensureContactField(list, 'bioLink', '<Field Type="Note" Name="bioLink" StaticName="bioLink" DisplayName="Biography Link" NumLines="3" RichText="FALSE" Group="SPS Contacts" />');
      await this.ensureContactField(list, 'displayPhoto', '<Field Type="Boolean" Name="displayPhoto" StaticName="displayPhoto" DisplayName="Display Photo" Group="SPS Contacts"><Default>1</Default></Field>');
      await this.ensureContactField(list, 'imageLink', '<Field Type="Note" Name="imageLink" StaticName="imageLink" DisplayName="Image Link" NumLines="3" RichText="FALSE" Group="SPS Contacts" />');
      await this.ensureContactField(list, 'sVoip', '<Field Type="Text" Name="sVoip" StaticName="sVoip" DisplayName="VOIP" MaxLength="255" Group="SPS Contacts" />');

      if (!listId) {
        throw new Error('SharePoint created the list but did not return its ID. Select it manually from the list picker.');
      }

      this.properties.propertyCheckbox = true;
      this.properties.listSourceMode = 'create';
      this.properties.ContactsListId = listId;
      this.properties.idField = 'Id';
      this.properties.titleField = 'Title';
      this.properties.directorateField = 'directorate';
      this.properties.divisionField = 'division';
      this.properties.branchField = 'branch';
      this.properties.groupField = 'Group';
      this.properties.groupHeiarchyField = 'groupHeiarchyValue';
      this.properties.nameField = 'name';
      this.properties.eMailField = 'eMail';
      this.properties.phoneNumberField = 'phoneNumber';
      this.properties.bioLinkField = 'bioLink';
      this.properties.displayPhotoField = 'displayPhoto';
      this.properties.imageLinkField = 'imageLink';
      this.properties.sVoipField = 'sVoip';
      this.properties.directorate = '';
      this.properties.division = '';
      this.dirs = undefined;
      this.divs = undefined;
      this.columnOptions = await this.loadListColumns(listId);
      this.provisionListStatus = 'Created "' + listTitle + '" and connected it to this web part.';
      this.render();
    } catch (error) {
      const message = error && (error as any).message ? (error as any).message : String(error);
      this.provisionListStatus = 'Unable to create the contact list: ' + message;
      console.error(LOG_SOURCE + this.provisionListStatus, error);
    } finally {
      this.isProvisioningList = false;
      this.context.propertyPane.refresh();
    }
  }

  private async ensureContactField(list: any, internalName: string, schemaXml: string): Promise<void> {
    try {
      await list.fields.getByInternalNameOrTitle(internalName).select('Id').get();
      return;
    } catch (fieldNotFoundError) {
      try {
        await list.fields.createFieldAsXml(schemaXml);
      } catch (createError) {
        const message = createError && (createError as any).message
          ? (createError as any).message
          : String(createError);
        throw new Error('Failed to create field "' + internalName + '": ' + message);
      }
    }
  }

  
public onInit(): Promise<void> {
  this.logDiagnostic('onInit called');
  pnp.setup({ spfxContext: this.context });   // bind SPFx context for auth/headers
  return Promise.resolve();
}


  protected onDispose(): void {
    this.logDiagnostic('onDispose called');
    ReactDom.unmountComponentAtNode(this.domElement);
  }

/*   protected get dataVersion(): Version {
    return Version.parse('1.0');
  } */

    private columnOptions: IPropertyPaneDropdownOption[] = [];

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let customListControl: any = [];
    let splist: any;
    let directorateControl: any = [];
    let divisionControl: any = [];
    //let linkControl: any=[];
    let colorHeaderControl: any = [];
    let colorTextControl: any = [];
    let imgWidthControl: any = [];
    let infoControl: any = [];

    const fontFamilyOptions: IPropertyPaneDropdownOption[] = [
      { key: 'Arial, sans-serif', text: 'Arial' },
      { key: '"Segoe UI", sans-serif', text: 'Segoe UI' },
      { key: 'Calibri, sans-serif', text: 'Calibri' },
      { key: 'Cambria, serif', text: 'Cambria' },
      { key: 'Georgia, serif', text: 'Georgia' },
      { key: 'Tahoma, sans-serif', text: 'Tahoma' },
      { key: '"Times New Roman", serif', text: 'Times New Roman' },
      { key: '"Trebuchet MS", sans-serif', text: 'Trebuchet MS' },
      { key: 'Verdana, sans-serif', text: 'Verdana' },
      { key: '"Comic Sans MS", cursive', text: 'Comic Sans MS' }
    ];
    const fontStyleOptions: IPropertyPaneDropdownOption[] = [
      { key: 'normal', text: 'Normal' },
      { key: 'italic', text: 'Italic' },
      { key: 'oblique', text: 'Oblique' }
    ];
    const alignmentOptions: IPropertyPaneDropdownOption[] = [
      { key: 'left', text: 'Left' },
      { key: 'center', text: 'Center' },
      { key: 'right', text: 'Right' }
    ];
    const topCornersOptions: IPropertyPaneDropdownOption[] = [
      { key: 'rounded', text: 'Rounded' },
      { key: 'squared', text: 'Squared' }
    ];
    const imageShapeOptions: IPropertyPaneDropdownOption[] = [
      { key: 'square', text: 'Square' },
      { key: 'rounded', text: 'Rounded' },
      { key: 'circle', text: 'Circle' }
    ];

    const conditionalGroupFields: IPropertyPaneGroup["groupFields"] = [
      PropertyPaneCheckbox("propertyCheckbox", {
        text: 'Use a List',
        checked: this.properties.propertyCheckbox === true,
        disabled: false
      })
    ];

    const configGroupFields: IPropertyPaneGroup["groupFields"] = [];

    // if (this.properties.propertyCheckbox===false) {

    
 customListControl = PropertyFieldCollectionData('customList', {
  label: 'Enter Contacts data with BIO Links and Pictures',
  panelHeader: 'Contacts',
  manageBtnLabel: 'Manage Contacts',
  key: 'customListFieldId',
  disabled: false,
  enableSorting: true,
  fields: [
    { id: 'Title', title: 'Job Title', required: true, type: CustomCollectionFieldType.string },
    { id: 'name', title: 'Full Name and Rank', required: true, type: CustomCollectionFieldType.string },
    { id: 'email', title: 'EMail', required: true, type: CustomCollectionFieldType.string },
    { id: 'phoneNumber', title: 'Phone Number', required: true, type: CustomCollectionFieldType.string },
    {
      id: 'bioLink',
      title: 'Link to Biography',
      required: false,
      type: CustomCollectionFieldType.custom,
      onCustomRender: this.renderBiographyPicker.bind(this)
    },
    {
      id: 'imageLink',
      title: 'Link to Image',
      required: false,
      type: CustomCollectionFieldType.custom,
      onCustomRender: this.renderImagePicker.bind(this)
    },
    { id: 'displayPhoto', title: 'Display Photo?', required: false, type: CustomCollectionFieldType.boolean, defaultValue: false },
    { id: 'sVoip', title: 'VOIP', required: false, type: CustomCollectionFieldType.string },
    { id: 'group', title: 'Group', required: true, type: CustomCollectionFieldType.string },
    { id: 'groupHeiarchyValue', title: 'Display Order', required: true, type: CustomCollectionFieldType.number }
  ],
  value: this.getNormalizedCustomList()
});


    //} else {
    splist = PropertyFieldListPicker('ContactsListId', {
      label: 'Select a list',
      selectedList: this.properties.ContactsListId,
      includeHidden: false,
      orderBy: PropertyFieldListPickerOrderBy.Title,
      disabled: false,
      onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
      properties: this.properties,
      context: this.context,
      deferredValidationTime: 0,
      multiSelect: false,
      key: 'listPickerFieldId'
    });
    directorateControl = PropertyPaneDropdown('directorate', {
      label: strings.DirectorateFieldLabel,
      options: this.dirs,
      disabled: this.dirsDropdownDisabled
    });
    divisionControl = PropertyPaneDropdown('division', {
      label: strings.DivisionFieldLabel,
      options: this.divs,
      disabled: this.divsDropdownDisabled
    });
    const imageShapeControl = PropertyPaneDropdown('imageShape', {
      label: 'Image Shape',
      options: imageShapeOptions,
      selectedKey: this.properties.imageShape || 'square'
    });
    //};
    colorHeaderControl = PropertyFieldColorPicker('PersonnelPanelHeaderBackColor', {
      label: "Panel Header Background Color",
      selectedColor: this.properties.PersonnelPanelHeaderBackColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      alphaSliderHidden: false,
      style: PropertyFieldColorPickerStyle.Full,
      iconName: 'HeaderBackColor',
      properties: this.properties,
      key: 'PersonnelPanelHeaderBackColorField'
    });
    colorTextControl = PropertyFieldColorPicker('PersonnelPanelHeaderTextColor', {
      label: "Panel Header Text Color",
      selectedColor: this.properties.PersonnelPanelHeaderTextColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      style: PropertyFieldColorPickerStyle.Full,
      iconName: 'HeaderTextColor',
      properties: this.properties,
      key: 'PersonnelPanelHeaderTextColorField'
    });
    imgWidthControl = PropertyPaneTextField('imageWidth', {
      label: "Image Diameter (ex:150, no px or %)"
    });

    const headerFontFamilyControl = PropertyPaneDropdown('headerFontFamily', {
      label: 'Header Font Family',
      options: fontFamilyOptions,
      selectedKey: this.properties.headerFontFamily || 'Segoe UI'
    });
    const headerFontStyleControl = PropertyPaneDropdown('headerFontStyle', {
      label: 'Header Font Style',
      options: fontStyleOptions,
      selectedKey: this.properties.headerFontStyle || 'normal'
    });
    const headerBoldControl = PropertyPaneCheckbox('headerFontBold', {
      text: 'Bold Header Text',
      checked: this.properties.headerFontBold === true
    });
    const headerAlignmentControl = PropertyPaneDropdown('headerAlignment', {
      label: 'Header Text Alignment',
      options: alignmentOptions,
      selectedKey: this.properties.headerAlignment || 'left'
    });
    const headerTopCornersControl = PropertyPaneDropdown('headerTopCorners', {
      label: 'Header Top Corners',
      options: topCornersOptions,
      selectedKey: this.properties.headerTopCorners || 'rounded'
    });
    const tileTitleColorControl = PropertyFieldColorPicker('tileTitleColor', {
      label: 'Tile Title Color',
      selectedColor: this.properties.tileTitleColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      alphaSliderHidden: false,
      style: PropertyFieldColorPickerStyle.Full,
      iconName: 'Font',
      properties: this.properties,
      key: 'tileTitleColorField'
    });
    const tileInfoBackgroundColorControl = PropertyFieldColorPicker('tileInfoBackgroundColor', {
      label: 'Tile Info Background Color',
      selectedColor: this.properties.tileInfoBackgroundColor,
      onPropertyChange: this.onPropertyPaneFieldChanged,
      alphaSliderHidden: false,
      style: PropertyFieldColorPickerStyle.Full,
      iconName: 'Color',
      properties: this.properties,
      key: 'tileInfoBackgroundColorField'
    });
    const titleFontFamilyControl = PropertyPaneDropdown('titleFontFamily', {
      label: 'Title Font Family',
      options: fontFamilyOptions,
      selectedKey: this.properties.titleFontFamily || 'Segoe UI'
    });
    const titleFontStyleControl = PropertyPaneDropdown('titleFontStyle', {
      label: 'Title Font Style',
      options: fontStyleOptions,
      selectedKey: this.properties.titleFontStyle || 'normal'
    });
    const titleBoldControl = PropertyPaneCheckbox('titleFontBold', {
      text: 'Bold Title Text',
      checked: this.properties.titleFontBold !== false
    });
    const diagnosticsControl = PropertyPaneCheckbox('enableDiagnostics', {
      text: strings.PropEnableDiagnosticsLabel,
      checked: this.properties.enableDiagnostics !== false
    });
    // linkControl = PropertyPaneLink('Link', {
    //   text: 'CLICK HERE TO ADD/CHANGE EUCOM CONTACTS DATA', href: strings.EucomListUrl, target:'_blank'                  
    // });
    // infoControl = PropertyPaneWebPartInformation({
    //   description: `<b>Lists Requires the following structure:</b><br><ul>
    //   <li>Title: string;</li>
    //   <li>directorate: string;</li>
    //   <li>division: string;</li>
    //   <li>branch: string;</li>
    //   <li>group: choice;</li>
    //   <li>groupHeiarchyValue: number;</li>
    //   <li>name: string;</li>
    //   <li>eMail: string;</li>
    //   <li>phoneNumber: string;</li>
    //   <li>bioLink: string;</li>
    //   <li>displayPhoto: boolean;</li>
    //   <li>imageLink: string;</li>
    //   <li>sVoip : string;</li>
    //   </ul>`,
    //   moreInfoLink: `https://pnp.github.io/sp-dev-fx-property-controls/`,
    //   /* videoProperties: {
    //     embedLink: `https://www.youtube.com/embed/d_9o3tQ90zo`,
    //     properties: { allowFullScreen: true}
    //   }, */
    //   key: 'webPartInfoId'
    // });  

    
infoControl = PropertyPaneLabel('webPartInfoId', {
  text: 'Lists Requires the Following Fields:\n' +
        '- Job Title\n' +
        '- Directorate\n' +
        '- Division\n' +
        '- Branch\n' +
        '- Group By\n' +
        '- Order By (Number)\n' +
        '- Person Name\n' +
        '- Email\n' +
        '- Phone Number\n' +
        '- Bio Link\n' +
        '- Display Photo\n' +
        '- Image Link\n' +
        '- sVoip'
});




    if (this.properties.propertyCheckbox) {
      conditionalGroupFields.push(PropertyPaneDropdown('listSourceMode', {
        label: 'List Source',
        selectedKey: this.properties.listSourceMode || 'existing',
        options: [
          { key: 'existing', text: 'Use an Existing List' },
          { key: 'create', text: 'Create a New List' }
        ]
      }));

      if ((this.properties.listSourceMode || 'existing') === 'existing') {
        conditionalGroupFields.push(splist);
      } else {
        conditionalGroupFields.push(
          PropertyPaneTextField('provisionListName', {
            label: 'New Contact List Name',
            placeholder: 'Contacts'
          }),
          PropertyPaneButton('provisionContactList', {
            text: this.isProvisioningList ? 'Creating...' : 'Create Contact List',
            buttonType: PropertyPaneButtonType.Primary,
            disabled: this.isProvisioningList,
            onClick: this.provisionContactList.bind(this)
          }),
          PropertyPaneLabel('provisionContactListStatus', {
            text: this.provisionListStatus || 'Creates the required columns and connects the new list automatically. Requires permission to manage lists.'
          })
        );
      }

      if (this.properties.ContactsListId) {
        configGroupFields.push(
          PropertyPaneDropdown('titleField', { label: 'Job Title Column', options: this.columnOptions, selectedKey: this.properties.titleField }),
          PropertyPaneDropdown('directorateField', { label: 'Directorate Column', options: this.columnOptions, selectedKey: this.properties.directorateField }),
          PropertyPaneDropdown('divisionField', { label: 'Division Column', options: this.columnOptions, selectedKey: this.properties.divisionField }),
          PropertyPaneDropdown('branchField', { label: 'Branch Column', options: this.columnOptions, selectedKey: this.properties.branchField }),
          PropertyPaneDropdown('groupField', { label: 'Group Column', options: this.columnOptions, selectedKey: this.properties.groupField }),
          PropertyPaneDropdown('groupHeiarchyField', { label: 'Display Order Column', options: this.columnOptions, selectedKey: this.properties.groupHeiarchyField }),
          PropertyPaneDropdown('nameField', { label: 'Person Name Column', options: this.columnOptions, selectedKey: this.properties.nameField }),
          PropertyPaneDropdown('eMailField', { label: 'Email Column', options: this.columnOptions, selectedKey: this.properties.eMailField }),
          PropertyPaneDropdown('phoneNumberField', { label: 'Phone Number Column', options: this.columnOptions, selectedKey: this.properties.phoneNumberField }),
          PropertyPaneDropdown('bioLinkField', { label: 'Biography Link Column', options: this.columnOptions, selectedKey: this.properties.bioLinkField }),
          PropertyPaneDropdown('displayPhotoField', { label: 'Display Photo Column', options: this.columnOptions, selectedKey: this.properties.displayPhotoField }),
          PropertyPaneDropdown('imageLinkField', { label: 'Image Link Column', options: this.columnOptions, selectedKey: this.properties.imageLinkField }),
          PropertyPaneDropdown('sVoipField', { label: 'VOIP Column', options: this.columnOptions, selectedKey: this.properties.sVoipField })
        );
      } else {
        configGroupFields.push(PropertyPaneLabel('selectListBeforeColumns', {
          text: 'Select an existing list or create a new list to configure columns.'
        }));
      }
    } else {
      conditionalGroupFields.push(customListControl);
    }
    //non conditional
    // conditionalGroupFields.push(
    //   colorHeaderControl,
    //   colorTextControl,
    //   imgWidthControl
    // );

    const propertyPaneGroups: IPropertyPaneGroup[] = [
      {
        groupName: 'Version: ' + this.getWebPartVersion(),
        groupFields: [
          PropertyPaneLabel('propertyPaneVersionInfo', {
            text: ' '
          })
        ]
      },
      {
        groupName: strings.BasicGroupName,
        groupFields: conditionalGroupFields
      }
    ];

    if (this.properties.propertyCheckbox) {
      propertyPaneGroups.push(
        {
          groupName: strings.ListConfigGroupName,
          groupFields: configGroupFields
        },
        {
          groupName: strings.ListChoicesGroupName,
          groupFields: [
            directorateControl,
            divisionControl
          ]
        }
      );
    }

    propertyPaneGroups.push(
      {
              groupName: strings.HeaderStyleGroupName,
              groupFields: [
                colorHeaderControl,
                colorTextControl,
                headerFontFamilyControl,
                headerFontStyleControl,
                headerBoldControl,
                headerAlignmentControl,
                headerTopCornersControl
              ]
          },
          {
              groupName: strings.TitleStyleGroupName,
              groupFields: [
                tileTitleColorControl,
                titleFontFamilyControl,
                titleFontStyleControl,
                titleBoldControl,
                tileInfoBackgroundColorControl,
                imgWidthControl,
                imageShapeControl,
                PropertyPaneTextField('overridecss', {
                  label: strings.overRideCSS,
                  validateOnFocusOut: true,
                }),
              ]
          },
          {
              groupName: 'Diagnostics',
              groupFields: [
                diagnosticsControl
              ]
      }
    );

    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: ''
          },
          groups: propertyPaneGroups
        }
      ]
    };
  }
}
