import { WebPartContext } from "@microsoft/sp-webpart-base";
import RecordSvc from "./shared/RecordService";

export type ConfigData = {
  personnelListName: string;
  personnelImgWidth: number;
  personnelDefaultImgUrl: string;
  personnelShowDefaultImg: boolean;
  siteUrl:string;
  directorateField:string;
  divisionField:string;
  groupField:string;
  branchField: string;
  groupHeiarchyField: string;
  nameField: string;
  eMailField: string;
  phoneNumberField: string;
  bioLinkField: string;
  displayPhotoField: string;
  imageLinkField: string;
  sVoipField : string;
  imageIsCircle: boolean;
};

export interface IContactsProps {
  directorate: string;
  division: string;
  propertyCheckbox: boolean;
  recSvc: RecordSvc;
  customList: IRecord[];
  spfxContext: WebPartContext;
  imageWidth: number;
  ContactsListId: string;
  directorateField:string;
  divisionField:string;
  groupField:string;
  branchField: string;
  groupHeiarchyField: string;
  nameField: string;
  eMailField: string;
  phoneNumberField: string;
  bioLinkField: string;
  displayPhotoField: string;
  imageLinkField: string;
  sVoipField : string;
}

export interface IRecord{
   title:string;
   name: string;
   email: string;
   phoneNumber: string;                                   
   bioLink: string;
   imageLink: string; 
   displayPhoto: boolean;
   imageWidth: number;
   sVoip: string;
   group: string;
   groupHeiarchyValue: number; 
}

