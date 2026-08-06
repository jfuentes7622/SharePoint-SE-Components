declare interface IContactsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListConfigGroupName: string;
  ListChoicesGroupName: string;
  CommonGroupName: string;
  HeaderStyleGroupName: string;
  TitleStyleGroupName: string;
  PropEnableDiagnosticsLabel: string;
  DirectorateFieldLabel: string;
  DivisionFieldLabel: string;
  EucomListUrl: string;
  PersonnelImgWidth: string;
  overRideCSS: string;
  
  //CustomList labels
  CustomListAddItem: string;
  CustomListBack: string;
  CustomListTrue: string;
  CustomListFalse: string;
  CustomListOK: string;
  CustomListCancel: string;
  CustomListEdit: string;
  CustomListDel: string;
  CustomListYes: string;
  CustomListNo: string;
  CustomListConfirmDel: string;
  CustomListConfirmDelMssg: string;
  CustomListFieldMissing: string;

  //PicturePicker labels
  PicturePickerTitle: string;
  PicturePickerRecent: string;
  PicturePickerSite: string;
  PicturePickerButtonSelect: string;
  PicturePickerButtonReset: string;

  //DocumentPicker Labels
  DocumentPickerTitle: string;
  DocumentPickerRecent: string;
  DocumentPickerSite: string;
  DocumentPickerButtonSelect: string;
  DocumentPickerButtonReset: string;
}


declare module 'ContactsWebPartStrings' {
  const strings: IContactsWebPartStrings;
  export = strings;
}
