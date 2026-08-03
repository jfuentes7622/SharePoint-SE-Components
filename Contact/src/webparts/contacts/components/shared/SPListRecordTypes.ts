export type personnelRecord = {
    Id: string;
    Title: string;
    directorate: string;
    division: string;
    branch: string;
    group: string;
    groupHeiarchyValue: number;
    name: string;
    eMail: string;
    phoneNumber: string;
    bioLink: string;
    displayPhoto: boolean;
    imageLink: string;
    sVoip : string;
    imageIsCircle: boolean;
  };
  
  export type fieldInfo = {
    Id: string;
    Title: string;
    FieldTypeKind: number;
    InternalName: string;
    Required: boolean;
    SchemaXml: string;
    Choices?: Array<string>;
  };
  