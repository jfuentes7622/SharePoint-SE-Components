import { PersonGroupModel } from './PersonGroup';

export interface IRecordService {
    GetPersonnel(dir:string, div:string): Promise<Array<PersonGroupModel>>;
   // SortPersonnel(cList: any):  Promise<Array<PersonGroupModel>>;
    SortPersonnel(cList: any):  Array<PersonGroupModel>;
    GetSPListFieldChoices(listName: string, fieldName: string): Promise<Array<string>>;
}