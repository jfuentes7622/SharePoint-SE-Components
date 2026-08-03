import { personnelRecord, fieldInfo} from './SPListRecordTypes';

export interface ISPListService { 
    GetPersonnelData(dir: string, div:string): Promise<Array<personnelRecord>>;
    GetFieldInfo(listName: string, fieldName: string): Promise<fieldInfo>;
}
