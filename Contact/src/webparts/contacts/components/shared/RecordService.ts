import { ISPListService } from './ISPListService';
import SPListService from './SPListService';
import { PersonGroupModel } from './PersonGroup';
import { PersonObjectModel } from './PersonObject';
import { ConfigData } from '../IContactsProps';
import { SPContexts } from './SPContexts';
import { IRecordService } from './IRecordService';

//import PnPTelemetry from "@pnp/telemetry-js";
// import {
//   Logger,
//   ConsoleListener,
//   LogLevel
// } from "@pnp/logging";
import { Logger,  LogLevel } from 'sp-pnp-js';

const LOG_SOURCE: string = 'KM Contacts - ';
//const telemetry = PnPTelemetry.getInstance();

declare global {
  interface Window { 
       iskmActiveLog: boolean;
  }
}

export default class RecordSvc implements IRecordService {
  private _spListService: ISPListService;
  private _configData: ConfigData;


  constructor(configData: ConfigData, spContext: SPContexts) {
    this._configData = configData;
    this._spListService = new SPListService(configData, spContext);

    //Logger.subscribe(ConsoleListener());
    if (!window.iskmActiveLog || window.iskmActiveLog===undefined) {
        Logger.activeLogLevel = LogLevel.Error;
      }
      else {Logger.activeLogLevel=LogLevel.Info;
    }
    
    //telemetry.optOut();

    Logger.write(LOG_SOURCE + 'In RecordService.ts constructor', LogLevel.Info);
  }

 public async GetPersonnel(dir: string, div: string): Promise<Array<PersonGroupModel>> {
    if (this._configData.directorateField)
      {
    const groups = await this.GetSPListFieldChoices(this._configData.personnelListName, 'Group');
    Logger.write('GetPersonnel DIR:' + dir,LogLevel.Info);
    Logger.write('GetPersonnel DIV:' + div, LogLevel.Info);
    const records = await this._spListService.GetPersonnelData(dir, div);
    return groups.map(d => {
      const personList = records
        .filter(d1 => { return d1.group === d; })
        .sort((a, b) => a.groupHeiarchyValue - b.groupHeiarchyValue)
        .map(d1_1 => {
          const retObj = (d1_1 as PersonObjectModel);
          retObj.imageWidth = this._configData.personnelImgWidth;  
          if (this._configData.personnelShowDefaultImg) {
            retObj.imageLink = retObj.imageLink || this._configData.personnelDefaultImgUrl;
          }
          else {
            // We not showing the default image if the record has no image link
            if (!retObj.imageLink) {
              // There's no image; make sure the text doesn't have 75+ pixels of left-padding
              retObj.imageWidth = 1;
            }
          }
          return retObj;
        });
      return personList.length > 0 ? { groupTitle: d, personList: personList } as PersonGroupModel : undefined;
    })
      .filter(f => {
        return f;
      }) as any;
    }
    return [];
  }

  public async GetSPListFieldChoices(listName: string, fieldName: string): Promise<Array<string>> {
    const data = await this._spListService.GetFieldInfo(listName, fieldName);
    return data.Choices || [];
  } 

  //Create personnel like list from the Collection Data Input
   public SortPersonnel(custList: any): Array<PersonGroupModel> {
   //const groups= [...new Set(custList.map((item: { group: any; }) => item.group))];
   const groups: any[] = [];
    for (let i = 0; i < custList.length; i++) {
      const group = custList[i].group;
      if (groups.indexOf(group) === -1) {
        groups.push(group);
      }
    }
   //Logger.write('Groups in sort Personnel' + groups,LogLevel.Info);
    return groups.map(d => {
      const personList = custList
        //.filter((d1: { group: unknown; }) => { return d1.group === d; })
        .filter((d1: { group: any }) => { return d1.group === d; })
        .sort((a: { groupHeiarchyValue: number; }, b: { groupHeiarchyValue: number; }) => a.groupHeiarchyValue - b.groupHeiarchyValue)
        .map((d1_1: { groupHeiarchyValue?: any; Id?: string; Title?: string; name?: string; eMail?: string; phoneNumber?: string; bioLink?: string; imageLink?: string; displayPhoto?: boolean; imageWidth?: number | undefined; sVoip?: string; }) => {
          const retObj = (d1_1 as PersonObjectModel);
          retObj.imageWidth = this._configData.personnelImgWidth;
          retObj.Id=d1_1.groupHeiarchyValue;
          if (this._configData.personnelShowDefaultImg) {
            retObj.imageLink = retObj.imageLink || this._configData.personnelDefaultImgUrl;
          }
          else {
            // We not showing the default image if the record has no image link
            if (!retObj.imageLink) {
              // There's no image; make sure the text doesn't have 75+ pixels of left-padding
              retObj.imageWidth = 1;
            }
          }
          return retObj;
        });
      return personList.length > 0 ? { groupTitle: d, personList: personList } as PersonGroupModel : undefined;
    })
      .filter(f => {
        return f;
      }) as any;
  }

}
