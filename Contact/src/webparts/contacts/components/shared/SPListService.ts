/*
The types exposed here need to resemble what SharePoint would deliver after a REST call
for the respective lists... don't try to convert between models here
*/
import { ISPListService } from './ISPListService';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPContexts } from './SPContexts';
import { personnelRecord, fieldInfo} from './SPListRecordTypes';
import { ConfigData } from '../IContactsProps';
//import PnPTelemetry from "@pnp/telemetry-js";
// import {
//   Logger,
//   ConsoleListener,
//   LogLevel
// } from "@pnp/logging";
import  { Logger, LogLevel } from 'sp-pnp-js';

const LOG_SOURCE: string = 'KM Contacts - ';
//const telemetry = PnPTelemetry.getInstance();

declare global {
  interface Window { 
       iskmActiveLog: boolean;
  }
}

export default class SPListService implements ISPListService {
  private _spContexts: SPContexts;
  private _configData: ConfigData;

  constructor(configData: ConfigData, spContexts: SPContexts) {
    this._configData = configData;
    this._spContexts = spContexts;

    //Logger.subscribe(ConsoleListener());
    if (!window.iskmActiveLog || window.iskmActiveLog===undefined) {
        Logger.activeLogLevel = LogLevel.Error;
      }
      else {Logger.activeLogLevel=LogLevel.Info;
    }
    
   // telemetry.optOut();

    Logger.write(LOG_SOURCE + 'In SPListService.ts constructor', LogLevel.Info);
  }

 public async GetPersonnelData(dir:string, div:string): Promise<Array<personnelRecord>> {  
  //Logger.write('absUrl:' + this._spContexts.absUrl,LogLevel.Info);
  const url=`${this._spContexts.absUrl}/_api/web/Lists/GetById('${this._configData.personnelListName}')/items?$top=999`;
  const personalRec:Array<personnelRecord> = new Array<personnelRecord>();
  Logger.write('GetPersonnelData URL:' + url,LogLevel.Info);
  Logger.write('GetPersonnelData Dir:'+ dir,LogLevel.Info);
  Logger.write('GetPersonnelData Div:' + div,LogLevel.Info);
    if (this._configData.directorateField){
      return new Promise<Array<personnelRecord>>((resolve) => {
        this._spContexts.spHttpClient.get(url, SPHttpClient.configurations.v1)    
        .then((response: SPHttpClientResponse) => {
            response.json()
              .then((responseJSON: any) => { return responseJSON.value; })
              .then((responseJSON: any) => {
                const config:ConfigData = this._configData;
                responseJSON.filter((f: any) => {
                  return (f[config.directorateField] || '').toLowerCase() === (dir || '').toLowerCase() &&
                    (f[config.divisionField] || '').toLowerCase() === (div || '').toLowerCase() ;         
                }).forEach((rec:any)=>{
                  //map data
                  personalRec.push({Id: rec.Id,
                    Title: rec.Title,
                    directorate: rec[config.directorateField],
                    division: rec[config.divisionField],
                    branch: rec[config.branchField],
                    group: rec[config.groupField],
                    groupHeiarchyValue: rec[config.groupHeiarchyField],
                    name: rec[config.nameField],
                    eMail: rec[config.eMailField],
                    phoneNumber: rec[config.phoneNumberField],
                    bioLink: rec[config.bioLinkField],
                    displayPhoto: rec[config.displayPhotoField],
                    imageLink: rec[config.imageLinkField],
                    sVoip : rec[config.sVoipField],
                  imageIsCircle: config.imageIsCircle});
              });
                resolve(personalRec);
              });
          })
          .catch(e => {
            resolve([]);
          });
      });
    }
    return personalRec;
  }

  public async GetFieldInfo(listName: string, fieldName: string): Promise<fieldInfo> {    
    Logger.write('absUrl:'+ this._spContexts.absUrl,LogLevel.Info);
    const url=`${this._spContexts.absUrl}/_api/web/Lists/GetById('${listName}')/fields/GetByTitle('${fieldName}')`;
    Logger.write('GetFieldInfo Url:'+ url,LogLevel.Info);
    return new Promise<fieldInfo | void>((resolve) => {
      this._spContexts.spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then(response => {
          response.json()
            .then((responseJSON: any) => {
              return responseJSON;
            })
            .then((data: fieldInfo) => {
              resolve(data);
            });
        })
        .catch(e => {
          resolve();
        });
    }) as any;
  }

}


