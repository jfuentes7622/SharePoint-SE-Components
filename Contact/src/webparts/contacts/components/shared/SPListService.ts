/*
The types exposed here need to resemble what SharePoint would deliver after a REST call
for the respective lists... don't try to convert between models here
*/
import { ISPListService } from './ISPListService';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPContexts } from './SPContexts';
import { personnelRecord, fieldInfo} from './SPListRecordTypes';
import { ConfigData } from '../IContactsProps';

const LOG_SOURCE: string = '[SPListService] ';

export default class SPListService implements ISPListService {
  private _spContexts: SPContexts;
  private _configData: ConfigData;

  constructor(configData: ConfigData, spContexts: SPContexts) {
    this._configData = configData;
    this._spContexts = spContexts;

    this.logDiagnostic('In SPListService.ts constructor');
  }

  private logDiagnostic(message: string): void {
    if (this._configData && this._configData.enableDiagnostics === false) {
      return;
    }
    console.log(LOG_SOURCE + message);
  }

 public async GetPersonnelData(dir:string, div:string): Promise<Array<personnelRecord>> {  
  const url=`${this._spContexts.absUrl}/_api/web/Lists/GetById('${this._configData.personnelListName}')/items?$top=999`;
  const personalRec:Array<personnelRecord> = new Array<personnelRecord>();
  this.logDiagnostic('GetPersonnelData URL:' + url);
  this.logDiagnostic('GetPersonnelData Dir:'+ dir);
  this.logDiagnostic('GetPersonnelData Div:' + div);
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
                    Title: rec[config.titleField || 'Title'],
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
                  imageShape: config.imageShape});
              });
                resolve(personalRec);
              });
          })
          .catch(e => {
            console.error(LOG_SOURCE + 'GetPersonnelData failed: ' + e);
            resolve([]);
          });
      });
    }
    return personalRec;
  }

  public async GetFieldInfo(listName: string, fieldName: string): Promise<fieldInfo> {    
    const url=`${this._spContexts.absUrl}/_api/web/Lists/GetById('${listName}')/fields/GetByTitle('${fieldName}')`;
    this.logDiagnostic('GetFieldInfo Url:'+ url);
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
          console.error(LOG_SOURCE + 'GetFieldInfo failed: ' + e);
          resolve();
        });
    }) as any;
  }

}


