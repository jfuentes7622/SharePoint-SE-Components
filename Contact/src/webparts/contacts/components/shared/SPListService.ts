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
        this.logDiagnostic('REST request: GET ' + url);
        this._spContexts.spHttpClient.get(url, SPHttpClient.configurations.v1)    
        .then((response: SPHttpClientResponse) => {
            this.logDiagnostic('REST response: GET ' + url + ' -> HTTP '
              + String(response.status) + ' ' + response.statusText);
            if (!response.ok) {
              throw new Error('HTTP ' + String(response.status) + ' ' + response.statusText);
            }
            response.json()
              .then((responseJSON: any) => { return responseJSON.value; })
              .then((responseJSON: any) => {
                this.logDiagnostic('REST payload summary: GET ' + url + ' -> items='
                  + String(Array.isArray(responseJSON) ? responseJSON.length : 0));
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
                this.logDiagnostic('GetPersonnelData filtered result count=' + String(personalRec.length));
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
    const escapedListName = String(listName || '').replace(/'/g, "''");
    const escapedFieldName = String(fieldName || '').replace(/'/g, "''");
    const url=`${this._spContexts.absUrl}/_api/web/Lists/GetById('${escapedListName}')/fields/GetByInternalNameOrTitle('${escapedFieldName}')`;
    this.logDiagnostic('GetFieldInfo Url:'+ url);
    this.logDiagnostic('REST request: GET ' + url);
    const response: SPHttpClientResponse = await this._spContexts.spHttpClient.get(url, SPHttpClient.configurations.v1);
    this.logDiagnostic('REST response: GET ' + url + ' -> HTTP '
      + String(response.status) + ' ' + response.statusText);
    if (!response.ok) {
      const responseText = await response.text();
      throw new Error('GetFieldInfo failed. HTTP ' + String(response.status) + ' ' + response.statusText
        + ', url=' + url + ', response=' + responseText);
    }
    const responseJSON: any = await response.json();
    const data: fieldInfo = responseJSON && responseJSON.d ? responseJSON.d : responseJSON;
    this.logDiagnostic('REST payload summary: GET ' + url + ' -> keys='
      + Object.keys(data || {}).slice(0, 20).join(','));
    if (!data) {
      throw new Error('GetFieldInfo returned an empty response for field ' + fieldName + '.');
    }
    return data;
  }

}
