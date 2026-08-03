
import { ISPListService } from './ISPListService';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPContexts } from '../webparts/carousel/components/SPContexts';
import { carouselSlideRecord } from '../webparts/carousel/components/SPListRecordTypes';
import { ConfigData } from '../webparts/carousel/components/ConfigData';
const LOG_SOURCE: string='KM Carousel- ';

// import {
//   Logger,
//   ConsoleListener,
//   LogLevel
// } from "@pnp/logging";
import { Logger, LogLevel } from 'sp-pnp-js';

export default class SPListService implements ISPListService {
  private _spContexts: SPContexts;
 // private _configData: ConfigData;

  constructor(configData: ConfigData, spContexts: SPContexts) {
    //this._configData = configData;
    this._spContexts = spContexts;
  }

  public async GetSlideData(slideList:string): Promise<Array<carouselSlideRecord>> {
      
      //Logger.subscribe(new ConsoleListener());
      if (!window.iskmActiveLog || window.iskmActiveLog===undefined) {
      Logger.activeLogLevel = LogLevel.Error;
      }
      else {Logger.activeLogLevel=LogLevel.Info;}

      Logger.write(LOG_SOURCE + 'SlideList in GetSlideData:' + slideList, LogLevel.Info);
      if (slideList !== undefined) {   
      const today = new Date();
      const compToday=today.toISOString();
     
      Logger.write(LOG_SOURCE + 'Today:' + compToday, LogLevel.Info);
      const urlC = `${this._spContexts.absUrl}/_api/web/Lists/GetByTitle('${slideList}')/items?$top=999&$select=*,FileRef/FileRef&$filter=(Display eq 'Yes') and (Expiration ge '` + compToday + `') and (StartDate le '` + compToday + `')`; 
   
      Logger.write(LOG_SOURCE +'urlC:' + urlC, LogLevel.Info);

      return new Promise<Array<carouselSlideRecord>>((resolve) => {
        this._spContexts.spHttpClient.get(urlC, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
            response.json()
              .then((responseJSON: any) => {
                
                Logger.write(LOG_SOURCE + `responseJSON:` + responseJSON.value, LogLevel.Info);
                return responseJSON.value;
              })
              .then((responseJSON: Array<carouselSlideRecord>) => {
                resolve(responseJSON);
              });
          })
          .catch(e => {
            resolve([]);
          });
      });
    }
    return([]);
 } 

}



