import { ISPListService } from './ISPListService';
import SPListService from './SPListService';

import { SPContexts } from '../webparts/carousel/components/SPContexts';
import { ConfigData } from '../webparts/carousel/components/ConfigData';
import { IRecordService } from './IRecordService';

import SlideItemModel from '../webparts/carousel/components/SlideItem';
import { carouselSlideRecord} from '../webparts/carousel/components/SPListRecordTypes';

const LOG_SOURCE: string = '[RecordSvc] ';

export default class RecordSvc implements IRecordService {
  private _spListService: ISPListService;
  private _spContexts: SPContexts;
  private _configData: ConfigData;

  constructor(configData: ConfigData, spContext: SPContexts) {
    this._spListService = new SPListService(configData, spContext);
    this._spContexts = spContext;
    this._configData = configData;
  }

  private logDiagnostic(message: string): void {
    if (this._configData && this._configData.enableDiagnostics === false) {
      return;
    }
    console.log(LOG_SOURCE + message);
  }

  public async GetSlides(slideListName:string): Promise<Array<SlideItemModel>> {

    this.logDiagnostic('GetSlides called for library: ' + slideListName);

    return this._spListService.GetSlideData(slideListName)
      .then((data: Array<carouselSlideRecord>) => {
          this.logDiagnostic('GetSlideData returned ' + data.length + ' record(s)');
          return data
          .map(d => {
            const rawPath = (d.File && d.File.ServerRelativeUrl) || d.FileRef || '';
            const normalizedPath = rawPath.indexOf('http') === 0 ? rawPath : this._spContexts.aUrl + rawPath;
            const imageUrl = encodeURI(normalizedPath);
            const navUrl = d.LinkTarget || d.ClickLink || '';
            const fallbackOrder = Number(d.Id) || 0;
            const slideOrder = typeof d.SlideOrder === 'number' ? d.SlideOrder : fallbackOrder;

            const retModel = ({
              slideId: d.Id,
              slideText: d.Title || '',
              slideNumber: slideOrder,
              slideImgUrl: imageUrl,
              slideNavigationUrl: navUrl
            } as SlideItemModel);
            return retModel;

          });
      });
  }


}
