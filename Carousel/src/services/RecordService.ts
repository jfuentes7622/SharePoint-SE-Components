import { ISPListService } from './ISPListService';
import SPListService from './SPListService';

import { SPContexts } from '../webparts/carousel/components/SPContexts';
import { ConfigData } from '../webparts/carousel/components/ConfigData';
import { IRecordService } from './IRecordService';

import SlideItemModel from '../webparts/carousel/components/SlideItem';
import { carouselSlideRecord} from '../webparts/carousel/components/SPListRecordTypes';

export default class RecordSvc implements IRecordService {
  private _spListService: ISPListService;
  private _spContexts: SPContexts;

  constructor(configData: ConfigData, spContext: SPContexts) {
    this._spListService = new SPListService(configData, spContext);
    this._spContexts = spContext;
  }


  public async GetSlides(slideListName:string): Promise<Array<SlideItemModel>> {
    
    return this._spListService.GetSlideData(slideListName)
      .then((data: Array<carouselSlideRecord>) => {
          return data
          .sort((a, b) => { return a.SlideOrder - b.SlideOrder; })
          .map(d => {

            const retModel = ({
              slideId: d.Id,
              slideText:d.Title,
              slideNumber:d.SlideOrder,
              slideImgUrl: this._spContexts.aUrl + d.FileRef,
              slideNavigationUrl:d.LinkTarget
            } as SlideItemModel);
            return retModel;

          });
      });
  }


}
