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
            const rawTitle: any = this._configData.slideTitleField ? (d as any)[this._configData.slideTitleField] : '';
            const rawDescription: any = this._configData.slideDescriptionField ? (d as any)[this._configData.slideDescriptionField] : '';
            const rawLink: any = this._configData.slideLinkField ? (d as any)[this._configData.slideLinkField] : '';
            const navUrl = typeof rawLink === 'string' ? rawLink : String(rawLink.Url || '');
            const linkText = typeof rawLink === 'string' ? '' : String(rawLink.Description || '');
            const slideTitle = typeof rawTitle === 'string' ? rawTitle : String((rawTitle && (rawTitle.Description || rawTitle.Url)) || '');
            const slideDescription = typeof rawDescription === 'string' ? rawDescription : String((rawDescription && (rawDescription.Description || rawDescription.Url)) || '');
            const fallbackOrder = Number(d.Id) || 0;
            const slideOrder = typeof d.SlideOrder === 'number' ? d.SlideOrder : fallbackOrder;

            const retModel = ({
              slideId: d.Id,
              slideTitle: slideTitle,
              slideText: slideDescription,
              slideNumber: slideOrder,
              slideImgUrl: imageUrl,
              slideNavigationUrl: navUrl,
              slideLinkText: linkText
            } as SlideItemModel);
            return retModel;

          });
      });
  }


}
