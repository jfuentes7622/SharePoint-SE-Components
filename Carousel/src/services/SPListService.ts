
import { ISPListService } from './ISPListService';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPContexts } from '../webparts/carousel/components/SPContexts';
import { carouselSlideRecord } from '../webparts/carousel/components/SPListRecordTypes';
import { ConfigData } from '../webparts/carousel/components/ConfigData';
const LOG_SOURCE: string = '[SPListService] ';

interface IODataListResponse {
  value: carouselSlideRecord[];
}

export default class SPListService implements ISPListService {
  private _spContexts: SPContexts;
  private _configData: ConfigData;

  constructor(configData: ConfigData, spContexts: SPContexts) {
    this._configData = configData;
    this._spContexts = spContexts;
  }

  private logDiagnostic(message: string): void {
    if (this._configData && this._configData.enableDiagnostics === false) {
      return;
    }
    console.log(LOG_SOURCE + message);
  }

  public async GetSlideData(slideList: string): Promise<Array<carouselSlideRecord>> {

    this.logDiagnostic('SlideList in GetSlideData: ' + slideList);

    if (!slideList) {
      return [];
    }

    const now = new Date();
    const nowIso = now.toISOString();
    this.logDiagnostic('Today: ' + nowIso);

    const listTitle = slideList.replace(/'/g, "''");
    const base = `${this._spContexts.absUrl}/_api/web/Lists/GetByTitle('${listTitle}')/items`;
    const baseSelect = '$top=999&$orderby=Id asc&$select=Id,Title,FileRef,FileLeafRef,FSObjType,File/ServerRelativeUrl,File/Name&$expand=File';
    const optionalSelect = '$top=999&$orderby=Id asc&$select=Id,Title,SlideOrder,ClickLink,LinkTarget,FileRef,FileLeafRef,FSObjType,Display,Expiration,StartDate,File/ServerRelativeUrl,File/Name&$expand=File';

    const fetchSlides = (query: string): Promise<carouselSlideRecord[]> => {
      const url = `${base}?${query}`;
      this.logDiagnostic('urlC: ' + url);

      return this._spContexts.spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (!response.ok) {
            this.logDiagnostic('HTTP ' + response.status + ' ' + response.statusText);
            return undefined;
          }
          return response.json().then((json: IODataListResponse) => {
            if (!json || !json.value) {
              this.logDiagnostic('Malformed response JSON');
              return undefined;
            }
            return json.value;
          });
        })
        .catch(() => {
          return undefined;
        });
    };

    // Try the fuller optional-column query first; fall back to the minimal schema
    // if the library is missing the custom columns (SlideOrder/Display/etc.).
    let rawItems: carouselSlideRecord[] = await fetchSlides(optionalSelect);
    if (!rawItems) {
      this.logDiagnostic('Optional slide columns unavailable, retrying with minimal schema.');
      rawItems = await fetchSlides(baseSelect);
    }

    if (!rawItems) {
      return [];
    }

    const imageExtensions: string[] = ['.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg'];

    const filtered = rawItems.filter((item: carouselSlideRecord) => {
      // Files only (exclude folders) when FSObjType is available
      if (typeof item.FSObjType === 'number' && item.FSObjType !== 0) {
        return false;
      }

      const nameOrPath: string = ((item.FileLeafRef || (item.File && item.File.Name) || item.FileRef || '') as string).toLowerCase();
      const isImage: boolean = imageExtensions.some((ext: string) => nameOrPath.indexOf(ext, nameOrPath.length - ext.length) !== -1);
      if (!isImage) {
        return false;
      }

      // Respect Display if present; otherwise include by default
      if (typeof item.Display !== 'undefined') {
        const displayValue = String(item.Display).toLowerCase();
        if (displayValue !== 'yes' && displayValue !== 'true' && displayValue !== '1') {
          return false;
        }
      }

      // Respect date windows only when those columns exist
      if (item.StartDate) {
        const start = new Date(item.StartDate);
        if (!isNaN(start.getTime()) && start > now) {
          return false;
        }
      }

      if (item.Expiration) {
        const expiration = new Date(item.Expiration);
        if (!isNaN(expiration.getTime()) && expiration < now) {
          return false;
        }
      }

      return true;
    });

    const hasSlideOrder: boolean = filtered.some((item: carouselSlideRecord) => typeof item.SlideOrder === 'number');

    const sorted = filtered.sort((a: carouselSlideRecord, b: carouselSlideRecord) => {
      const aId = Number(a.Id) || 0;
      const bId = Number(b.Id) || 0;

      if (hasSlideOrder) {
        const aOrder = typeof a.SlideOrder === 'number' ? a.SlideOrder : Number.MAX_VALUE;
        const bOrder = typeof b.SlideOrder === 'number' ? b.SlideOrder : Number.MAX_VALUE;
        if (aOrder !== bOrder) {
          return aOrder - bOrder;
        }
      }

      // Default behavior when SlideOrder is not present: library item order by ID.
      return aId - bId;
    });

    this.logDiagnostic('Items received: ' + rawItems.length + ', usable image files: ' + sorted.length);
    return sorted;
  }

}



