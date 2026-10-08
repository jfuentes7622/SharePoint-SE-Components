
import { ISPListService } from './ISPListService';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPContexts } from '../webparts/carousel/components/SPContexts';
import { carouselSlideRecord } from '../webparts/carousel/components/SPListRecordTypes';
import { ConfigData } from '../webparts/carousel/components/ConfigData';
const LOG_SOURCE: string = '[SPListService] ';

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

  private extractResults(json: any): carouselSlideRecord[] {
    if (json && json.value) {
      return json.value;
    }
    if (json && json.d && json.d.results) {
      return json.d.results;
    }
    return undefined;
  }

  private buildViewXml(viewData: any, fieldNames: string[]): string {
    const query: string = String(viewData && viewData.ViewQuery || '');
    const scopeValue: string = String(viewData && viewData.Scope !== undefined ? viewData.Scope : '').toLowerCase();
    const scopes: { [key: string]: string } = {
      '1': 'Recursive',
      '2': 'RecursiveAll',
      '3': 'FilesOnly',
      'recursive': 'Recursive',
      'recursiveall': 'RecursiveAll',
      'filesonly': 'FilesOnly'
    };
    const scope: string = scopes[scopeValue] ? ' Scope="' + scopes[scopeValue] + '"' : '';
    const fields: string = fieldNames.filter((fieldName: string): boolean => {
      return /^[A-Za-z_][A-Za-z0-9_]*$/.test(fieldName);
    }).map((fieldName: string): string => '<FieldRef Name="' + fieldName + '" />').join('');
    const rowLimit: number = parseInt(String(viewData && viewData.RowLimit || ''), 10);
    const rowLimitXml: string = !isNaN(rowLimit) && rowLimit > 0 ? '<RowLimit>' + rowLimit + '</RowLimit>' : '';
    return '<View' + scope + '><Query>' + query + '</Query><ViewFields>' + fields + '</ViewFields>' + rowLimitXml + '</View>';
  }

  private async loadSelectedViewXml(listTitle: string, fieldNames: string[]): Promise<string> {
    const viewId: string = String(this._configData.slideViewId || '').replace(/[{}]/g, '');
    if (!viewId || viewId === '__all__') {
      return '';
    }
    const viewBase = `${this._spContexts.absUrl}/_api/web/Lists/GetByTitle('${listTitle}')/views`;
    const urls: string[] = [
      `${viewBase}/getById('${encodeURIComponent(viewId)}')?$select=ViewQuery,RowLimit,Scope`,
      `${viewBase}(guid'${encodeURIComponent(viewId)}')?$select=ViewQuery,RowLimit,Scope`
    ];
    for (let index = 0; index < urls.length; index += 1) {
      this.logDiagnostic('REST request: GET ' + urls[index]);
      const response: SPHttpClientResponse = await this._spContexts.spHttpClient.get(urls[index], SPHttpClient.configurations.v1);
      this.logDiagnostic('REST response: GET ' + urls[index] + ' -> HTTP '
        + String(response.status) + ' ' + response.statusText);
      if (!response.ok) {
        continue;
      }
      const json: any = await response.json();
      const viewData: any = json && json.d ? json.d : json;
      this.logDiagnostic('REST payload summary: GET ' + urls[index] + ' -> keys='
        + Object.keys(viewData || {}).slice(0, 20).join(','));
      const viewXml: string = this.buildViewXml(viewData, fieldNames);
      this.logDiagnostic('Selected view loaded. HasFilter=' + String(/<Where(?:\s|>)/i.test(viewXml)));
      return viewXml;
    }
    throw new Error('Unable to load the selected SharePoint view.');
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
    const configuredFields = [
      this._configData.slideTitleField,
      this._configData.slideDescriptionField,
      this._configData.slideLinkField
    ].filter((fieldName: string, index: number, fields: string[]): boolean => {
      return !!fieldName && /^[A-Za-z_][A-Za-z0-9_]*$/.test(fieldName) && fields.indexOf(fieldName) === index;
    });
    const requiredFields = ['Id', 'FileRef', 'FileLeafRef', 'FSObjType', 'File/ServerRelativeUrl', 'File/Name'].concat(configuredFields);
    const optionalFields = requiredFields.concat(['SlideOrder', 'Display', 'Expiration', 'StartDate']);
    const baseSelect = '$top=999&$orderby=Id asc&$select=' + requiredFields.join(',') + '&$expand=File';
    const optionalSelect = '$top=999&$orderby=Id asc&$select=' + optionalFields.join(',') + '&$expand=File';

    const viewXml: string = await this.loadSelectedViewXml(listTitle, optionalFields);
    const fetchSlides = (query: string, selectedViewXml: string): Promise<carouselSlideRecord[]> => {
      const viewQuery: string = query.replace(/(?:^|&)\$top=[^&]*/g, '')
        .replace(/(?:^|&)\$orderby=[^&]*/g, '')
        .replace(/^&/, '');
      const url = selectedViewXml
        ? base.replace(/\/items$/, '/GetItems') + '?' + viewQuery
        : `${base}?${query}`;
      this.logDiagnostic('REST request: ' + (selectedViewXml ? 'POST ' : 'GET ') + url);

      const postSelectedView = async (): Promise<SPHttpClientResponse> => {
        const acceptValues: string[] = [
          'application/json',
          'application/json;odata=verbose',
          'application/json;odata=nometadata'
        ];
        let response: SPHttpClientResponse;
        for (let index: number = 0; index < acceptValues.length; index += 1) {
          if (index > 0) {
            this.logDiagnostic('REST retry: POST ' + url + ' with Accept=' + acceptValues[index]);
          }
          response = await this._spContexts.spHttpClient.post(url, SPHttpClient.configurations.v1, {
            headers: {
              'Accept': acceptValues[index],
              'Content-Type': 'application/json;odata=verbose'
            },
            body: JSON.stringify({
              query: {
                'ViewXml': selectedViewXml
              }
            })
          });
          if (response.status !== 406) {
            return response;
          }
          this.logDiagnostic('REST response: POST ' + url + ' with Accept=' + acceptValues[index]
            + ' -> HTTP ' + String(response.status) + ' ' + response.statusText);
        }
        return response;
      };

      const request: Promise<SPHttpClientResponse> = selectedViewXml
        ? postSelectedView()
        : this._spContexts.spHttpClient.get(url, SPHttpClient.configurations.v1);

      return request
        .then((response: SPHttpClientResponse) => {
          this.logDiagnostic('REST response: ' + (selectedViewXml ? 'POST ' : 'GET ') + url + ' -> HTTP '
            + String(response.status) + ' ' + response.statusText);
          if (!response.ok) {
            if (selectedViewXml) {
              return response.clone().text().then((responseText: string) => {
                console.error(LOG_SOURCE + 'Selected-view GetItems POST failed. HTTP '
                  + String(response.status) + ' ' + response.statusText + ', url=' + url
                  + (responseText ? ', response=' + responseText.substring(0, 1000) : ''));
                return undefined;
              });
            }
            return undefined;
          }
          return response.json().then((json: any) => {
            const results: carouselSlideRecord[] = this.extractResults(json);
            if (!results) {
              this.logDiagnostic('Malformed response JSON');
              return undefined;
            }
            this.logDiagnostic('REST payload summary: ' + (selectedViewXml ? 'POST ' : 'GET ') + url
              + ' -> items=' + String(results.length));
            return results;
          });
        })
        .catch((error: any) => {
          console.error(LOG_SOURCE + 'REST request failed: ' + (selectedViewXml ? 'POST ' : 'GET ')
            + url + '. Error=' + String(error && error.message ? error.message : error));
          return undefined;
        });
    };

    // Try the fuller optional-column query first; fall back to the minimal schema
    // if the library is missing the custom columns (SlideOrder/Display/etc.).
    let rawItems: carouselSlideRecord[] = await fetchSlides(optionalSelect, viewXml);
    if (!rawItems) {
      this.logDiagnostic('Optional control columns unavailable, retrying with selected caption columns only.');
      const baseViewXml: string = viewXml ? await this.loadSelectedViewXml(listTitle, requiredFields) : '';
      rawItems = await fetchSlides(baseSelect, baseViewXml);
    }

    if (!rawItems) {
      return [];
    }

    const imageExtensions: string[] = ['.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg'];
    let folderCount: number = 0;
    let nonImageCount: number = 0;
    let hiddenCount: number = 0;
    let futureCount: number = 0;
    let expiredCount: number = 0;

    const filtered = rawItems.filter((item: carouselSlideRecord) => {
      // Files only (exclude folders) when FSObjType is available
      if (typeof item.FSObjType === 'number' && item.FSObjType !== 0) {
        folderCount += 1;
        return false;
      }

      const nameOrPath: string = ((item.FileLeafRef || (item.File && item.File.Name) || item.FileRef || '') as string).toLowerCase();
      const isImage: boolean = imageExtensions.some((ext: string) => nameOrPath.indexOf(ext, nameOrPath.length - ext.length) !== -1);
      if (!isImage) {
        nonImageCount += 1;
        return false;
      }

      // SharePoint returns null for blank optional columns. Only a meaningful
      // Display value should override the default behavior of showing the slide.
      if (typeof item.Display === 'string' || typeof item.Display === 'boolean' || typeof item.Display === 'number') {
        const displayValue = String(item.Display).toLowerCase().trim();
        if (displayValue !== 'yes' && displayValue !== 'true' && displayValue !== '1') {
          if (displayValue) {
            hiddenCount += 1;
            return false;
          }
        }
      }

      // Respect date windows only when those columns exist
      if (item.StartDate) {
        const start = new Date(item.StartDate);
        if (!isNaN(start.getTime()) && start > now) {
          futureCount += 1;
          return false;
        }
      }

      if (item.Expiration) {
        const expiration = new Date(item.Expiration);
        if (!isNaN(expiration.getTime()) && expiration < now) {
          expiredCount += 1;
          return false;
        }
      }

      return true;
    });

    if (nonImageCount > 0 && rawItems.length > 0) {
      this.logDiagnostic('First returned item keys: ' + Object.keys(rawItems[0] || {}).slice(0, 30).join(','));
    }

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

    this.logDiagnostic('Items received: ' + rawItems.length + ', usable image files: ' + sorted.length
      + ', excluded folders=' + folderCount + ', non-images=' + nonImageCount + ', hidden=' + hiddenCount
      + ', future=' + futureCount + ', expired=' + expiredCount);
    return sorted;
  }

}
