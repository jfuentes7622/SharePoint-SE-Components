import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IRestResult {
  data: any;
  response: SPHttpClientResponse;
}

export default class SharePointRestClient {
  private webUrl: string;
  private client: SPHttpClient;
  private diagnostics: boolean;

  constructor(webUrl: string, client: SPHttpClient, diagnostics: boolean) {
    this.webUrl = String(webUrl || '').replace(/\/$/, '');
    this.client = client;
    this.diagnostics = diagnostics;
  }

  public setDiagnostics(enabled: boolean): void {
    this.diagnostics = enabled;
  }

  public api(relativeUrl: string): string {
    return this.webUrl + '/_api/' + String(relativeUrl || '').replace(/^\/+/, '');
  }

  public get(relativeUrl: string): Promise<IRestResult> {
    var url = relativeUrl.indexOf('http') === 0 ? relativeUrl : this.api(relativeUrl);
    this.log('GET ' + url);
    return this.client.get(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'OData-Version': '3.0'
      }
    }).then((response: SPHttpClientResponse): Promise<IRestResult> => {
      return this.parse(response, 'GET', url);
    });
  }

  public post(relativeUrl: string, body: any, headers?: { [key: string]: string }): Promise<IRestResult> {
    var url = relativeUrl.indexOf('http') === 0 ? relativeUrl : this.api(relativeUrl);
    var requestHeaders: { [key: string]: string } = {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=verbose',
      'OData-Version': '3.0'
    };
    Object.keys(headers || {}).forEach(function(key: string): void {
      requestHeaders[key] = headers[key];
    });
    this.log('POST ' + url);
    return this.client.post(url, SPHttpClient.configurations.v1, {
      headers: requestHeaders,
      body: body === undefined ? '' : JSON.stringify(body)
    }).then((response: SPHttpClientResponse): Promise<IRestResult> => {
      return this.parse(response, 'POST', url);
    });
  }

  public rows(data: any): any[] {
    if (data && Array.isArray(data.value)) {
      return data.value;
    }
    if (data && data.d && Array.isArray(data.d.results)) {
      return data.d.results;
    }
    return [];
  }

  public item(data: any): any {
    return data && data.d ? data.d : data;
  }

  public escape(value: string): string {
    return String(value || '').replace(/'/g, "''");
  }

  public errorMessage(error: any): string {
    if (!error) {
      return 'Unknown error';
    }
    return String(error.message || error.statusText || error);
  }

  private parse(response: SPHttpClientResponse, method: string, url: string): Promise<IRestResult> {
    this.log(method + ' ' + url + ' -> HTTP ' + String(response.status));
    return response.text().then((text: string): IRestResult => {
      var data: any = {};
      if (text) {
        try {
          data = JSON.parse(text);
        } catch (error) {
          data = { text: text };
        }
      }
      if (!response.ok) {
        var detail = data && data.error && data.error.message
          ? (data.error.message.value || data.error.message)
          : (data && data['odata.error'] && data['odata.error'].message
            ? data['odata.error'].message.value : response.statusText);
        var restError: any = new Error(method + ' ' + url + ' failed with HTTP ' + String(response.status)
          + ': ' + String(detail || 'SharePoint request failed.'));
        restError.status = response.status;
        restError.response = response;
        restError.data = data;
        throw restError;
      }
      return { data: data, response: response };
    });
  }

  private log(message: string): void {
    if (this.diagnostics) {
      console.log('[EventRegistrationRest] ' + message);
    }
  }
}
