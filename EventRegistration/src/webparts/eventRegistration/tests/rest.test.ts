import { assert } from 'chai';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import SharePointRestClient from '../services/SharePointRestClient';

describe('Event Registration REST protocol', function(): void {
  var rest: SharePointRestClient;
  var requestedUrl: string;
  var requestOptions: ISPHttpClientOptions;
  var status: number;
  var responseBody: string;

  beforeEach(function(): void {
    status = 200;
    responseBody = '{"value":[]}';
    var send = function(url: string, configuration: any, options: ISPHttpClientOptions): Promise<SPHttpClientResponse> {
      requestedUrl = url;
      requestOptions = options;
      return Promise.resolve({
        ok: status >= 200 && status < 300,
        status: status,
        statusText: 'Bad Request',
        text: function(): Promise<string> { return Promise.resolve(responseBody); }
      } as SPHttpClientResponse);
    };
    var client = { get: send, post: send } as SPHttpClient;
    rest = new SharePointRestClient('https://example.test', client, false);
  });

  it('uses OData 3.0 for list discovery instead of the SDK v1 default of 4.0', function(): Promise<void> {
    return rest.get('web/lists?$select=Title').then(function(result): void {
      assert.equal(requestedUrl, 'https://example.test/_api/web/lists?$select=Title');
      assert.equal(requestOptions.headers['OData-Version'], '3.0');
      assert.equal(requestOptions.headers['Accept'], 'application/json;odata=nometadata');
      assert.deepEqual(rest.rows(result.data), []);
    });
  });

  it('sends list creation metadata using the matching OData 3.0 protocol', function(): Promise<void> {
    var body = { __metadata: { type: 'SP.List' }, BaseTemplate: 100, Title: 'Events' };
    responseBody = '{"d":{"Id":"list-id"}}';
    return rest.post('web/lists', body).then(function(result): void {
      assert.equal(requestedUrl, 'https://example.test/_api/web/lists');
      assert.equal(requestOptions.headers['OData-Version'], '3.0');
      assert.equal(requestOptions.headers['Content-Type'], 'application/json;odata=verbose');
      assert.equal(requestOptions.headers['Accept'], 'application/json;odata=nometadata');
      assert.deepEqual(JSON.parse(requestOptions.body), body);
      assert.equal(rest.item(result.data).Id, 'list-id');
    });
  });

  it('preserves field creation parameters and exact ETags for item updates', function(): Promise<void> {
    var body = { parameters: {
      __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
      SchemaXml: '<Field Type="Text" Name="ERLocation" />',
      Options: 0
    } };
    return rest.post("web/lists/getByTitle('Events')/fields/CreateFieldAsXml", body).then(function(): Promise<any> {
      assert.equal(requestOptions.headers['OData-Version'], '3.0');
      assert.deepEqual(JSON.parse(requestOptions.body), body);
      return rest.post("web/lists/getByTitle('Slots')/items(1)", { ERStatus: 'claimed' }, {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '"7"'
      });
    }).then(function(): void {
      assert.equal(requestOptions.headers['OData-Version'], '3.0');
      assert.equal(requestOptions.headers['X-HTTP-Method'], 'MERGE');
      assert.equal(requestOptions.headers['IF-MATCH'], '"7"');
      assert.deepEqual(JSON.parse(requestOptions.body), { ERStatus: 'claimed' });
    });
  });

  it('keeps SharePoint setup errors visible', function(): Promise<void> {
    status = 400;
    responseBody = '{"error":{"message":{"value":"Invalid list schema"}}}';
    return rest.post('web/lists', {}).then(function(): void {
      assert.fail('Expected list creation to fail');
    }, function(error): void {
      assert.equal(error.status, 400);
      assert.include(error.message, 'POST https://example.test/_api/web/lists failed with HTTP 400');
      assert.include(error.message, 'Invalid list schema');
    });
  });
});
