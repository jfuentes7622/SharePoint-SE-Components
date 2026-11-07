import { DisplayMode } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IInstanceNameProperties {
  instanceName?: string;
  instanceNameGeneratedFor?: string;
}

interface IPageName {
  instanceId: string;
  name: string;
}

interface INameRegistry {
  pageKey: string;
  liveNames: { [instanceId: string]: string };
  pageNames?: Promise<IPageName[]>;
}

interface INameDocument extends Document {
  __spseInstanceNames?: INameRegistry;
}

function normalizeId(value: string): string {
  return String(value || '').replace(/[{}]/g, '').toLowerCase();
}

function readControl(data: any): IPageName {
  var webPart = data && (data.webPartData || data);
  return {
    instanceId: normalizeId(webPart && (webPart.instanceId || data.id)),
    name: String(webPart && webPart.properties && webPart.properties.instanceName || '').trim()
  };
}

function readDocumentNames(ownerDocument: Document): IPageName[] {
  var names: IPageName[] = [];
  var controls = ownerDocument.querySelectorAll('[data-sp-controldata]');
  for (var index = 0; index < controls.length; index += 1) {
    names.push(readControl(JSON.parse(controls[index].getAttribute('data-sp-controldata'))));
  }
  return names;
}

function readCanvasNames(canvas: string, ownerDocument: Document): IPageName[] {
  var source = String(canvas || '').trim();
  if (!source) { return []; }
  if (source.charAt(0) === '[') {
    var controls = JSON.parse(source);
    if (!Array.isArray(controls)) { throw new Error('Invalid page canvas metadata.'); }
    return controls.map(readControl);
  }
  var snapshot = ownerDocument.implementation.createHTMLDocument('');
  snapshot.body.innerHTML = source;
  return readDocumentNames(snapshot);
}

function needsName(properties: IInstanceNameProperties, prefix: string, instanceId: string): boolean {
  var name = String(properties.instanceName || '').trim();
  return !name || (!!properties.instanceNameGeneratedFor
    && normalizeId(properties.instanceNameGeneratedFor) !== instanceId
    && new RegExp('^' + prefix + '[1-9][0-9]*$', 'i').test(name));
}

export function ensureInstanceName(
  properties: IInstanceNameProperties, prefix: string, context: WebPartContext,
  element: HTMLElement, getDisplayMode: () => DisplayMode
): Promise<boolean> {
  if (getDisplayMode() !== DisplayMode.Edit) { return Promise.resolve(false); }
  var ownerDocument = element.ownerDocument as INameDocument;
  var pageContext = context.pageContext;
  var pageKey = pageContext.web.absoluteUrl + '|'
    + (pageContext.list && pageContext.list.id.toString() || '') + '|'
    + (pageContext.listItem && pageContext.listItem.id || ownerDocument.location.pathname);
  var registry = ownerDocument.__spseInstanceNames;
  if (!registry || registry.pageKey !== pageKey) {
    registry = { pageKey: pageKey, liveNames: {} };
    ownerDocument.__spseInstanceNames = registry;
  }
  var instanceId = normalizeId(context.instanceId);
  var registerName = function(): void {
    registry.liveNames[instanceId] = String(properties.instanceName || '').trim();
    element.setAttribute('data-spse-instance-id', instanceId);
    element.setAttribute('data-spse-instance-name', registry.liveNames[instanceId]);
  };
  if (!needsName(properties, prefix, instanceId)) {
    registerName();
    return Promise.resolve(false);
  }
  if (!registry.pageNames) {
    if (pageContext.list && pageContext.listItem) {
      var listId = normalizeId(pageContext.list.id.toString());
      if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(listId) || !(pageContext.listItem.id > 0)) {
        return Promise.reject(new Error('Cannot allocate a web part name: invalid page identity.'));
      }
      var url = pageContext.web.absoluteUrl + "/_api/web/lists(guid'" + listId + "')/items(" + pageContext.listItem.id + ')';
      registry.pageNames = context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=verbose', 'OData-Version': '3.0' }
      }).then(function(response): Promise<IPageName[]> {
        if (!response.ok) { throw new Error('Cannot read page names: HTTP ' + response.status); }
        return response.json().then(function(data): IPageName[] {
          var page = data && (data.d || data);
          if (!page || typeof page !== 'object' || page.error) { throw new Error('Invalid page metadata response.'); }
          return readCanvasNames(page.CanvasContent1, ownerDocument);
        });
      }).catch(function(error): never {
        registry.pageNames = undefined;
        throw error;
      });
    } else {
      registry.pageNames = Promise.resolve([]);
    }
  }
  return registry.pageNames.then(function(savedNames): boolean {
    if (ownerDocument.__spseInstanceNames !== registry || getDisplayMode() !== DisplayMode.Edit) { return false; }
    if (!needsName(properties, prefix, instanceId)) { registerName(); return false; }
    var namesById: { [instanceId: string]: string } = {};
    savedNames.concat(readDocumentNames(ownerDocument)).forEach(function(item: IPageName, index: number): void {
      if (item.name) { namesById[item.instanceId || 'unidentified-' + index] = item.name; }
    });
    Object.keys(registry.liveNames).forEach(function(id: string): void { namesById[id] = registry.liveNames[id]; });
    var used: { [name: string]: boolean } = {};
    Object.keys(namesById).forEach(function(id: string): void {
      if (id !== instanceId) { used[namesById[id].toLowerCase()] = true; }
    });
    var number = 1;
    while (used[(prefix + number).toLowerCase()]) { number += 1; }
    properties.instanceName = prefix + number;
    properties.instanceNameGeneratedFor = instanceId;
    registerName();
    return true;
  });
}
