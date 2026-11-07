export interface ILibraryFolder {
  path: string;
  name: string;
}

export interface ILibraryCondition {
  field: string;
  operator?: string;
  logical?: string;
  valueType?: string;
  value: any;
}

export interface ILibraryLoadOptions {
  webUrl: string;
  listName: string;
  viewXml: string;
  fields: string[];
  conditions: ILibraryCondition[];
  time: number;
  batchSize: number;
  folder?: string;
  get: (url: string) => Promise<any>;
  post: (url: string, body: any) => Promise<any>;
  extract: (data: any) => { rows: any[]; fields: any[]; nextHref: string };
  current: () => boolean;
}

export function validateLibraryFilterJson(sources: string[]): void {
  sources.forEach(function(source: string): void {
    if (!String(source || '').trim()) { return; }
    var parsed = JSON.parse(source);
    if (!Array.isArray(parsed) || parsed.some(function(condition: any): boolean {
      return !condition || typeof condition.field !== 'string' || !condition.field.trim();
    })) { throw new Error('Library filters must be an array of conditions with field names.'); }
  });
}

export function resolveNowExpression(expression: string, time: number): string {
  var normalized = String(expression || '').trim().toLowerCase();
  if (normalized === 'now') { return new Date(time).toISOString(); }
  var match = /^now\(\s*([+-]?\d+)h\s*\)$/.exec(normalized);
  var value = match ? time + Number(match[1]) * 3600000 : NaN;
  if (!isFinite(value) || Math.abs(value) > 8640000000000000) {
    throw new Error('Invalid elapsed-hour expression: ' + expression);
  }
  return new Date(value).toISOString();
}

export function documentUrl(webUrl: string, path: string): string {
  var value = String(path || '').trim();
  if (!value || /^\/\//.test(value) || /^(?!https?:)[a-z][a-z0-9+.-]*:/i.test(value) || /[\\\r\n]/.test(value)) { return ''; }
  var origin = /^https?:\/\/[^/]+/i.exec(webUrl);
  if (!origin) { return ''; }
  var absolute = /^https?:\/\//i.test(value) ? value
    : value.charAt(0) === '/' ? origin[0] + value : webUrl.replace(/\/$/, '') + '/' + value;
  // FileRef can already contain escapes. Preserve those, but escape literal %, # and spaces.
  return encodeURI(absolute.replace(/%(?![0-9a-f]{2})/ig, '%25'))
    .replace(/%25([0-9a-f]{2})/ig, '%$1').replace(/#/g, '%23');
}

function cleanPath(path: string): string {
  var value = String(path || '').replace(/\/+$/, '');
  if (!value || value.charAt(0) !== '/' || /[\\\r\n]/.test(value)
    || value.split('/').some(function(segment: string): boolean {
      return segment === '.' || segment === '..' || /^(%2e){1,2}$/i.test(segment);
    })) {
    throw new Error('Invalid library path: ' + path);
  }
  return value;
}

function withinRoot(root: string, path: string): boolean {
  return path.toLowerCase() === root.toLowerCase()
    || path.toLowerCase().indexOf(root.toLowerCase() + '/') === 0;
}

export function folderDocuments(current: string, documents: any[], recursive: boolean = false): any[] {
  var path = cleanPath(current);
  return documents.filter(function(row: any): boolean {
    var parent = cleanPath(row.FileDirRef);
    return recursive ? withinRoot(path, parent) : parent.toLowerCase() === path.toLowerCase();
  });
}

export function folderDocumentCount(current: string, documents: any[]): number {
  return folderDocuments(current, documents, true).length;
}

export function ancestorPaths(rootPath: string, rows: any[]): string[] {
  var root = cleanPath(rootPath);
  var paths: { [key: string]: string } = {};
  rows.forEach(function(row: any): void {
    var parent = cleanPath(row.FileDirRef);
    if (!withinRoot(root, parent)) { throw new Error('Document is outside the library root: ' + parent); }
    while (parent.toLowerCase() !== root.toLowerCase()) {
      paths[parent.toLowerCase()] = parent;
      parent = parent.substring(0, parent.lastIndexOf('/'));
    }
  });
  return Object.keys(paths).map(function(key: string): string { return paths[key]; });
}

export function folderLevel(root: string, current: string, documents: any[], folders: ILibraryFolder[], prune: boolean = true): any[] {
  var required = ancestorPaths(root, documents).map(function(path: string): string { return path.toLowerCase(); });
  var children = folders.filter(function(folder: ILibraryFolder): boolean {
    return (!prune || required.indexOf(folder.path.toLowerCase()) >= 0)
      && folder.path.substring(0, folder.path.lastIndexOf('/')).toLowerCase() === current.toLowerCase();
  }).sort(function(left: ILibraryFolder, right: ILibraryFolder): number { return left.name.localeCompare(right.name); });
  return children.map(function(folder: ILibraryFolder): any { return { libraryFolder: folder }; })
    .concat(folderDocuments(current, documents));
}

export function nearestFolder(root: string, current: string, paths: string[]): string {
  var keys = paths.map(function(path: string): string { return path.toLowerCase(); });
  var candidate = current || root;
  while (withinRoot(root, candidate) && candidate.toLowerCase() !== root.toLowerCase()) {
    if (keys.indexOf(candidate.toLowerCase()) >= 0) { return candidate; }
    candidate = candidate.substring(0, candidate.lastIndexOf('/'));
  }
  return root;
}

export function recursiveView(viewXml: string, fields: string[], conditions: ILibraryCondition[], time: number, batchSize: number, dateFields?: string[], recursive: boolean = true): string {
  var doc = new DOMParser().parseFromString(viewXml || '<View/>', 'text/xml');
  if (doc.getElementsByTagName('parsererror').length) { throw new Error('Invalid source view XML.'); }
  var view = doc.documentElement;
  if (recursive) { view.setAttribute('Scope', 'RecursiveAll'); }
  ['GroupBy', 'Aggregations', 'Folder'].forEach(function(tag: string): void {
    var nodes = view.getElementsByTagName(tag);
    while (nodes.length) { nodes[0].parentNode.removeChild(nodes[0]); }
  });
  var query = view.getElementsByTagName('Query')[0];
  if (!query) { query = doc.createElement('Query'); view.appendChild(query); }
  var where = query.getElementsByTagName('Where')[0];
  if (!where) { where = doc.createElement('Where'); query.appendChild(where); }
  function predicate(tag: string, field: string, type: string, value: string): Element {
    var node = doc.createElement(tag);
    var ref = doc.createElement('FieldRef'); ref.setAttribute('Name', field);
    var val = doc.createElement('Value'); val.setAttribute('Type', type);
    if (type === 'DateTime') {
      val.setAttribute('IncludeTimeValue', value.indexOf('T') >= 0 ? 'TRUE' : 'FALSE');
      if (value.indexOf('T') >= 0) { val.setAttribute('StorageTZ', 'TRUE'); }
    }
    val.appendChild(doc.createTextNode(value)); node.appendChild(ref); node.appendChild(val);
    return node;
  }
  function appendPredicate(node: Element): void {
    var existing = where.firstElementChild;
    if (existing) {
      var and = doc.createElement('And'); where.removeChild(existing);
      and.appendChild(existing); and.appendChild(node); where.appendChild(and);
    } else { where.appendChild(node); }
  }
  appendPredicate(predicate('Eq', 'FSObjType', 'Integer', '0'));
  // Push down only safe conjuncts. Mixed OR expressions retain their complete client evaluation.
  if (!conditions.some(function(condition: ILibraryCondition, index: number): boolean {
    return index > 0 && String(condition.logical || '').toLowerCase() === 'or';
  })) {
    var operators: { [key: string]: string } = { eq: 'Eq', ge: 'Geq', le: 'Leq', gt: 'Gt', lt: 'Lt', ne: 'Neq' };
    conditions.forEach(function(condition: ILibraryCondition): void {
      if (dateFields && dateFields.indexOf(condition.field) < 0) { return; }
      var expression = String(condition.value || '').trim().toLowerCase();
      var value = '';
      if (String(condition.valueType).toLowerCase() === 'expression') {
        if (/^now(?:\(|$)/.test(expression)) { value = resolveNowExpression(expression, time); }
        var days = /^date\(\s*([+-]?\d+)\s*\)$/.exec(expression);
        if (expression === 'today' || days) {
          var local = new Date(time);
          if (days) { local.setDate(local.getDate() + Number(days[1])); }
          value = String(local.getFullYear()) + '-' + ('0' + String(local.getMonth() + 1)).slice(-2) + '-' + ('0' + String(local.getDate())).slice(-2);
        }
      } else if (typeof condition.value === 'string' && /^\d{4}-\d\d-\d\d(?:T.*)?$/.test(condition.value)) {
        value = condition.value;
      }
      var operator = operators[String(condition.operator || 'eq').toLowerCase()];
      if (value && operator) { appendPredicate(predicate(operator, condition.field, 'DateTime', value)); }
    });
  }
  var viewFields = view.getElementsByTagName('ViewFields')[0];
  if (!viewFields) { viewFields = doc.createElement('ViewFields'); view.appendChild(viewFields); }
  var existingNames: string[] = [];
  var refs = viewFields.getElementsByTagName('FieldRef');
  for (var index = 0; index < refs.length; index += 1) { existingNames.push(String(refs[index].getAttribute('Name') || '').toLowerCase()); }
  fields.forEach(function(field: string): void {
    if (existingNames.indexOf(field.toLowerCase()) >= 0) { return; }
    existingNames.push(field.toLowerCase());
    var ref = doc.createElement('FieldRef'); ref.setAttribute('Name', field); viewFields.appendChild(ref);
  });
  var limit = view.getElementsByTagName('RowLimit')[0];
  if (!limit) { limit = doc.createElement('RowLimit'); view.appendChild(limit); }
  limit.setAttribute('Paged', 'TRUE'); limit.textContent = String(Math.max(1, batchSize || 200));
  return new XMLSerializer().serializeToString(doc);
}

async function json(response: any, operation: string): Promise<any> {
  if (!response.ok) {
    var body = typeof response.text === 'function' ? await response.text() : '';
    var requestId = response.headers && response.headers.get
      ? response.headers.get('SPRequestGuid') || response.headers.get('request-id') : '';
    var detail = String(body || '').substring(0, 2000);
    var guidance = /SPQueryThrottledException|list view threshold|exceeds.*threshold/i.test(detail)
      ? ' Use folder-scoped loading and a selective indexed view/filter. A folder can also exceed the list view threshold.' : '';
    throw new Error(operation + ': HTTP ' + String(response.status) + ' ' + String(response.statusText || '')
      + (detail ? ' - ' + detail : '') + (requestId ? ' [SPRequestGuid: ' + requestId + ']' : '') + guidance);
  }
  return response.json();
}

function values(data: any): any[] {
  var rows = data && (data.value || (data.d && data.d.results));
  if (!Array.isArray(rows)) { throw new Error('Invalid SharePoint collection response.'); }
  return rows;
}

export async function detectDocumentLibrary(webUrl: string, listName: string, get: (url: string) => Promise<any>): Promise<boolean> {
  var data = await json(await get(webUrl + "/_api/web/lists/getByTitle('"
    + encodeURIComponent(listName.replace(/'/g, "''")) + "')?$select=BaseType"), 'List type discovery');
  var source = data && (data.d || data);
  if (!source || !/^[0-5]$/.test(String(source.BaseType))) {
    throw new Error('SharePoint did not return a valid list BaseType.');
  }
  return Number(source.BaseType) === 1;
}

export async function loadLibraryFolders(
  webUrl: string, listName: string, root: string, documents: any[],
  get: (url: string) => Promise<any>, current: () => boolean, batchSize: number
): Promise<ILibraryFolder[]> {
  var next = webUrl + "/_api/web/lists/getByTitle('" + encodeURIComponent(listName.replace(/'/g, "''"))
    + "')/items?$select=FileRef,FileDirRef,FileLeafRef&$filter=FSObjType%20eq%201&$top=" + String(Math.max(1, batchSize || 200));
  var seenPages: { [url: string]: boolean } = {};
  var byPath: { [path: string]: ILibraryFolder } = {};
  while (next) {
    if (!current()) { throw new Error('Superseded library folder request.'); }
    if (seenPages[next]) { throw new Error('Library folder paging did not advance.'); }
    seenPages[next] = true;
    var data = await json(await get(next), 'Library folders');
    values(data).forEach(function(item: any): void {
      var path = cleanPath(item.FileRef);
      if (!withinRoot(root, path) || path.toLowerCase() === root.toLowerCase() || !item.FileLeafRef
        || path.substring(0, path.lastIndexOf('/')).toLowerCase() !== cleanPath(item.FileDirRef).toLowerCase()) {
        throw new Error('Invalid library folder metadata: ' + path);
      }
      byPath[path.toLowerCase()] = { path: path, name: String(item.FileLeafRef) };
    });
    next = data['@odata.nextLink'] || data['odata.nextLink'] || (data.d && data.d.__next) || '';
  }
  var folders = Object.keys(byPath).map(function(key: string): ILibraryFolder { return byPath[key]; });
  var required = ancestorPaths(root, documents.concat(folders.map(function(folder: ILibraryFolder): any {
    return { FileDirRef: folder.path.substring(0, folder.path.lastIndexOf('/')) };
  })));
  required.forEach(function(path: string): void {
    if (!byPath[path.toLowerCase()]) { throw new Error('Library ancestor folder is unavailable: ' + path); }
  });
  return folders;
}

export async function defaultLibraryViewId(webUrl: string, listName: string, get: (url: string) => Promise<any>): Promise<string> {
  var data = await json(await get(webUrl + "/_api/web/lists/getByTitle('"
    + encodeURIComponent(listName.replace(/'/g, "''")) + "')/DefaultView?$select=Id"), 'Library default view');
  var source = data.d || data;
  if (!source.Id) { throw new Error('SharePoint did not return the default library view.'); }
  return String(source.Id);
}

export async function loadLibraryDocuments(options: ILibraryLoadOptions): Promise<{ root: string; folder: string; flat: boolean; rows: any[]; fields: any[]; metadata: any[] }> {
  var listUrl = options.webUrl + "/_api/web/lists/getByTitle('" + encodeURIComponent(options.listName.replace(/'/g, "''")) + "')";
  var sourceData = await json(await options.get(listUrl + '?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder'), 'Library discovery');
  var source = sourceData.d || sourceData;
  if (Number(source.BaseType) !== 1) { throw new Error('Matching document folders requires a document library.'); }
  var root = cleanPath(source.RootFolder && source.RootFolder.ServerRelativeUrl);
  var folder = options.folder === undefined ? '' : cleanPath(options.folder || root);
  if (folder && !withinRoot(root, folder)) { throw new Error('Folder is outside the library root: ' + folder); }
  var metadata = values(await json(await options.get(listUrl
    + '/fields?$select=InternalName,Title,TypeAsString,DisplayFormat,Choices,LookupList,LookupField,AllowMultipleValues'), 'Library field metadata'));
  var names = options.fields.concat(['ID', 'FSObjType', 'FileRef', 'FileDirRef', 'FileLeafRef', 'Modified', 'Created']);
  options.conditions.forEach(function(condition: ILibraryCondition): void {
    if (String(condition.valueType).toLowerCase() === 'expression') {
      var expression = String(condition.value || '').trim().toLowerCase();
      if (/^now(?:\(|$)/.test(expression)) { resolveNowExpression(expression, options.time); }
      else if (['today', 'me', 'me.email', 'me.login', 'me.id'].indexOf(expression) < 0 && !/^date\(\s*[+-]?\d+\s*\)$/.test(expression)) {
        throw new Error('Unsupported library filter expression: ' + expression);
      }
    }
    if (names.indexOf(condition.field) < 0) { names.push(condition.field); }
  });
  function internalName(name: string): string {
    var match = metadata.filter(function(field: any): boolean {
      return String(field.InternalName).toLowerCase() === name.toLowerCase() || String(field.Title).toLowerCase() === name.toLowerCase();
    })[0];
    return match ? String(match.InternalName) : name;
  }
  names = names.map(internalName);
  names = names.filter(function(name: string, index: number): boolean { return !!name && names.indexOf(name) === index; });
  var dateFields = metadata.filter(function(field: any): boolean { return field.TypeAsString === 'DateTime'; })
    .map(function(field: any): string { return field.InternalName; });
  var queryConditions = options.conditions.map(function(condition: ILibraryCondition): ILibraryCondition {
    return Object.assign({}, condition, { field: internalName(condition.field) });
  });
  var viewXml = recursiveView(options.viewXml, names, queryConditions, options.time, options.batchSize, dateFields, !folder);
  var sourceScope = new DOMParser().parseFromString(options.viewXml || '<View/>', 'text/xml').documentElement.getAttribute('Scope') || '';
  var flat = /^(Recursive|RecursiveAll)$/i.test(sourceScope);
  var rows: any[] = [];
  var fields: any[] = [];
  var paging = '';
  var seen: { [key: string]: boolean } = {};
  do {
    if (!options.current()) { throw new Error('Superseded library request.'); }
    var parameters: any = { RenderOptions: 7, ViewXml: viewXml, DatesInUtc: true };
    if (folder) { parameters.FolderServerRelativeUrl = folder; }
    if (paging) { parameters.Paging = paging.replace(/^\?/, ''); }
    var data = await json(await options.post(listUrl + '/RenderListDataAsStream', { parameters: parameters }), 'Library documents');
    var renderData = data && data.d && data.d.RenderListDataAsStream ? data.d.RenderListDataAsStream : data;
    if (typeof renderData === 'string') { renderData = JSON.parse(renderData); }
    var listData = renderData.ListData || renderData;
    if (typeof listData === 'string') { listData = JSON.parse(listData); }
    var renderError = renderData.ErrorInfo || renderData.ErrorMessage || listData.ErrorInfo || listData.ErrorMessage;
    if (renderError) {
      throw new Error('SharePoint could not render the library view: ' + JSON.stringify(renderError));
    }
    var batch = options.extract(data);
    rows = rows.concat(batch.rows);
    if (!fields.length) { fields = batch.fields; }
    paging = batch.nextHref || '';
    if (paging && seen[paging]) { throw new Error('Library paging did not advance.'); }
    seen[paging] = true;
  } while (paging);
  // REST supplies authoritative UTC dates and decoded paths, independently of display formatting.
  var dateNames = metadata.filter(function(field: any): boolean {
    return field.TypeAsString === 'DateTime' && names.indexOf(field.InternalName) >= 0;
  }).map(function(field: any): string { return field.InternalName; });
  var select = ['ID', 'FileRef', 'FileDirRef', 'FileLeafRef', 'FSObjType'].concat(dateNames);
  var hydrated: { [id: string]: any } = {};
  if (folder && !flat && rows.length) {
    // Expand only this folder's files; a library-wide OR of item IDs can still be throttled.
    var filesUrl = options.webUrl + "/_api/web/GetFolderByServerRelativeUrl('"
      + encodeURIComponent(folder.replace(/'/g, "''")) + "')/Files?$expand=ListItemAllFields&$select="
      + select.map(function(name: string): string { return 'ListItemAllFields/' + name; }).join(',')
      + '&$top=' + String(Math.max(1, options.batchSize || 200));
    var filePages: { [url: string]: boolean } = {};
    while (filesUrl) {
      if (!options.current()) { throw new Error('Superseded library request.'); }
      if (filePages[filesUrl]) { throw new Error('Folder document metadata paging did not advance.'); }
      filePages[filesUrl] = true;
      var fileData = await json(await options.get(filesUrl), 'Folder document metadata');
      values(fileData).forEach(function(file: any): void {
        var item = file.ListItemAllFields;
        var id = Number(item && item.ID);
        if (!item || !isFinite(id) || id <= 0 || Math.floor(id) !== id || cleanPath(item.FileDirRef).toLowerCase() !== folder.toLowerCase()) {
          throw new Error('Invalid folder document metadata.');
        }
        hydrated[String(item.ID)] = item;
      });
      filesUrl = fileData['@odata.nextLink'] || fileData['odata.nextLink'] || (fileData.d && fileData.d.__next) || '';
    }
    rows.forEach(function(row: any): void {
      if (!hydrated[String(row.ID || row.Id)]) { throw new Error('Document metadata is no longer accessible: ' + String(row.ID || row.Id)); }
    });
  }
  var cursor = 0;
  var hydrationError: any;
  async function hydrateFlatRows(): Promise<void> {
    while (cursor < rows.length && !hydrationError) {
      var row = rows[cursor++];
      var id = Number(row.ID || row.Id);
      try {
        if (!options.current()) { throw new Error('Superseded library request.'); }
        if (!isFinite(id) || id <= 0 || Math.floor(id) !== id) { throw new Error('Library returned an invalid document ID.'); }
        var itemData = await json(await options.get(listUrl + '/items(' + String(id) + ')?$select=' + select.join(',')), 'Library document metadata');
        var item = itemData.d || itemData;
        if (Number(item.ID || item.Id) !== id) { throw new Error('Document metadata is no longer accessible: ' + String(id)); }
        hydrated[String(id)] = item;
      } catch (error) { hydrationError = error; }
    }
  }
  if (flat) {
    // Indexed item reads avoid both library-wide OR filters and scanning unrelated folder files.
    await Promise.all([hydrateFlatRows(), hydrateFlatRows(), hydrateFlatRows(), hydrateFlatRows()]);
    if (hydrationError) { throw hydrationError; }
  }
  for (var offset = 0; !folder && !flat && offset < rows.length; offset += 50) {
    if (!options.current()) { throw new Error('Superseded library request.'); }
    var ids = rows.slice(offset, offset + 50).map(function(row: any): number { return Number(row.ID || row.Id); });
    if (ids.some(function(id: number): boolean { return !isFinite(id) || id <= 0 || Math.floor(id) !== id; })) {
      throw new Error('Library returned an invalid document ID.');
    }
    var next = listUrl + '/items?$top=50&$select=' + select.join(',') + '&$filter='
      + encodeURIComponent(ids.map(function(id: number): string { return 'ID eq ' + String(id); }).join(' or '));
    var itemPages: { [url: string]: boolean } = {};
    while (next) {
      if (!options.current()) { throw new Error('Superseded library request.'); }
      if (itemPages[next]) { throw new Error('Library metadata paging did not advance.'); }
      itemPages[next] = true;
      var itemData = await json(await options.get(next), 'Library document metadata');
      values(itemData).forEach(function(item: any): void { hydrated[String(item.ID || item.Id)] = item; });
      next = itemData['@odata.nextLink'] || itemData['odata.nextLink'] || (itemData.d && itemData.d.__next) || '';
    }
    ids.forEach(function(id: number): void {
      if (!hydrated[String(id)]) { throw new Error('Document metadata is no longer accessible: ' + String(id)); }
    });
  }
  rows = rows.map(function(row: any): any { return Object.assign({}, row, hydrated[String(row.ID || row.Id)]); })
    .filter(function(row: any): boolean { return Number(row.FSObjType || 0) === 0; });
  return { root: root, folder: folder || root, flat: flat, rows: rows, fields: fields, metadata: metadata };
}

export async function loadChildFolders(webUrl: string, root: string, folder: string, get: (url: string) => Promise<any>,
  current: () => boolean, batchSize: number): Promise<ILibraryFolder[]> {
  var path = cleanPath(folder);
  if (!withinRoot(cleanPath(root), path)) { throw new Error('Folder is outside the library root: ' + path); }
  var next = webUrl + "/_api/web/GetFolderByServerRelativeUrl('" + encodeURIComponent(path.replace(/'/g, "''"))
    + "')/Folders?$select=Name,ServerRelativeUrl&$top=" + String(Math.max(1, batchSize || 200));
  var pages: { [url: string]: boolean } = {};
  var folders: ILibraryFolder[] = [];
  while (next) {
    if (!current()) { throw new Error('Superseded library request.'); }
    if (pages[next]) { throw new Error('Child folder paging did not advance.'); }
    pages[next] = true;
    var data = await json(await get(next), 'Child folders');
    values(data).forEach(function(item: any): void {
      var child = cleanPath(item.ServerRelativeUrl);
      if (child.substring(0, child.lastIndexOf('/')).toLowerCase() !== path.toLowerCase() || !item.Name) {
        throw new Error('Invalid child folder metadata: ' + child);
      }
      if (path.toLowerCase() !== root.toLowerCase() || String(item.Name).toLowerCase() !== 'forms') {
        folders.push({ path: child, name: String(item.Name) });
      }
    });
    next = data['@odata.nextLink'] || data['odata.nextLink'] || (data.d && data.d.__next) || '';
  }
  return folders;
}

export async function loadAncestorFolders(webUrl: string, root: string, rows: any[], get: (url: string) => Promise<any>): Promise<ILibraryFolder[]> {
  var paths = ancestorPaths(root, rows);
  var folders: ILibraryFolder[] = [];
  var cursor = 0;
  var failure: any;
  async function worker(): Promise<void> {
    while (cursor < paths.length) {
      var path = paths[cursor++];
      var endpoint = webUrl + "/_api/web/GetFolderByServerRelativeUrl('"
        + encodeURIComponent(path.replace(/'/g, "''")) + "')?$select=Name,ServerRelativeUrl";
      var data = await json(await get(endpoint), 'Folder ' + path);
      var source = data.d || data;
      if (cleanPath(source.ServerRelativeUrl).toLowerCase() !== path.toLowerCase() || !source.Name) {
        throw new Error('Invalid folder metadata: ' + path);
      }
      folders.push({ path: path, name: String(source.Name) });
    }
  }
  function boundedWorker(): Promise<void> {
    return worker().catch(function(error: any): void {
      failure = error;
      cursor = paths.length;
    });
  }
  await Promise.all([boundedWorker(), boundedWorker(), boundedWorker(), boundedWorker()]);
  if (failure) { throw failure; }
  return folders;
}
