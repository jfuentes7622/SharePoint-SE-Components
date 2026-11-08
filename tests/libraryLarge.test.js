'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

async function run(project) {
  const root = path.join(__dirname, '..', project);
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const React = require(path.join(root, 'node_modules', 'react'));
  const ReactDomServer = require(path.join(root, 'node_modules', 'react-dom', 'server'));
  const document = require(path.join(root, 'node_modules', 'jsdom')).jsdom('<html></html>');
  const window = document.defaultView;
  window.setInterval = () => 42;
  window.clearInterval = () => {};
  const directory = path.join(root, 'src', 'webparts', project === 'ListControl' ? 'listControl' : 'gridControl');
  const library = require('./loadLibraryBrowser')(root, window);
  let strings;
  vm.runInNewContext(fs.readFileSync(path.join(directory, 'loc', 'en-us.js'), 'utf8'), {
    define: (dependencies, factory) => { strings = factory(); }
  });
  const exports = {};
  const errors = [];
  vm.runInNewContext(ts.transpileModule(fs.readFileSync(path.join(directory, 'components', project + '.tsx'), 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
  }).outputText, {
    exports, window, document, Date,
    console: { log() {}, error(...args) { errors.push(args); } },
    require(name) {
      if (name === 'react') { return React; }
      if (name === './LibraryBrowser') { return library; }
      if (name === project + 'WebPartStrings') { return strings; }
      if (name === '@microsoft/sp-http') { return { SPHttpClient: { configurations: { v1: {} } } }; }
      if (['./DateTime24HourInput', './GridValidation', './GridRichTextEditor'].includes(name) || /\.css$/.test(name)) { return {}; }
      throw new Error('Unexpected import: ' + name);
    }
  });
  const web = 'https://example.test/sites/team/sub';
  const libraryRoot = '/sites/team/sub/Documents';
  const folderNames = Array.from({ length: 21 }, (_, index) => index === 0 ? "Team #1's files" : 'Team ' + index);
  const documents = Array.from({ length: 21000 }, (_, index) => ({
    ID: index + 1, FSObjType: 0, Title: 'Document ' + (index + 1),
    FileDirRef: libraryRoot + '/' + folderNames[Math.floor(index / 1000)],
    FileLeafRef: 'Document ' + (index + 1) + '.docx',
    Modified: '2026-11-05T12:00:00Z', Created: '2026-01-01T00:00:00Z',
    Status: index % 3 ? 'Active' : 'Inactive', Category: index % 2 ? 'Red' : 'Blue'
  }));
  documents.forEach(item => { item.FileRef = item.FileDirRef + '/' + item.FileLeafRef; });
  assert.strictEqual(documents.length, 21000);
  const metadata = ['Title', 'Status', 'Category', 'Modified', 'Created'].map(name => ({
    InternalName: name, Title: name, TypeAsString: /Modified|Created/.test(name) ? 'DateTime' : 'Text'
  }));
  const fieldMap = {};
  metadata.forEach(item => { fieldMap[item.InternalName.toLowerCase()] = {
    internalName: item.InternalName, title: item.Title, typeAsString: item.TypeAsString, choices: []
  }; });
  const requests = [];
  const selections = [];
  const counts = [];
  let failure = '';
  let hold;
  let release;
  const response = (data, status = 200) => ({
    ok: status < 400, status, statusText: status === 500 ? 'Internal Server Error' : status === 403 ? 'Forbidden' : 'OK',
    json: async () => data, text: async () => JSON.stringify(data),
    headers: { get: name => name === 'SPRequestGuid' ? 'large-library-correlation-id' : undefined },
    clone() { return response(data, status); }
  });
  const threshold = () => response({
    error: { code: 'Microsoft.SharePoint.SPQueryThrottledException', message: {
      value: 'The attempted operation is prohibited because it exceeds the list view threshold.'
    } }
  }, 500);
  const folderFromUrl = url => decodeURIComponent(url.match(/RelativeUrl\('(.+?)'\)/)[1]).replace(/''/g, "'");
  const client = {
    async get(url, configuration, options) {
      requests.push({ method: 'GET', url, options });
      if (url.includes('/fields?')) { return response({ value: metadata }); }
      if (url.includes('/items?')) { throw new Error('Folder mode must not issue library-wide item-ID OR queries.'); }
      if (url.includes('GetFolderByServerRelativeUrl')) {
        const folder = folderFromUrl(url);
        assert(folder === libraryRoot || folder.startsWith(libraryRoot + '/'));
        if (url.includes('/Folders?')) {
          if (failure === 'folders') { return response({ error: 'Folder access denied' }, 403); }
          const names = folder === libraryRoot ? folderNames.concat(['Empty', 'Forms']) : [];
          return response({ value: names.map(name => ({ Name: name, ServerRelativeUrl: folder + '/' + name })) });
        }
        assert(url.includes('/Files?'), 'Only immediate child folder/file collections are requested.');
        if (failure === 'files') { return response({ error: 'File metadata access denied' }, 403); }
        assert(url.includes('$expand=ListItemAllFields') && url.includes('ListItemAllFields/Modified'));
        const offset = Number((url.match(/filePage=(\d+)/) || [])[1] || 0);
        const files = documents.filter(item => item.FileDirRef === folder);
        return response({ value: files.slice(offset, offset + 200).map(item => ({ ListItemAllFields: item })),
          '@odata.nextLink': offset + 200 < files.length ? url.replace(/&filePage=\d+/, '') + '&filePage=' + (offset + 200) : '' });
      }
      return response({ BaseType: 1, RootFolder: { ServerRelativeUrl: libraryRoot } });
    },
    async post(url, configuration, options) {
      requests.push({ method: 'POST', url, options });
      const parameters = JSON.parse(options.body).parameters;
      if (hold) { await hold; }
      if (!parameters.FolderServerRelativeUrl) { return threshold(); }
      assert(!parameters.ViewXml.includes('RecursiveAll'), 'A folder-based source view must not become recursive.');
      assert(!parameters.ViewXml.includes('GroupBy'), 'Library requests fetch items, not server group summaries.');
      assert(parameters.ViewXml.includes('Status') && parameters.ViewXml.includes('OrderBy'), 'Source restrictions and ordering are retained.');
      const offset = Number(parameters.Paging || 0);
      if (failure === 'documents' && offset) { return threshold(); }
      const rows = documents.filter(item => item.FileDirRef === parameters.FolderServerRelativeUrl && item.Status === 'Active');
      return response({ Row: rows.slice(offset, offset + 200).map(item => ({ ...item, Modified: 'localized display date' })),
        NextHref: offset + 200 < rows.length ? '?' + (offset + 200) : '',
        ListSchema: { Field: metadata.map(item => ({ Name: item.InternalName, TypeAsString: item.TypeAsString })) } });
    }
  };
  const control = new exports[project]({
    context: { pageContext: { web: { absoluteUrl: web } }, spHttpClient: client },
    views: [], viewColumns: [], gridSchemaJson: '', defaultViewId: 'view', listName: 'Documents',
    matchingDocumentFolders: false, enableDiagnostics: false, isEditMode: false,
    pageSize: 50, fetchBatchSize: 200, buttonDisplayMode: 'text', filterJson: '', showRefresh: true,
    onSelectionChange: (id, mode) => selections.push([id, mode]), onFilteredCountChange: count => counts.push(count)
  });
  control.setState = function(update, callback) {
    this.state = Object.assign({}, this.state, typeof update === 'function' ? update(this.state, this.props) : update);
    if (callback) { callback(); }
  };
  Object.assign(control.state, { accessLoading: false, hasReadAccess: true, hasEditAccess: true });
  control.loadSelectedViewFieldNames = async () => ['Title', 'Status', 'Category'];
  control.loadSelectedViewXml = async () => '<View><Query><Where><Eq><FieldRef Name="Status"/><Value Type="Text">Active</Value></Eq></Where><OrderBy><FieldRef Name="ID"/></OrderBy><GroupBy><FieldRef Name="Category"/></GroupBy></Query></View>';
  control.loadListFieldMetadataMap = async () => fieldMap;
  control.loadGridFieldMetadata = async () => fieldMap;
  control.loadLookupOptionsByField = async () => ({});
  const settle = async () => {
    for (let attempt = 0; attempt < 100; attempt++) {
      await new Promise(resolve => setImmediate(resolve));
      if (!control._libraryLoading && !control._libraryRefreshPending) { return; }
    }
    throw new Error('Library request did not settle.');
  };
  const footer = () => {
    const container = document.createElement('div');
    container.innerHTML = ReactDomServer.renderToStaticMarkup(control.render());
    return container.querySelector('.lc-total-items').textContent;
  };
  await control.loadRows();
  assert.strictEqual(control.isFolderScopedLibrary(), true, 'Unconfigured normal browsing defaults to folder scope.');
  assert.strictEqual(control.state.libraryFolders.length, 22, 'Shows immediate accessible folders, including empty ones but not the system Forms folder.');
  assert.strictEqual(control.state.rows.length, 0);
  assert.strictEqual(requests.filter(request => request.method === 'POST').length, 1, 'Root listing does not fetch 21,000 descendant documents.');
  assert.strictEqual(footer(), '0 matching files in this folder (excluding subfolders)');
  const firstFolder = libraryRoot + '/' + folderNames[0];
  requests.length = 0;
  control.navigateLibraryFolder(firstFolder);
  await settle();
  const expected = documents.filter(item => item.FileDirRef === firstFolder && item.Status === 'Active');
  assert.deepStrictEqual(Array.from(control.state.rows, item => item.ID), expected.map(item => item.ID));
  assert.strictEqual(requests.filter(request => request.method === 'POST').length, Math.ceil(expected.length / 200), 'Completes server paging only within the selected folder.');
  assert.strictEqual(requests.filter(request => request.url.includes('/Files?')).length, 5, 'Hydrates only the current folder, with server paging.');
  assert(control.state.rows.every(item => item.Modified === '2026-11-05T12:00:00Z'), 'Authoritative REST dates replace localized render dates.');
  assert.strictEqual(footer(), expected.length + ' matching files in this folder (excluding subfolders)');
  assert.strictEqual(counts[counts.length - 1], expected.length, 'Count is complete for the current folder, not a partial page or the whole library.');
  assert.deepStrictEqual(selections.pop(), [0, 'view'], 'Folder navigation clears connected selection.');
  control.state.selectedItemId = expected[0].ID;
  const selectionEvents = selections.length;
  await control.loadRows();
  assert.strictEqual(control.state.selectedItemId, expected[0].ID);
  assert.strictEqual(selections.length, selectionEvents, 'Refresh does not republish selection.');
  control.props.filterJson = JSON.stringify([{ field: 'Category', operator: 'eq', value: 'Red' }]);
  await control.loadRows();
  assert.strictEqual(control.getProcessedRows().length, expected.filter(item => item.Category === 'Red').length);
  control.props.filterJson = '';
  await control.loadRows();
  const snapshot = control.state.libraryDocuments;
  for (const error of ['documents', 'files', 'folders']) {
    failure = error;
    requests.length = 0;
    await control.loadRows();
    assert.strictEqual(control.state.libraryDocuments, snapshot, 'Failure retains the last complete snapshot: ' + error);
    assert(control.state.libraryError.includes('large-library-correlation-id'));
    assert.strictEqual(control.state.loading, false);
    if (error === 'documents') {
      assert(control.state.libraryError.includes('SPQueryThrottledException') && control.state.libraryError.includes('indexed'));
      assert.strictEqual(requests.filter(request => request.method === 'POST').length, 2, 'A failing page is not retried with other metadata headers.');
    } else {
      assert(control.state.libraryError.includes('403') && control.state.libraryError.includes('access denied'));
    }
    assert.strictEqual(footer(), expected.length + ' matching files in this folder (excluding subfolders)', 'Failure does not publish a partial count.');
  }
  assert(errors.length >= 3, 'Real failures are logged even when optional diagnostics are disabled.');
  failure = '';
  hold = new Promise(resolve => { release = resolve; });
  const stale = control.loadRows();
  await new Promise(resolve => setImmediate(resolve));
  control.navigateLibraryFolder(libraryRoot);
  hold = undefined;
  release();
  await stale;
  await settle();
  assert.strictEqual(control.state.libraryFolderLoaded, libraryRoot, 'A superseded folder load cannot overwrite a newer navigation.');
  assert.strictEqual(control.state.rows.length, 0);
  control.navigateLibraryFolder(libraryRoot + '/Empty');
  await settle();
  assert.strictEqual(control.state.libraryFolder, libraryRoot + '/Empty', 'Empty folders remain navigable.');
  assert.strictEqual(control.state.rows.length, 0);
  assert.strictEqual(footer(), '0 matching files in this folder (excluding subfolders)');
  if (project === 'GridControl') {
    control.state.editingItemId = 0;
    control.state.editingValues = { Title: 'Unsaved document' };
    control.navigateLibraryFolder(libraryRoot);
    assert.strictEqual(control.state.libraryFolder, libraryRoot + '/Empty');
    await control.loadRows();
    assert.strictEqual(control.state.editingValues.Title, 'Unsaved document');
    control.state.editingItemId = -1;
    control._libraryRefreshPending = false;
  }
  control.props.libraryLoadingMode = 'recursive';
  control.props.matchingDocumentFolders = true;
  requests.length = 0;
  await control.loadRows();
  assert.strictEqual(control.isFolderScopedLibrary(), false, 'Explicit whole-library matching retains recursive behavior.');
  assert(control.state.libraryError.includes('SPQueryThrottledException'));
  assert.strictEqual(requests.filter(request => request.method === 'POST').length, 1, 'A whole-library threshold failure is not silently capped, retried or bypassed.');
  for (const method of ['get', 'post']) {
    let attempts = 0;
    control.props.context.spHttpClient = {
      [method]: async (url, configuration, options) => {
        if (++attempts === 1) { return response({}, 406); }
        assert.strictEqual(options.headers.Accept, 'application/json;odata=verbose');
        assert.strictEqual(options.headers['OData-Version'], '3.0', 'Verbose fallback must pair with OData v3.');
        return response({});
      }
    };
    const result = method === 'get' ? await control.getJsonWithFallback(web + '/test')
      : await control.postJsonWithFallback(web + '/test', {});
    assert.strictEqual(result.ok, true, 'Metadata negotiation still retries an actual HTTP 406.');
    assert.strictEqual(attempts, 2);
  }
  await assert.rejects(() => library.loadChildFolders(web, libraryRoot, '/Other', client.get, () => true, 200), /outside/);
  await assert.rejects(() => library.loadChildFolders(web, libraryRoot, libraryRoot, async () => response({
    value: [], '@odata.nextLink': 'same'
  }), () => true, 200), /paging/);
  await assert.rejects(() => library.loadChildFolders(web, libraryRoot, libraryRoot, async () => response({
    value: [{ Name: 'Outside', ServerRelativeUrl: '/Outside' }]
  }), () => true, 200), /metadata/);
  console.log(project + ': 21,000-document folder scope, complete paging/counts, metadata, navigation, refresh, draft and error regressions passed');
}

(async () => {
  const list = fs.readFileSync(path.join(__dirname, '..', 'ListControl', 'src', 'webparts', 'listControl', 'components', 'LibraryBrowser.ts'), 'utf8');
  const grid = fs.readFileSync(path.join(__dirname, '..', 'GridControl', 'src', 'webparts', 'gridControl', 'components', 'LibraryBrowser.ts'), 'utf8');
  assert.strictEqual(list, grid, 'Keep shared library helpers identical.');
  for (const project of ['ListControl', 'GridControl']) { await run(project); }
})().catch(error => { console.error(error); process.exitCode = 1; });
