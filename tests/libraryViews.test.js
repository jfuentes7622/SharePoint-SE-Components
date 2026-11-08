'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const loadLibraryBrowser = require('./loadLibraryBrowser');

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
  const library = loadLibraryBrowser(root, window);
  let strings;
  vm.runInNewContext(fs.readFileSync(path.join(directory, 'loc', 'en-us.js'), 'utf8'), {
    define: (dependencies, factory) => { strings = factory(); }
  });
  const errors = [];
  function loadComponent(name) {
    const exports = {};
    vm.runInNewContext(ts.transpileModule(fs.readFileSync(path.join(directory, 'components', name + '.tsx'), 'utf8'), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
    }).outputText, {
      exports, window, document, Date, DOMParser: window.DOMParser,
      XMLSerializer: function() { this.serializeToString = loadLibraryBrowser.serializeXml; },
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
    return exports[name];
  }
  const Control = loadComponent(project);
  const web = 'https://example.test/sites/team';
  const libraryRoot = '/sites/team/Documents';
  const ids = {
    all: '11111111-1111-4111-8111-111111111111',
    active: '22222222-2222-4222-8222-222222222222',
    ascending: '33333333-3333-4333-8333-333333333333',
    folder: '44444444-4444-4444-8444-444444444444',
    filesOnly: '55555555-5555-4555-8555-555555555555',
    grouped: '66666666-6666-4666-8666-666666666666'
  };
  const activeQuery = '<Where><Eq><FieldRef Name="Status"/><Value Type="Text">Active</Value></Eq></Where>';
  const views = {
    [ids.all]: { fields: ['Title'], Scope: 0, ViewQuery: '<OrderBy><FieldRef Name="Title"/></OrderBy>' },
    [ids.active]: { fields: ['Status', 'Title'], Scope: 1, ViewQuery: activeQuery + '<OrderBy><FieldRef Name="ID" Ascending="FALSE"/></OrderBy>' },
    [ids.ascending]: { fields: ['Title', 'Category'], Scope: 2, ViewQuery: activeQuery + '<OrderBy><FieldRef Name="ID" Ascending="TRUE"/></OrderBy>' },
    [ids.folder]: { fields: ['Status', 'Title'], Scope: 0, ViewQuery: activeQuery + '<OrderBy><FieldRef Name="ID" Ascending="FALSE"/></OrderBy>' },
    [ids.filesOnly]: { fields: ['Title'], Scope: 3, ViewQuery: activeQuery },
    [ids.grouped]: { fields: ['Title', 'Status'], Scope: 1, ViewQuery: activeQuery
      + '<OrderBy><FieldRef Name="ID"/></OrderBy><GroupBy Collapse="TRUE"><FieldRef Name="Category"/><FieldRef Name="Status"/></GroupBy>' }
  };
  const documents = Array.from({ length: 21000 }, (_, index) => {
    const folder = libraryRoot + '/Team ' + Math.floor(index / 1000);
    return { ID: index + 1, FSObjType: 0, Title: 'Document ' + (index + 1), Status: index % 1000 === 0 && index < 5000 ? 'Active' : 'Archived',
      Category: 'Red', FileDirRef: folder, FileLeafRef: 'File ' + (index + 1) + '.docx',
      FileRef: folder + '/File ' + (index + 1) + '.docx', Modified: '2026-11-06T12:00:00Z' };
  });
  const metadata = ['Title', 'Status', 'Category', 'Modified'].map(name => ({
    InternalName: name, Title: name, TypeAsString: name === 'Modified' ? 'DateTime' : 'Text'
  }));
  const requests = [];
  const counts = [];
  const selections = [];
  let hold;
  let release;
  let failure = '';
  const response = (data, status = 200) => ({
    ok: status < 400, status, statusText: status >= 400 ? 'Failed' : 'OK',
    json: async () => data, text: async () => JSON.stringify(data),
    headers: { get: name => name === 'SPRequestGuid' ? 'view-correlation-id' : undefined },
    clone() { return response(data, status); }
  });
  const client = {
    async get(url) {
      requests.push({ method: 'GET', url });
      if (url.includes('/views')) {
        const key = Object.values(ids).find(id => decodeURIComponent(url).includes(id));
        assert(key, 'Selected view GUID must be present in metadata URL.');
        if (failure === 'view') { return response({ error: 'View inaccessible' }, 403); }
        return response(url.includes('/ViewFields') ? { value: views[key].fields }
          : { ViewQuery: views[key].ViewQuery, Scope: views[key].Scope, RowLimit: 2 });
      }
      if (url.includes('/fields?')) { return response({ value: metadata }); }
      if (url.includes('/DefaultView?')) { return response({ Id: ids.all }); }
      const itemMatch = /\/items\((\d+)\)/.exec(url);
      if (itemMatch) {
        if (failure === 'item') { return response({ error: 'Document inaccessible' }, 403); }
        return response(documents[Number(itemMatch[1]) - 1]);
      }
      if (url.includes('/items?')) {
        const filter = decodeURIComponent(url.split('$filter=')[1] || '');
        assert(filter.startsWith('FSObjType eq 1'), 'No library-wide item-ID OR hydration is permitted for flat views.');
        return response({ value: Array.from({ length: 21 }, (_, index) => ({
          FileRef: libraryRoot + '/Team ' + index, FileDirRef: libraryRoot, FileLeafRef: 'Team ' + index
        })) });
      }
      if (url.includes('GetFolderByServerRelativeUrl')) {
        const folder = decodeURIComponent(url.match(/RelativeUrl\('(.+?)'\)/)[1]);
        if (url.includes('/Folders?')) {
          return response({ value: folder === libraryRoot ? Array.from({ length: 21 }, (_, index) => ({
            Name: 'Team ' + index, ServerRelativeUrl: libraryRoot + '/Team ' + index
          })) : [] });
        }
        if (url.includes('/Files?')) {
          return response({ value: documents.filter(item => item.FileDirRef === folder).map(item => ({ ListItemAllFields: item })) });
        }
        return response({ Name: folder.split('/').pop(), ServerRelativeUrl: folder });
      }
      if (url.includes('BaseType')) { return response({ BaseType: 1, RootFolder: { ServerRelativeUrl: libraryRoot } }); }
      throw new Error('Unexpected GET ' + url);
    },
    async post(url, configuration, options) {
      const parameters = JSON.parse(options.body).parameters;
      requests.push({ method: 'POST', url, parameters });
      if (hold) { await hold; }
      if (failure === 'threshold') { return response({ error: 'SPQueryThrottledException: list view threshold' }, 500); }
      const xml = new window.DOMParser().parseFromString(parameters.ViewXml, 'text/xml');
      const scope = xml.documentElement.getAttribute('Scope');
      const recursive = scope === 'Recursive' || scope === 'RecursiveAll';
      const folder = parameters.FolderServerRelativeUrl || libraryRoot;
      let rows = documents.filter(item => recursive ? item.FileDirRef === folder || item.FileDirRef.startsWith(folder + '/') : item.FileDirRef === folder);
      if (xml.getElementsByTagName('Value').length > 1) {
        assert(parameters.ViewXml.includes('Active'), 'View filter reaches the server, not a default-view query.');
        rows = rows.filter(item => item.Status === 'Active');
      }
      const order = xml.getElementsByTagName('OrderBy')[0];
      if (order) {
        const descending = order.firstElementChild.getAttribute('Ascending') === 'FALSE';
        rows.sort((a, b) => descending ? b.ID - a.ID : a.ID - b.ID);
      }
      assert(rows.length <= 1000, 'No unrestricted whole-library fetch during a selective view switch.');
      const offset = Number(parameters.Paging || 0);
      const names = Array.from(xml.getElementsByTagName('ViewFields')[0].getElementsByTagName('FieldRef'), node => node.getAttribute('Name'));
      return response({ Row: rows.slice(offset, offset + 2).map(item => {
        const row = { ID: item.ID };
        names.forEach(name => { row[name] = name === 'Modified' ? 'localized' : item[name]; });
        return row;
      }),
        NextHref: offset + 2 < rows.length ? '?' + (offset + 2) : '',
        ListSchema: { Field: metadata.filter(item => names.includes(item.InternalName)).map(item => ({ Name: item.InternalName, DisplayName: item.Title })) } });
    }
  };
  const control = new Control({
    context: { pageContext: { web: { absoluteUrl: web } }, spHttpClient: client },
    views: Object.values(ids).map(key => ({ key, text: key, isDefault: key === ids.all })),
    defaultViewId: ids.active, viewColumns: [], gridSchemaJson: '', listName: 'Documents',
    libraryLoadingMode: 'folder', matchingDocumentFolders: false, enableDiagnostics: false, isEditMode: false,
    pageSize: 50, fetchBatchSize: 2, filterJson: '', showRefresh: true, buttonDisplayMode: 'text',
    onSelectionChange: (id, mode) => selections.push([id, mode]), onFilteredCountChange: count => counts.push(count)
  });
  Object.assign(control.state, { accessLoading: false, hasReadAccess: true, hasEditAccess: true });
  control.setState = function(update, callback) {
    const previousState = this.state;
    this.state = Object.assign({}, this.state, typeof update === 'function' ? update(this.state, this.props) : update);
    this.componentDidUpdate(this.props, previousState);
    if (callback) { callback(); }
  };
  const settle = async () => {
    for (let attempt = 0; attempt < 100; attempt++) {
      await new Promise(resolve => setImmediate(resolve));
      if (!control._libraryLoading && !control._libraryRefreshPending) { return; }
    }
    throw new Error('View request did not settle.');
  };
  const configure = async (update, wait = true) => {
    const oldProps = control.props;
    control.props = Object.assign({}, oldProps, update);
    control.componentDidUpdate(oldProps, control.state);
    if (wait) { await settle(); }
  };
  const markup = () => {
    const container = document.createElement('div');
    container.innerHTML = ReactDomServer.renderToStaticMarkup(control.render());
    return container;
  };
  const rowIds = () => Array.from(control.state.rows, row => row.ID);
  await control.loadRows();
  assert.deepStrictEqual(rowIds(), [4001, 3001, 2001, 1001, 1], 'Initial configured view is used instead of the library default.');
  await configure({ defaultViewId: ids.all });
  assert.strictEqual(control.state.libraryViewFlat, false);
  assert.strictEqual(control.state.libraryFolders.length, 21);
  assert.strictEqual(markup().querySelectorAll('tbody tr.lc-library-folder').length, 0, 'Navigation folders never masquerade as filtered results.');
  assert.strictEqual(markup().querySelectorAll('details.lc-library-folders li').length, 21, 'Accessible folders remain available in explicitly separate navigation.');

  requests.length = 0;
  await configure({ defaultViewId: ids.active });
  assert.strictEqual(control.state.libraryError, '');
  assert.deepStrictEqual(rowIds(), [4001, 3001, 2001, 1001, 1], 'A configured filtered flat view shows exactly five documents across folders in source-view order.');
  assert.strictEqual(control.state.libraryViewFlat, true);
  assert.strictEqual(markup().querySelectorAll('tbody tr').length, 5);
  assert.strictEqual(markup().querySelector('.lc-total-items').textContent, '5 matching files in this view (including subfolders)');
  assert.strictEqual(counts[counts.length - 1], 5);
  assert.strictEqual(requests.filter(request => request.method === 'POST').length, 3, 'All filtered view pages complete.');
  assert.strictEqual(requests.filter(request => /\/items\(\d+\)/.test(request.url)).length, 5, 'Only returned IDs are hydrated.');
  assert(!requests.some(request => request.url.includes('/Files?')), 'A flat filtered view does not enumerate every file in matching folders.');
  assert(control.state.rows.every(row => row.Modified === '2026-11-06T12:00:00Z'));
  assert.deepStrictEqual(Array.from(control.getDisplayFields(), field => field.RealFieldName || field.Name), ['Status', 'Title']);

  control.state.sortFieldName = 'ID';
  control.state.sortDirection = 'desc';
  control.setState({ selectedViewId: ids.ascending });
  await settle();
  assert.deepStrictEqual(rowIds(), [1, 1001, 2001, 3001, 4001], 'Runtime view changes use their own ordering.');
  assert.strictEqual(control.state.sortFieldName, '', 'A previous user sort cannot mask the new view ordering.');
  assert.deepStrictEqual(Array.from(control.getDisplayFields(), field => field.RealFieldName || field.Name), ['Title', 'Category']);

  await configure({ defaultViewId: ids.folder });
  assert.strictEqual(control.state.libraryViewFlat, false);
  assert.deepStrictEqual(rowIds(), [], 'A folder-based view remains scoped at root instead of silently querying descendants.');
  control.navigateLibraryFolder(libraryRoot + '/Team 0');
  await settle();
  assert.deepStrictEqual(rowIds(), [1]);
  assert.strictEqual(markup().querySelector('.lc-total-items').textContent, '1 matching file in this folder (excluding subfolders)');
  await configure({ defaultViewId: ids.filesOnly });
  assert.deepStrictEqual(rowIds(), []);
  assert.strictEqual(requests.filter(request => request.method === 'POST').pop().parameters.ViewXml.includes('Scope="FilesOnly"'), true);

  hold = new Promise(resolve => { release = resolve; });
  await configure({ defaultViewId: ids.active }, false);
  await new Promise(resolve => setImmediate(resolve));
  await configure({ defaultViewId: ids.ascending }, false);
  hold = undefined;
  release();
  await settle();
  assert.deepStrictEqual(rowIds(), [1, 1001, 2001, 3001, 4001], 'A pending old-view response cannot overwrite a newer settings selection.');

  await configure({ libraryLoadingMode: 'recursive', matchingDocumentFolders: true });
  assert.deepStrictEqual(rowIds(), [1, 1001, 2001, 3001, 4001]);
  assert.strictEqual(markup().querySelectorAll('tbody tr').length, 5, 'Whole-library matching also displays a flat view without synthetic folder rows.');
  const cachedRequests = requests.length;
  control.navigateLibraryFolder(libraryRoot + '/Team 0');
  assert.deepStrictEqual(Array.from(control.getProcessedRows(), row => row.ID), [1], 'A cached flat snapshot narrows to the chosen folder subtree.');
  assert.strictEqual(markup().querySelector('.lc-total-items').textContent, '1 matching file in this view (including subfolders)');
  control.navigateLibraryFolder(libraryRoot);
  assert.strictEqual(requests.length, cachedRequests, 'Whole-library mode navigates its completed snapshot without another scan.');
  if (project === 'GridControl') {
    control.setState({ editingItemId: 1, editingValues: { Title: 'Unsaved title' } });
    await configure({ defaultViewId: ids.active }, false);
    assert.strictEqual(control.state.editingValues.Title, 'Unsaved title');
    assert.strictEqual(control.state.selectedViewId, ids.ascending, 'A settings change waits for the Grid draft.');
    control.setState({ editingItemId: -1 });
    await settle();
    assert.deepStrictEqual(rowIds(), [4001, 3001, 2001, 1001, 1]);
  }
  {
    const Designer = loadComponent(project === 'ListControl' ? 'ListDesigner' : 'GridDesigner');
    let importedSchema;
    let importedColumns;
    const designer = new Designer({
      context: control.props.context, listName: 'Documents', viewId: ids.grouped, schemaJson: '',
      viewColumns: [], groupingJson: '', enableDiagnostics: false,
      onSave: (value, groupingJson) => {
        if (project === 'ListControl') { importedColumns = value; importedSchema = groupingJson; }
        else { importedSchema = value; }
      }
    });
    designer.setState = function(update) { this.state = Object.assign({}, this.state, update); };
    await designer.loadFields();
    await designer.loadFromView();
    designer.save();
    assert(importedSchema, 'Library Designer import produces a saved grouping schema.');
    await configure({ libraryLoadingMode: 'folder', matchingDocumentFolders: false });
    await configure(Object.assign({ defaultViewId: ids.grouped }, project === 'ListControl'
      ? { groupingJson: importedSchema, viewColumns: importedColumns } : { gridSchemaJson: importedSchema }));
    assert.strictEqual(control.state.libraryError, '');
    const grouping = project === 'ListControl' ? control.getGroupingConfig() : control.getGridGroupingConfig();
    assert.strictEqual(grouping.enabled, true, 'Folder support must not disable Designer grouping for libraries.');
    assert.deepStrictEqual(rowIds(), [1, 1001, 2001, 3001, 4001]);
    assert(control.state.rows.every(row => row.Category === 'Red' && row.Title), 'A hidden grouping field and actual item cells are loaded.');
    assert(control.state.fields.some(field => field.Name === 'Category' && field.RuntimeFilterOnly), 'Hidden group keys stay out of displayed columns.');
    const groups = control.buildRowGroups(control.getProcessedRows(), grouping);
    assert.strictEqual(groups[0].label, 'Red');
    assert.strictEqual(groups[0].subGroups[0].label, 'Active');
    assert.strictEqual(control.countGroupRows(groups[0]), 5);
    const groupClass = 'gc-group-row';
    assert.strictEqual(markup().querySelectorAll('tbody tr.' + groupClass).length, 1, 'Imported groups start collapsed.');
    control.toggleGroupCollapsed(groups[0].key, true);
    control.toggleGroupCollapsed(groups[0].key + '/' + groups[0].subGroups[0].key, true);
    assert.strictEqual(markup().querySelectorAll('tbody tr:not(.' + groupClass + ')').length, 5, 'Expanding groups reveals all five real documents.');
    assert(markup().querySelector('tbody').textContent.includes('Document 1'), 'Expanded document cells contain their actual Title values.');
    control.navigateLibraryFolder(libraryRoot + '/Team 0');
    await settle();
    assert.deepStrictEqual(rowIds(), [1], 'Grouped flat views retain current-folder navigation scope.');
    control.navigateLibraryFolder(libraryRoot);
    await settle();
    const singleGroupSchema = JSON.parse(importedSchema);
    const singleGrouping = project === 'ListControl' ? singleGroupSchema : singleGroupSchema.grouping;
    singleGrouping.field2 = '';
    singleGrouping.collapsedByDefault = false;
    await configure(project === 'ListControl' ? { groupingJson: JSON.stringify(singleGroupSchema) }
      : { gridSchemaJson: JSON.stringify(singleGroupSchema) });
    assert.strictEqual(markup().querySelectorAll('tbody tr.' + groupClass).length, 1, 'Single-level groups are supported.');
    assert.strictEqual(markup().querySelectorAll('tbody tr:not(.' + groupClass + ')').length, 5, 'Expanded-by-default groups show populated records immediately.');
    await configure({ defaultViewId: ids.folder });
    assert.strictEqual(markup().querySelectorAll('tbody tr.' + groupClass).length, 0, 'An empty folder does not fabricate groups.');
    assert.strictEqual(markup().querySelectorAll('details.lc-library-folders li').length, 21, 'Empty grouped results retain separate folder navigation.');
    control.navigateLibraryFolder(libraryRoot + '/Team 0');
    await settle();
    assert.deepStrictEqual(rowIds(), [1], 'Folder-based views group only the current folder documents.');
    assert.strictEqual(markup().querySelectorAll('tbody tr.' + groupClass).length, 1);
    assert(markup().querySelector('tbody').textContent.includes('Document 1'));
    await configure({ defaultViewId: ids.active });
  }
  if (project === 'GridControl') {
    await configure({ gridSchemaJson: JSON.stringify({ steps: [{ fields: [{ id: 'category', fieldName: 'Category', type: 'text' }] }] }) });
    assert.deepStrictEqual(Array.from(control.getDisplayFields(), field => field.RealFieldName || field.Name), ['FileLeafRef', 'Category'], 'Explicit Grid schema columns retain precedence over source-view columns.');
    assert.deepStrictEqual(rowIds(), [4001, 3001, 2001, 1001, 1], 'Grid schema does not replace the source-view filter or ordering.');
  }
  for (const kind of ['item', 'threshold', 'view']) {
    const previous = control.state.libraryDocuments;
    failure = kind;
    requests.length = 0;
    await control.loadRows();
    assert.strictEqual(control.state.libraryDocuments, previous, 'A failed selected view keeps only the last complete snapshot: ' + kind);
    assert(control.state.libraryError, 'Failures must be visible: ' + kind);
    if (kind === 'threshold') {
      assert(control.state.libraryError.includes('view-correlation-id'));
      assert.strictEqual(requests.filter(request => request.method === 'POST').length, 1, 'No threshold retry or All Documents fallback.');
    }
    if (kind === 'view') { assert.strictEqual(requests.filter(request => request.method === 'POST').length, 0, 'Missing view metadata must not fall back to default results.'); }
    failure = '';
  }
  assert(errors.length >= 3, 'Genuine failures remain logged.');
  assert(selections.length > 0, 'Source/folder changes continue publishing empty connected selection.');
  console.log(project + ': actual metadata, five-of-21k filtered views, settings/runtime switching, scope/order/columns, stale requests, separate navigation, drafts and explicit failures passed');
}

(async () => {
  await run('ListControl');
  await run('GridControl');
})().catch(error => { console.error(error); process.exitCode = 1; });
