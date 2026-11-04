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
  const document = require(path.join(root, 'node_modules', 'jsdom')).jsdom('<html><body></body></html>');
  const window = document.defaultView;
  const directory = path.join(root, 'src', 'webparts', project === 'ListControl' ? 'listControl' : 'gridControl');
  const library = require('./loadLibraryBrowser')(root, window);
  let now = Date.parse('2026-11-04T16:53:06Z');
  class Clock extends Date {
    constructor(...args) { super(...(args.length ? args : [now])); }
    static now() { return now; }
  }
  let timer;
  let started = 0;
  let cleared = 0;
  window.setInterval = (callback, delay) => { started++; timer = { callback, delay }; return 42; };
  window.clearInterval = id => { assert.strictEqual(id, 42); cleared++; };
  let strings;
  vm.runInNewContext(fs.readFileSync(path.join(directory, 'loc', 'en-us.js'), 'utf8'), {
    define: (deps, factory) => { strings = factory(); }
  });
  const exports = {};
  const errors = [];
  vm.runInNewContext(ts.transpileModule(fs.readFileSync(path.join(directory, 'components', project + '.tsx'), 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
  }).outputText, { exports, document, window, Date: Clock,
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
  const Control = exports[project];
  const selections = [];
  const counts = [];
  const rootPath = '/sites/team/sub/ActualLibraryUrl';
  const metadata = [
    { InternalName: 'Title', Title: 'Title', TypeAsString: 'Text' },
    { InternalName: 'Modified', Title: 'Modified', TypeAsString: 'DateTime', DisplayFormat: 1, ReadOnlyField: true },
    { InternalName: 'Created', Title: 'Created', TypeAsString: 'DateTime', DisplayFormat: 1, ReadOnlyField: true },
    { InternalName: 'Category', Title: 'Category', TypeAsString: 'Choice', Choices: ['Red', 'Blue'] }
  ];
  const fieldMap = {};
  metadata.forEach(item => { fieldMap[item.InternalName.toLowerCase()] = {
    internalName: item.InternalName, title: item.Title, typeAsString: item.TypeAsString, displayFormat: item.DisplayFormat,
    choices: item.Choices || [], lookupList: '', lookupField: '', readOnly: !!item.ReadOnlyField
  }; });
  const filters = [
    { field: 'Modified', operator: 'ge', valueType: 'expression', value: 'now(-24h)' },
    { field: 'Modified', operator: 'le', logical: 'and', valueType: 'expression', value: 'now' }
  ];
  let documents;
  function resetDocuments() {
    documents = [0, -86400000, -86400001, 1, -1000].map((delta, index) => ({
      ID: index + 1, Title: 'Document ' + (index + 1), FSObjType: 0,
      Modified: new Date(now + delta).toISOString(), Created: '2020-01-01T00:00:00Z',
      Category: index === 4 ? 'Blue' : 'Red',
      FileDirRef: rootPath + (index === 0 ? '' : index === 1 ? '/A/B' : index === 4 ? '/X/B' : '/Old'),
      FileLeafRef: 'Document ' + (index + 1) + '.docx'
    }));
    documents.forEach(item => { item.FileRef = item.FileDirRef + '/' + item.FileLeafRef; });
  }
  resetDocuments();
  const instance = new Control({
    views: [], viewColumns: [], gridSchemaJson: '', defaultViewId: 'view', listName: 'Renamed Library',
    matchingDocumentFolders: true, filterJson: JSON.stringify(filters), enableDiagnostics: false, isEditMode: false,
    pageSize: 50, fetchBatchSize: 2, buttonDisplayMode: 'text', showRefresh: true, showLinkToItem: true, linkOpenBehavior: 'dialog',
    showDelete: true, showHistory: true, historyAvailable: true,
    context: { pageContext: { web: { absoluteUrl: 'https://example.test/sites/team/sub' } } },
    onSelectionChange: (id, mode) => selections.push([id, mode]), onFilteredCountChange: count => counts.push(count)
  });
  instance.setState = function(update, callback) {
    this.state = Object.assign({}, this.state, typeof update === 'function' ? update(this.state, this.props) : update);
    if (callback) { callback(); }
  };
  Object.assign(instance.state, { accessLoading: false, hasReadAccess: true, hasEditAccess: true });
  instance.loadSelectedViewFieldNames = async () => ['Title', 'Category'];
  instance.loadSelectedViewXml = async () => '<View><Query><Where><Eq><FieldRef Name="Status"/><Value Type="Text">Active</Value></Eq></Where><OrderBy><FieldRef Name="Title"/></OrderBy></Query><ViewFields><FieldRef Name="Title"/><FieldRef Name="Category"/></ViewFields></View>';
  instance.loadListFieldMetadataMap = async () => fieldMap;
  instance.loadGridFieldMetadata = async () => fieldMap;
  instance.loadLookupOptionsByField = async () => ({});
  const response = data => ({ ok: true, json: async () => data });
  let failFolder = false;
  let posts = 0;
  let hold;
  let release;
  let folderPages = 0;
  let typeRequests = 0;
  instance.getJsonWithFallback = async url => {
    if (url.includes('/fields?')) { return response({ value: metadata }); }
    if (url.includes('/items?')) {
      const filter = decodeURIComponent(url.split('$filter=')[1] || '');
      if (filter.startsWith('FSObjType eq 1')) {
        folderPages++;
        if (failFolder) { return { ok: false, status: 403 }; }
        const all = ['A', 'A/B', 'X', 'X/B', 'Old', 'Empty'].map(name => ({
          FileRef: rootPath + '/' + name, FileDirRef: rootPath + (name.includes('/') ? '/' + name.split('/')[0] : ''),
          FileLeafRef: name.split('/').pop()
        }));
        const offset = Number((url.match(/folderPage=(\d+)/) || [])[1] || 0);
        return response({ value: all.slice(offset, offset + 2), '@odata.nextLink': offset + 2 < all.length
          ? 'https://example.test/items?$filter=FSObjType%20eq%201&folderPage=' + (offset + 2) : '' });
      }
      return response({ value: documents.filter(item => new RegExp('\\bID eq ' + item.ID + '\\b').test(filter)) });
    }
    if (url.includes('GetFolderByServerRelativeUrl')) {
      if (failFolder) { return { ok: false, status: 403, statusText: 'Forbidden' }; }
      const folderPath = decodeURIComponent(url.match(/RelativeUrl\('(.+)'\)/)[1]).replace(/''/g, "'");
      return response({ ServerRelativeUrl: folderPath, Name: folderPath.split('/').pop() });
    }
    if (url.endsWith('?$select=BaseType')) { typeRequests++; }
    return response({ BaseType: 1, RootFolder: { ServerRelativeUrl: rootPath } });
  };
  instance.postJsonWithFallback = async (url, body) => {
    posts++;
    if (hold) { await hold; }
    assert(body.parameters.ViewXml.includes('RecursiveAll'));
    assert(body.parameters.ViewXml.includes('Status'), 'Retains source Where');
    const offset = Number(body.parameters.Paging || '0');
    return response({ Row: documents.slice(offset, offset + 2), NextHref: offset + 2 < documents.length ? '?' + String(offset + 2) : '',
      ListSchema: { Field: [{ Name: 'Title', TypeAsString: 'Text' }, { Name: 'Category', TypeAsString: 'Choice' }] } });
  };
  await instance.loadRows();
  assert.deepStrictEqual(Array.from(instance.getProcessedRows(), row => row.ID), [1, 2, 5], 'Inclusive exact millisecond window');
  assert.strictEqual(posts, 3, 'Reads beyond first batch');
  assert.strictEqual(instance.state.libraryFolders.length, 4, 'Retains old ancestor branches, not empty/old-only folders');
  assert.strictEqual(counts[counts.length - 1], 3, 'Dynamic data count excludes folders');
  let html = ReactDomServer.renderToStaticMarkup(instance.render());
  assert(html.includes('/_layouts/15/images/folder.gif'), 'Uses the native SharePoint folder icon in every toolbar mode');
  assert(html.includes('Document%201.docx'), 'Show link to item opens the document');
  assert(!html.includes('/sub/sites/'), 'No duplicated subsite prefix');
  const renderContainer = (control = instance) => {
    const container = document.createElement('div');
    container.innerHTML = ReactDomServer.renderToStaticMarkup(control.render());
    return container;
  };
  const footerText = (control = instance) => {
    const footer = renderContainer(control).querySelector('.lc-total-items');
    return footer ? footer.textContent : undefined;
  };
  assert.strictEqual(footerText(), '3 files in this folder and subfolders', 'Counts root files together with nested matching files');
  const originalDocuments = instance.state.libraryDocuments;
  instance.state.libraryDocuments = [
    { ...documents[1], ID: 70 },
    { ...documents[1], ID: 71, Title: 'Second file', FileLeafRef: 'Second.docx', FileRef: documents[1].FileDirRef + '/Second.docx' },
    { ...documents[4], ID: 72 }
  ];
  assert.strictEqual(renderContainer().querySelectorAll('tbody tr').length, 2, 'The displayed entries are two folders');
  assert.strictEqual(footerText(), '3 files in this folder and subfolders', 'Two folder rows represent three matching files below them');
  instance.props.pageSize = 1;
  instance.state.currentPage = 1;
  assert.strictEqual(renderContainer().querySelectorAll('tbody tr').length, 1, 'Pagination still pages folder entries');
  assert.strictEqual(footerText(), '3 files in this folder and subfolders', 'File total is independent of folder-row pagination');
  instance.state.libraryFolder = rootPath + '/A';
  assert.strictEqual(footerText(), '2 files in this folder and subfolders', 'Includes descendant files even when none are stored directly in this folder');
  assert(renderContainer().textContent.includes('3 matching files in the library'), 'Navigation count remains library-wide');
  instance.state.libraryFolder = rootPath + '/X/B';
  assert.strictEqual(footerText(), '1 file in this folder and subfolders', 'Singular count is correct after drill-down');
  instance.state.libraryDocuments = originalDocuments;
  instance.state.libraryFolder = rootPath;
  instance.state.currentPage = 0;
  instance.props.pageSize = 50;
  for (const mode of ['text', 'iconText', 'icon']) {
    instance.props.buttonDisplayMode = mode;
    const folderHtml = ReactDomServer.renderToStaticMarkup(instance.renderLibraryFolderRow({
      libraryFolder: { path: rootPath + '/Test Folder 1', name: 'Test Folder 1' }
    }, 8));
    const container = document.createElement('div');
    container.innerHTML = '<table><tbody>' + folderHtml + '</tbody></table>';
    const folderButton = container.querySelector('a.lc-library-folder-link');
    assert(folderButton, 'Folders are navigation links, not command buttons');
    assert(!container.querySelector('button'), 'Folder row has no button chrome');
    assert(!folderButton.classList.contains('gc-command-button-icon-only'), 'Folder names must not be squeezed into a fixed icon-only width: ' + mode);
    assert.strictEqual(folderButton.textContent.trim(), 'Test Folder 1');
    assert(!folderButton.querySelector('.gc-visually-hidden'), 'Folder name stays visible: ' + mode);
    assert(folderButton.querySelector('img.lc-library-folder-icon').getAttribute('src').endsWith('/_layouts/15/images/folder.gif'),
      'Uses SharePoint built-in folder image: ' + mode);
    assert(folderButton.getAttribute('href').includes('/Test%20Folder%201'));
    container.innerHTML = ReactDomServer.renderToStaticMarkup(instance.renderLibraryNavigation());
    const breadcrumb = container.querySelector('a[aria-current="page"]');
    assert(!breadcrumb.classList.contains('gc-command-button-icon-only'), 'Breadcrumb retains its readable label: ' + mode);
    assert(!breadcrumb.querySelector('.gc-visually-hidden'));
    assert.strictEqual(instance.getCommandButtonClass().includes('gc-command-button-icon-only'), mode === 'icon', 'Toolbar mode remains unchanged');
  }
  instance.props.buttonDisplayMode = 'text';
  instance.state.selectedItemId = 1;
  instance.state.selectedItemIds = [1];
  instance.navigateLibraryFolder(rootPath + '/A');
  const pathContainer = document.createElement('div');
  pathContainer.innerHTML = ReactDomServer.renderToStaticMarkup(instance.renderLibraryNavigation());
  assert.strictEqual(pathContainer.querySelector('.lc-library-separator').textContent, '\\', 'Breadcrumb uses backslash path separators');
  assert.strictEqual(pathContainer.querySelectorAll('a.lc-library-folder-link').length, 2);
  assert.strictEqual(instance.state.selectedItemId, 0);
  assert.deepStrictEqual(selections[selections.length - 1], [0, 'view']);
  if (project === 'GridControl') { assert.strictEqual(instance.state.selectedItemIds.length, 0); }
  const folderRow = instance.renderLibraryFolderRow({ libraryFolder: instance.state.libraryFolders[0] }, 8);
  html = ReactDomServer.renderToStaticMarkup(folderRow);
  assert(html.includes('colspan="8"'));
  assert(!html.includes('checkbox') && !html.includes('Edit') && !html.includes('History'), 'Folder has navigation only');
  instance.state.selectedItemId = 2;
  selections.length = 0;
  await instance.loadRows();
  assert.strictEqual(selections.length, 0, 'Background refresh does not republish unchanged selection');
  assert.strictEqual(instance.state.libraryFolder, rootPath + '/A');
  const previous = instance.state.libraryDocuments;
  failFolder = true;
  await instance.loadRows();
  assert.strictEqual(instance.state.libraryDocuments, previous, 'Failed ancestor resolution retains complete snapshot');
  assert.strictEqual(footerText(), '1 file in this folder and subfolders', 'Failed refresh retains the committed file count');
  assert(instance.state.libraryError.includes('403'));
  assert(errors.length > 0, 'Failure is logged');
  failFolder = false;
  instance.updateLibraryRefreshTimer();
  assert.strictEqual(timer.delay, 300000);
  now += 300000;
  documents = [documents[0], documents[4]];
  documents.push({ ...documents[0], ID: 6, Modified: new Date(now).toISOString(), Title: 'New file' });
  await timer.callback();
  assert.deepStrictEqual(Array.from(instance.getProcessedRows(), row => row.ID), [1, 5, 6], 'Timer refetches new and expired documents');
  assert.strictEqual(instance.state.libraryFolder, rootPath, 'Refresh falls back to nearest surviving ancestor');
  assert(instance.state.libraryNotice.includes(strings.RuntimeLibrarySelectionGone));
  assert.strictEqual(instance.state.selectedItemId, 2, 'Expired selection keeps its form binding');
  assert.strictEqual(selections.length, 0);
  instance.props.filterJson = '';
  await instance.loadRows();
  assert.strictEqual(instance.getProcessedRows().length, documents.length, 'No injected Modified filter');
  instance.props.filterJson = JSON.stringify([{ field: 'Category', value: ['Blue'], operator: 'eq' }]);
  await instance.loadRows();
  assert.deepStrictEqual(Array.from(instance.getProcessedRows(), row => row.ID), [5], 'Non-date and multi-value filters prune branches');
  assert.strictEqual(footerText(), '1 file in this folder and subfolders', 'Counts only files matching the configured filter');
  assert(renderContainer().textContent.includes('1 matching file in the library'), 'Library-wide count supports singular files');
  assert.strictEqual(instance.state.libraryFolders.length, 2);
  instance.props.filterJson = JSON.stringify(filters.map(filter => ({ ...filter, field: 'Created' })));
  await instance.loadRows();
  assert.strictEqual(instance.getProcessedRows().length, 0, 'Created and Modified are independently selectable');
  assert.strictEqual(footerText(), '0 files in this folder and subfolders', 'No matching documents means zero files');
  instance.props.matchingDocumentFolders = false;
  await instance.loadRows();
  assert.strictEqual(typeRequests, 1, 'Discovers document-library source even with pruning off');
  assert.strictEqual(folderPages, 3, 'Retrieves all folder pages');
  assert.strictEqual(instance.getProcessedRows().length, 0, 'Filters still apply to documents');
  assert.strictEqual(footerText(), '0 files in this folder and subfolders', 'Visible nonmatching folders never inflate the file total');
  html = ReactDomServer.renderToStaticMarkup(instance.render());
  assert(html.includes('Empty') && html.includes('Old') && html.includes('/_layouts/15/images/folder.gif'), 'Pruning off retains empty and nonmatching branches');
  const emptyFolder = instance.state.libraryFolders.find(folder => folder.name === 'Empty');
  const emptyRow = instance.renderLibraryFolderRow({ libraryFolder: emptyFolder }, 8);
  const beforeDialog = instance.state.linkDialogUrl;
  const folderLink = emptyRow.props.children.props.children;
  let prevented = false;
  folderLink.props.onClick({ preventDefault() { prevented = true; }, stopPropagation() {} });
  assert(prevented, 'Folder link prevents browser navigation to SharePoint');
  assert.strictEqual(instance.state.libraryFolder, rootPath + '/Empty', 'Folder button drills into the control');
  assert.strictEqual(instance.state.linkDialogUrl, beforeDialog, 'Folder never opens the configured document dialog');
  instance.navigateLibraryFolder(rootPath);
  folderLink.props.onKeyDown({ key: ' ', preventDefault() {}, stopPropagation() {} });
  assert.strictEqual(instance.state.libraryFolder, rootPath + '/Empty', 'Space retains accessible folder activation');
  assert.strictEqual(footerText(), '0 files in this folder and subfolders', 'Empty folders count no files');
  await instance.loadRows();
  assert.strictEqual(instance.state.libraryFolder, rootPath + '/Empty', 'Refresh preserves an accessible empty folder');
  assert.strictEqual(typeRequests, 1, 'Caches source discovery');
  const allSnapshot = instance.state.libraryDocuments;
  failFolder = true;
  await instance.loadRows();
  assert.strictEqual(instance.state.libraryDocuments, allSnapshot, 'Folder retrieval failure retains prior unpruned snapshot');
  assert(instance.state.libraryError.includes('403'));
  failFolder = false;
  instance.props.matchingDocumentFolders = true;
  await instance.loadRows();
  assert.strictEqual(instance.state.libraryFolders.length, 0, 'Pruning on removes branches without matches');
  instance.props.matchingDocumentFolders = false;
  instance.props.filterJson = '';
  hold = new Promise(resolve => { release = resolve; });
  const loading = instance.loadRows();
  await new Promise(resolve => setImmediate(resolve));
  const before = posts;
  await timer.callback();
  assert.strictEqual(posts, before, 'Requests do not overlap');
  hold = undefined; release();
  await loading;
  await new Promise(resolve => setImmediate(resolve));
  await new Promise(resolve => setImmediate(resolve));
  assert(!instance._libraryLoading);
  instance.state.libraryFolder = rootPath;
  assert.strictEqual(footerText(), '3 files in this folder and subfolders', 'Unpruned browsing counts the same matching files');
  instance.props.showLinkToItem = false;
  assert(!ReactDomServer.renderToStaticMarkup(instance.render()).includes('Document%201.docx'), 'Document link toggle is respected');
  instance.props.showLinkToItem = true;
  assert(ReactDomServer.renderToStaticMarkup(instance.render()).includes('Document%201.docx'), 'Document links work without pruning');
  if (project === 'ListControl') {
    let openedDocument;
    instance.openItemLinkDialog = url => { openedDocument = url; };
    const visit = node => {
      if (!node || typeof node !== 'object') { return; }
      if (Array.isArray(node)) { node.forEach(visit); return; }
      if (node.type === 'a' && String(node.props.href).includes('Document%201.docx')) {
        node.props.onClick({ stopPropagation() {}, preventDefault() {} });
      }
      if (node.props) { visit(node.props.children); }
    };
    visit(instance.render());
    assert(openedDocument.includes('Document%201.docx'), 'Document link retains Dialog opening behavior with pruning off');
  }
  hold = new Promise(resolve => { release = resolve; });
  const staleLoad = instance.loadRows();
  await new Promise(resolve => setImmediate(resolve));
  const oldProps = instance.props;
  const oldState = instance.state;
  instance.props = { ...oldProps, filterJson: JSON.stringify([{ field: 'Category', operator: 'eq', value: 'Blue' }]) };
  instance.componentDidUpdate(oldProps, oldState);
  hold = undefined; release();
  await staleLoad;
  for (let attempt = 0; instance._libraryLoading && attempt < 30; attempt++) {
    await new Promise(resolve => setImmediate(resolve));
  }
  assert.deepStrictEqual(Array.from(instance.getProcessedRows(), row => row.ID), [5], 'Superseded generation cannot overwrite a newer filter');
  instance.props.filterJson = '';
  await instance.loadRows();
  if (project === 'GridControl') {
    instance.state.libraryFolder = rootPath + '/X/B';
    instance.state.editingItemId = 0;
    instance.state.editingValues = { Title: 'Unsaved' };
    const before = posts;
    await timer.callback();
    assert.strictEqual(posts, before, 'Grid draft defers automatic refresh');
    instance.navigateLibraryFolder(rootPath + '/X');
    assert.strictEqual(instance.state.libraryFolder, rootPath + '/X/B');
    instance.applyActiveFilter();
    assert.strictEqual(instance.state.editingValues.Title, 'Unsaved');
    const editingState = instance.state;
    instance.cancelRowEdit();
    instance.componentDidUpdate(instance.props, editingState);
    for (let attempt = 0; instance._libraryLoading && attempt < 30; attempt++) {
      await new Promise(resolve => setImmediate(resolve));
    }
    assert.strictEqual(instance.state.libraryFolder, rootPath + '/X/B', 'Deferred refresh retains current folder after cancel');
    instance.state.editingItemId = 0;
    instance.state.libraryFolder = rootPath + '/X/B';
    let uploadUrl;
    instance.props.context.spHttpClient = { post: async url => {
      uploadUrl = url;
      return response({ ServerRelativeUrl: rootPath + '/X/B/New.docx', ListItemAllFields: { Id: 90 } });
    } };
    await instance.uploadDocumentFile({ name: 'New.docx' });
    assert(decodeURIComponent(uploadUrl).includes(rootPath + '/X/B'), 'Upload targets current folder');
    assert(uploadUrl.includes('overwrite=false'), 'Existing files cannot be overwritten/cleaned up');
    instance.state.pendingDocumentFile = { name: 'New.docx' };
    instance.validateEditingValues = () => ({});
    instance.buildEditingPayload = () => ({ Title: 'New' });
    instance.loadListItemEntityTypeName = async () => 'SP.Data.DocumentsItem';
    instance.buildVerboseEditingPayload = payload => payload;
    instance.uploadDocumentFile = async () => ({ itemId: 90, serverRelativeUrl: rootPath + '/X/B/New.docx' });
    instance.postJsonWithFallback = async () => ({ ok: false, text: async () => 'Metadata save denied' });
    const cleanup = [];
    instance.deleteUploadedDocument = async url => { cleanup.push(url); };
    await instance.saveEditingRow();
    assert.deepStrictEqual(cleanup, [rootPath + '/X/B/New.docx'], 'Only newly uploaded file is cleaned up after metadata failure');
    assert.strictEqual(instance.state.error, 'Metadata save denied');
    instance.state.editingItemId = 5;
    await instance.saveEditingRow();
    assert.strictEqual(cleanup.length, 1, 'Existing document metadata failure never deletes the file');
    instance._libraryRefreshPending = false;
  }
  instance.componentWillUnmount();
  assert(cleared > 0);
  const fresh = new Control({ ...instance.props, matchingDocumentFolders: false, filterJson: '' });
  fresh.setState = instance.setState;
  Object.assign(fresh.state, { accessLoading: false, hasReadAccess: true, hasEditAccess: true });
  ['loadSelectedViewFieldNames', 'loadSelectedViewXml', 'loadListFieldMetadataMap', 'loadGridFieldMetadata',
    'loadLookupOptionsByField', 'getJsonWithFallback'].forEach(name => { fresh[name] = instance[name]; });
  fresh.postJsonWithFallback = async (url, body) => response({ Row: documents,
    ListSchema: { Field: [{ Name: 'Title', TypeAsString: 'Text' }] } });
  const get = fresh.getJsonWithFallback;
  fresh.getJsonWithFallback = async url => url.endsWith('?$select=BaseType') ? { ok: false, status: 403 } : get(url);
  await fresh.loadRows();
  assert(fresh.state.error.includes('403'), 'Initial discovery failure is explicit; never falls back to legacy folder links');
  fresh.getJsonWithFallback = get;
  failFolder = true;
  await fresh.loadRows();
  assert(fresh.state.libraryError.includes('403'), 'Initial folder retrieval failure is explicit');
  assert.strictEqual(footerText(fresh), undefined, 'No zero-file total is shown before a successful snapshot');
  failFolder = false;
  await fresh.loadRows();
  assert.strictEqual(fresh.state.libraryFolders.length, 6, 'Fresh instances browse all folders with matching off');
  assert.strictEqual(timer.delay, 300000, 'Fresh unpruned libraries start automatic refresh');
  const settle = async () => {
    await new Promise(resolve => setImmediate(resolve));
    for (let attempt = 0; fresh._libraryLoading && attempt < 30; attempt++) {
      await new Promise(resolve => setImmediate(resolve));
    }
    assert(!fresh._libraryLoading, 'Expected folder refresh to finish');
  };
  fresh.state.selectedItemId = 5;
  selections.length = 0;
  let beforeProps = fresh.props;
  fresh.props = { ...beforeProps, matchingDocumentFolders: true };
  fresh.componentDidUpdate(beforeProps, fresh.state);
  await settle();
  assert.strictEqual(fresh.state.selectedItemId, 5, 'Pruning toggle retains document selection');
  assert.strictEqual(selections.length, 0, 'Pruning toggle does not clear connected form binding');
  if (project === 'GridControl') {
    fresh.state.editingItemId = 0;
    fresh.state.editingValues = { Title: 'Keep draft' };
  }
  beforeProps = fresh.props;
  fresh.props = { ...beforeProps, matchingDocumentFolders: false };
  fresh.componentDidUpdate(beforeProps, fresh.state);
  if (project === 'GridControl') {
    assert.strictEqual(fresh.state.editingValues.Title, 'Keep draft', 'Pruning toggle cannot discard a Grid draft');
    assert(fresh._libraryRefreshPending);
    const editingState = fresh.state;
    fresh.cancelRowEdit();
    fresh.componentDidUpdate(fresh.props, editingState);
  }
  await settle();
  assert.strictEqual(fresh.state.libraryFolders.length, 6, 'Checkbox change refreshes all folders after any deferred edit');
  let releaseDiscovery;
  fresh.props = { ...fresh.props, listName: 'Old source' };
  fresh.getJsonWithFallback = url => url.includes("getByTitle('Old%20source')?$select=BaseType")
    ? new Promise(resolve => { releaseDiscovery = resolve; }) : get(url);
  const oldDiscovery = fresh.loadRows();
  fresh.props = { ...fresh.props, listName: 'New source' };
  await fresh.loadRows();
  const currentRows = fresh.state.rows;
  releaseDiscovery(response({ BaseType: 0 }));
  await oldDiscovery;
  assert.strictEqual(fresh.state.rows, currentRows, 'Stale source discovery cannot replace new library results');
  fresh.componentWillUnmount();
  const dying = new Control({ ...fresh.props });
  dying.setState = fresh.setState;
  Object.assign(dying.state, { accessLoading: false, hasReadAccess: true });
  dying.getJsonWithFallback = () => new Promise(resolve => { releaseDiscovery = resolve; });
  const abandoned = dying.loadRows();
  dying.componentWillUnmount();
  const timersBefore = started;
  releaseDiscovery(response({ BaseType: 1 }));
  await abandoned;
  assert.strictEqual(started, timersBefore, 'Late discovery cannot start a timer after unmount');
  console.log(project + ': runtime library filtering, navigation, refresh, selection and action checks passed');
}

(async () => { for (const project of ['ListControl', 'GridControl']) { await run(project); } })().catch(error => {
  console.error(error.stack || error);
  process.exitCode = 1;
});
