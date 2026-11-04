'use strict';

const assert = require('assert');
const path = require('path');

async function run(project) {
  const root = path.join(__dirname, '..', project);
  const document = require(path.join(root, 'node_modules', 'jsdom')).jsdom('<html></html>');
  const lib = require('./loadLibraryBrowser')(root, document.defaultView);
  const time = Date.parse('2026-11-04T16:53:06Z');
  assert.throws(() => lib.validateLibraryFilterJson(['{}']), /array/);
  assert.throws(() => lib.validateLibraryFilterJson(['[{ "value": "Red" }]']), /field/);
  lib.validateLibraryFilterJson(['', '[{"field":"Modified","value":"now"}]']);
  assert.strictEqual(lib.resolveNowExpression('now(-24h)', Date.parse('2026-03-08T07:30:00Z')), '2026-03-07T07:30:00.000Z');
  assert.strictEqual(lib.resolveNowExpression('now(-24h)', time), '2026-11-03T16:53:06.000Z');
  assert.strictEqual(lib.resolveNowExpression('now(-6h)', time), '2026-11-04T10:53:06.000Z');
  assert.strictEqual(lib.resolveNowExpression('now(-48h)', time), '2026-11-02T16:53:06.000Z');
  assert.strictEqual(lib.resolveNowExpression('now(+2h)', time), '2026-11-04T18:53:06.000Z');
  assert.strictEqual(lib.resolveNowExpression('now', time), new Date(time).toISOString());
  ['now(-1.5h)', 'now(-24)', 'now(Infinityh)', 'now(999999999999999999h)'].forEach(value => {
    assert.throws(() => lib.resolveNowExpression(value, time));
  });
  const web = 'https://example.test/sites/team/sub';
  assert.strictEqual(lib.documentUrl(web, '/sites/team/Docs/A%20B.docx'), 'https://example.test/sites/team/Docs/A%20B.docx');
  assert.strictEqual(lib.documentUrl(web, 'Docs/A #&+.docx'), web + '/Docs/A%20%23&+.docx');
  assert.strictEqual(lib.documentUrl(web, 'javascript:alert(1)'), '');
  assert.strictEqual(lib.documentUrl(web, '//elsewhere.test/file'), '');
  assert.strictEqual(lib.documentUrl(web, '/Docs/100%.docx'), 'https://example.test/Docs/100%25.docx');
  assert.strictEqual(lib.documentUrl(web, '/Docs/%E6%96%87.docx'), 'https://example.test/Docs/%E6%96%87.docx');
  assert.throws(() => lib.ancestorPaths('/Docs', [{ FileDirRef: '/Other/Docs' }]));
  assert.throws(() => lib.ancestorPaths('/Docs', [{ FileDirRef: '/Docs/../Other' }]));
  const documents = [{ ID: 1, FileDirRef: '/Docs' }, { ID: 2, FileDirRef: '/Docs/A/B' },
    { ID: 3, FileDirRef: '/Docs/X/B' }];
  assert.deepStrictEqual(Array.from(lib.ancestorPaths('/Docs', documents)).sort(), ['/Docs/A', '/Docs/A/B', '/Docs/X', '/Docs/X/B']);
  const folders = ['/Docs/A', '/Docs/A/B', '/Docs/X', '/Docs/X/B'].map(path => ({ path, name: path.split('/').pop() }));
  const level = lib.folderLevel('/Docs', '/Docs', documents, folders);
  assert.strictEqual(lib.folderDocumentCount('/Docs', documents), 3);
  assert.strictEqual(lib.folderDocumentCount('/Docs/A', documents), 1);
  assert.strictEqual(lib.folderDocumentCount('/Docs/A/', documents), 1);
  assert.strictEqual(lib.folderDocumentCount('/docs/x', documents), 1);
  assert.strictEqual(lib.folderDocumentCount('/Docs/Empty', documents), 0);
  assert.strictEqual(lib.folderDocumentCount('/Docs/A', [{ FileDirRef: '/Docs/Another' }]), 0, 'Folder paths use segment boundaries');
  assert.strictEqual(level.length, 3);
  assert.strictEqual(level[0].libraryFolder.path, '/Docs/A');
  assert.strictEqual(level[2].ID, 1);
  assert.strictEqual(lib.folderLevel('/Docs', '/Docs/A', documents, folders)[0].libraryFolder.path, '/Docs/A/B');
  const allFolders = folders.concat([{ path: '/Docs/Empty', name: 'Empty' }]);
  assert(lib.folderLevel('/Docs', '/Docs', documents, allFolders, false).some(row => row.libraryFolder && row.libraryFolder.name === 'Empty'),
    'Without pruning, empty folders remain navigable');
  assert(!lib.folderLevel('/Docs', '/Docs', documents, allFolders, true).some(row => row.libraryFolder && row.libraryFolder.name === 'Empty'));
  assert.strictEqual(lib.nearestFolder('/Docs', '/Docs/A/B', ['/Docs/A']), '/Docs/A');
  assert.strictEqual(lib.nearestFolder('/Docs', '/Docs/A/B', []), '/Docs');
  const xml = lib.recursiveView('<View Scope="FilesOnly"><Query><Where><Eq><FieldRef Name="Status"/><Value Type="Text">Active</Value></Eq></Where><OrderBy><FieldRef Name="Title"/></OrderBy><GroupBy><FieldRef Name="Status"/></GroupBy></Query><ViewFields/></View>', ['Modified'], [
    { field: 'Modified', operator: 'ge', logical: 'and', valueType: 'expression', value: 'now(-24h)' },
    { field: 'Modified', operator: 'le', logical: 'and', valueType: 'expression', value: 'now' }
  ], time, 50);
  assert(xml.includes('RecursiveAll') && xml.includes('Status') && xml.includes('OrderBy'));
  assert(xml.includes('2026-11-03T16:53:06.000Z') && xml.includes('StorageTZ="TRUE"'));
  assert(!xml.includes('GroupBy'));
  const orXml = lib.recursiveView('<View/>', [], [
    { field: 'Modified', operator: 'ge', valueType: 'expression', value: 'now(-24h)' },
    { field: 'Title', logical: 'or', value: 'Always include' }
  ], time, 50, ['Modified']);
  assert(!orXml.includes('DateTime'), 'Mixed OR cannot be narrowed to one date conjunct');
  const typedXml = lib.recursiveView('<View/>', [], [
    { field: 'Title', valueType: 'expression', value: 'now' }
  ], time, 50, ['Modified']);
  assert(!typedXml.includes('DateTime'), 'Do not push a text-field expression as a DateTime predicate');
  const response = data => ({ ok: true, json: async () => data });
  assert.strictEqual(await lib.detectDocumentLibrary(web, 'Docs', async () => response({ BaseType: 1 })), true);
  assert.strictEqual(await lib.detectDocumentLibrary(web, 'Tasks', async () => response({ d: { BaseType: 0 } })), false);
  await assert.rejects(() => lib.detectDocumentLibrary(web, 'Docs', async () => response({})), /BaseType/);
  await assert.rejects(() => lib.detectDocumentLibrary(web, 'Docs', async () => ({ ok: false, status: 403 })), /403/);
  const folderItem = path => ({ FileRef: path, FileDirRef: path.substring(0, path.lastIndexOf('/')), FileLeafRef: path.split('/').pop() });
  const pages = [];
  const loadedFolders = await lib.loadLibraryFolders(web, "Team's Docs", '/Docs', documents, async url => {
    pages.push(url);
    return response(pages.length === 1
      ? { value: ['/Docs/A', '/Docs/A/B'].map(folderItem), '@odata.nextLink': 'second' }
      : { d: { results: ['/Docs/X', '/Docs/X/B', '/Docs/Empty'].map(folderItem) } });
  }, () => true, 2);
  assert.strictEqual(pages.length, 2, 'Completes folder paging independently of document filters');
  assert.strictEqual(loadedFolders.length, 5);
  assert.strictEqual(lib.folderLevel('/Docs', '/Docs', [], loadedFolders, false).length, 3, 'Empty filtered results retain folders');
  await assert.rejects(() => lib.loadLibraryFolders(web, 'Docs', '/Docs', [], async () => response({
    value: [], '@odata.nextLink': 'repeat'
  }), () => true, 2), /paging/);
  await assert.rejects(() => lib.loadLibraryFolders(web, 'Docs', '/Docs', [], async () => ({ ok: false, status: 403 }), () => true, 2), /403/);
  await assert.rejects(() => lib.loadLibraryFolders(web, 'Docs', '/Docs', [], async () => response({
    value: [folderItem('/Other/Outside')]
  }), () => true, 2), /metadata/);
  await assert.rejects(() => lib.loadLibraryFolders(web, 'Docs', '/Docs', documents, async () => response({ value: [] }), () => true, 2), /ancestor/);
  await assert.rejects(() => lib.loadLibraryFolders(web, 'Docs', '/Docs', [], async () => response({ value: [] }), () => false, 2), /Superseded/);
  let posts = 0;
  const options = {
    webUrl: web, listName: "Team's Docs", viewXml: '<View/>', fields: [],
    conditions: [], time, batchSize: 50, current: () => true,
    get: async url => response(url.includes('/fields') ? { value: [] } : url.includes('/items?')
      ? { value: [{ ID: 1, FileDirRef: '/Docs' }, { ID: 2, FileDirRef: '/Docs' }] }
      : { BaseType: 1, RootFolder: { ServerRelativeUrl: '/Docs' } }),
    post: async () => { posts++; return response({ rows: [{ ID: posts, FileDirRef: '/Docs' }], nextHref: posts === 1 ? '?p_ID=1' : '', fields: [] }); },
    extract: data => data
  };
  const result = await lib.loadLibraryDocuments(options);
  assert.strictEqual(result.rows.length, 2, 'Completes all candidate pages before publishing');
  assert.strictEqual(posts, 2);
  posts = 0;
  options.post = async () => ++posts === 1 ? response({ rows: [], fields: [], nextHref: '?p_ID=1' })
    : { ok: false, status: 403, statusText: 'Forbidden' };
  await assert.rejects(() => lib.loadLibraryDocuments(options), /403/);
  options.post = async () => response({ rows: [], fields: [], nextHref: '?same' });
  await assert.rejects(() => lib.loadLibraryDocuments(options), /paging/i);
  options.post = async () => response({ d: { RenderListDataAsStream: JSON.stringify({ ErrorInfo: { ErrorMessage: 'View threshold exceeded' } }) } });
  await assert.rejects(() => lib.loadLibraryDocuments(options), /threshold/);
  await assert.rejects(() => lib.loadAncestorFolders(web, '/Docs', documents, async () => ({ ok: false, status: 403 })), /403/);
  console.log(project + ': library expression, URL, hierarchy, query, paging and permission checks passed');
}

(async () => { for (const project of ['ListControl', 'GridControl']) { await run(project); } })().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
