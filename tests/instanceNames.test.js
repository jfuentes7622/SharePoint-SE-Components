'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const projects = [
  ['ListControl', 'listControl', 'ListControlWebPart', 'ListControl'],
  ['GridControl', 'gridControl', 'GridControlWebPart', 'GridControl'],
  ['DynamicForms', 'sharePointDynamicForm', 'SharePointDynamicFormWebPart', 'DynamicForm'],
  ['ReportForms', 'sharePointDynamicForm', 'SharePointDynamicFormWebPart', 'ReportForm']
];
const id = number => '00000000-0000-4000-8000-' + String(number).padStart(12, '0');
const response = data => ({ ok: true, json: async () => data });
const canvasControl = (number, name) => ({ id: id(number), webPartData: { instanceId: id(number), properties: { instanceName: name } } });

async function run(project) {
  const [name, folder, className, prefix] = project;
  const root = path.join(__dirname, '..', name);
  const helper = require('./loadInstanceNames')(root);
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const jsdom = require(path.join(root, 'node_modules', 'jsdom'));
  let strings;
  vm.runInNewContext(fs.readFileSync(path.join(root, 'src', 'webparts', folder, 'loc', 'en-us.js'), 'utf8'), {
    define: (deps, factory) => { strings = factory(); }
  });
  function fixture(number, document) {
    document = document || jsdom.jsdom('<!doctype html><body></body>', { url: 'https://example.test/SitePages/Test.aspx' });
    const element = document.createElement('div');
    document.body.appendChild(element);
    const context = {
      instanceId: id(number),
      manifest: { alias: name + 'WebPart', id: id(95) },
      pageContext: { web: { absoluteUrl: 'https://example.test' }, list: { id: { toString: () => id(90) } }, listItem: { id: 7 } },
      propertyPane: { refresh() {} },
      spHttpClient: { get: async () => response({ d: { CanvasContent1: '[]' } }) }
    };
    const properties = {};
    return { document, element, context, properties, ensure: mode => helper.ensureInstanceName(properties, prefix, context, element, () => mode === undefined ? 2 : mode) };
  }
  const first = fixture(1);
  const numbered = fixture(13);
  const nextNumbered = fixture(14, numbered.document);
  await numbered.ensure();
  await nextNumbered.ensure();
  assert.strictEqual(numbered.properties.instanceName, prefix + '1');
  assert.strictEqual(nextNumbered.properties.instanceName, prefix + '2', 'New controls receive sequential numbered defaults');
  let requests = 0;
  first.context.spHttpClient.get = async url => {
    requests++;
    assert(url.includes("/lists(guid'") && url.endsWith('/items(7)'));
    return response({ CanvasContent1: JSON.stringify([canvasControl(50, prefix + '1')]) });
  };
  await first.ensure(1);
  assert.strictEqual(first.properties.instanceName, undefined, 'Read mode never changes saved properties');
  assert.strictEqual(requests, 0);
  const hidden = first.document.createElement('div');
  hidden.style.display = 'none';
  hidden.setAttribute('data-sp-controldata', JSON.stringify(canvasControl(51, prefix + '2')));
  first.document.body.appendChild(hidden);
  const second = fixture(2, first.document);
  await Promise.all([first.ensure(), second.ensure()]);
  assert.strictEqual(first.properties.instanceName, prefix + '3', 'Reserves saved and hidden control names');
  assert.strictEqual(second.properties.instanceName, prefix + '4', 'Concurrent allocation cannot collide');
  assert.strictEqual(requests, 1, 'Page metadata is loaded once per page');
  assert.strictEqual(first.element.getAttribute('data-spse-instance-name'), prefix + '3');
  await first.ensure();
  assert.strictEqual(first.properties.instanceName, prefix + '3', 'No renumbering after reinitialization');
  const saved = JSON.parse(JSON.stringify(first.properties));
  const reopened = fixture(1);
  Object.assign(reopened.properties, saved);
  reopened.context.spHttpClient.get = () => { throw new Error('Named controls do not need an allocation query'); };
  await reopened.ensure();
  assert.deepStrictEqual(reopened.properties, saved, 'Persisted names and provenance survive reload');
  const clone = fixture(3, first.document);
  Object.assign(clone.properties, saved);
  await clone.ensure();
  assert.strictEqual(clone.properties.instanceName, prefix + '5', 'Copied generated names get their own number');
  assert.strictEqual(clone.properties.instanceNameGeneratedFor, id(3));
  first.properties.instanceName = 'My custom name';
  delete first.properties.instanceNameGeneratedFor;
  await first.ensure();
  assert.strictEqual(first.properties.instanceName, 'My custom name');
  assert.strictEqual(first.element.getAttribute('data-spse-instance-name'), 'My custom name');
  const legacy = fixture(4);
  delete legacy.context.pageContext.list;
  delete legacy.context.pageContext.listItem;
  legacy.context.spHttpClient.get = () => { throw new Error('Workbench does not have a persisted page item'); };
  await legacy.ensure();
  assert.strictEqual(legacy.properties.instanceName, prefix + '1', 'Workbench/legacy pages use live canvas reservations');
  const html = fixture(5);
  const wrapper = html.document.createElement('div');
  wrapper.setAttribute('data-sp-controldata', JSON.stringify(canvasControl(80, prefix.toUpperCase() + '1')));
  html.context.spHttpClient.get = async () => response({ CanvasContent1: wrapper.outerHTML });
  await html.ensure();
  assert.strictEqual(html.properties.instanceName, prefix + '2', 'Parses HTML canvas metadata and reserves names case-insensitively');
  const failed = fixture(6);
  failed.context.spHttpClient.get = async () => ({ ok: false, status: 403 });
  await assert.rejects(() => failed.ensure(), /403/);
  assert.strictEqual(failed.properties.instanceName, undefined, 'Discovery errors cannot produce a success-shaped default');
  failed.context.spHttpClient.get = async () => response({ CanvasContent1: '[]' });
  await failed.ensure();
  assert.strictEqual(failed.properties.instanceName, prefix + '1', 'Failed discovery can be retried');
  const delayed = fixture(7);
  let finish;
  delayed.context.spHttpClient.get = () => new Promise(resolve => { finish = resolve; });
  const allocation = delayed.ensure();
  delayed.properties.instanceName = 'Typed while loading';
  finish(response({ CanvasContent1: '[]' }));
  await allocation;
  assert.strictEqual(delayed.properties.instanceName, 'Typed while loading', 'Pending discovery cannot overwrite a user edit');
  const changingPage = fixture(8);
  changingPage.context.spHttpClient.get = () => new Promise(resolve => { finish = resolve; });
  const oldPage = changingPage.ensure();
  const newPage = fixture(9, changingPage.document);
  newPage.context.pageContext.listItem.id = 8;
  await newPage.ensure();
  finish(response({ CanvasContent1: '[]' }));
  assert.strictEqual(await oldPage, false, 'Navigation invalidates an old page allocation');
  assert.strictEqual(changingPage.properties.instanceName, undefined);
  const malformed = fixture(10);
  malformed.context.spHttpClient.get = async () => response({ CanvasContent1: '[broken' });
  await assert.rejects(() => malformed.ensure());
  assert.strictEqual(malformed.properties.instanceName, undefined);
  const cancelled = fixture(12);
  let mode = 2;
  cancelled.context.spHttpClient.get = () => new Promise(resolve => { finish = resolve; });
  const pendingEdit = helper.ensureInstanceName(cancelled.properties, prefix, cancelled.context, cancelled.element, () => mode);
  mode = 1;
  finish(response({ CanvasContent1: '[]' }));
  assert.strictEqual(await pendingEdit, false, 'Leaving edit mode invalidates pending allocation');
  assert.strictEqual(cancelled.properties.instanceName, undefined);

  function Base() {}
  Base.prototype.onPropertyPaneFieldChanged = function(property, oldValue, newValue) { this.properties[property] = newValue; };
  const field = kind => (targetProperty, properties) => ({ kind, targetProperty, properties });
  const exports = {};
  const file = path.join(root, 'src', 'webparts', folder, className + '.ts');
  vm.runInNewContext(ts.transpileModule(fs.readFileSync(file, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5 }
  }).outputText, {
    exports, window: first.document.defaultView, document: first.document, console, setTimeout, clearTimeout,
    require(module) {
      if (module === '../shared/instanceNames') { return helper; }
      if (module === '@microsoft/sp-core-library') { return { DisplayMode: { Read: 1, Edit: 2 } }; }
      if (module === '@microsoft/sp-webpart-base') {
        return {
          BaseClientSideWebPart: Base, PropertyPaneButtonType: { Normal: 0, Primary: 1 },
          PropertyPaneButton: field('button'), PropertyPaneCheckbox: field('checkbox'),
          PropertyPaneDropdown: field('dropdown'), PropertyPaneLabel: field('label'),
          PropertyPaneSlider: field('slider'), PropertyPaneTextField: field('text'), PropertyPaneToggle: field('toggle')
        };
      }
      if (module.endsWith('/PropertyPaneCustomField')) { return { PropertyPaneCustomField: properties => ({ kind: 'custom', properties }) }; }
      if (module.endsWith('/PropertyPaneCheckbox')) { return { PropertyPaneCheckbox: field('checkbox') }; }
      if (module.endsWith('/PropertyFieldColorPicker')) {
        return { PropertyFieldColorPicker: field('color'), PropertyFieldColorPickerStyle: { Inline: 0 } };
      }
      if (module.endsWith('/PropertyFieldCollectionData')) {
        return { PropertyFieldCollectionData: field('collection'), CustomCollectionFieldType: {} };
      }
      if (module === '../shared/overrideCss') { return { PropertyPaneOverrideCss: field('override') }; }
      if (module === './components/PageNavigation') { return { getPageNavigationTargets: () => [] }; }
      if (/WebPartStrings$/.test(module)) { return strings; }
      if (module === '../../../config/package-solution.json') { return require(path.join(root, 'config', 'package-solution.json')); }
      if (module === 'react' || module === 'react-dom' || /WebPartStrings$/.test(module)
        || /^(?:@microsoft\/|@pnp\/|\.\/components\/|\.\.\/shared\/)/.test(module)) { return {}; }
      throw new Error('Unexpected web part import: ' + module);
    }
  }, { filename: file });
  const part = new exports.default();
  const live = fixture(11);
  part.properties = live.properties;
  part.context = live.context;
  part.domElement = live.element;
  part.displayMode = 2;
  let initializedName;
  let publishedName;
  let renders = 0;
  part.initializeControl = async () => { initializedName = part.properties.instanceName; };
  part.notifyDynamicData = property => { if (property === 'instanceName') { publishedName = part.properties.instanceName; } };
  part.notifyDynamicSourceChanged = () => {};
  part.handleRuntimeStateChanged = () => {};
  part.promoteRuntimeDynamicProperties = () => {};
  part.ensureDynamicValueChangeHandler = () => {};
  part.render = () => { renders++; };
  await part.onInit();
  assert.strictEqual(initializedName, prefix + '1', 'Assigns the name before Dynamic Data initialization');
  assert.strictEqual(publishedName, prefix + '1');
  assert.strictEqual(part.metadata.title, prefix + '1', 'Dynamic Data displays the generated friendly name');
  assert.strictEqual(part.getPropertyValue('instanceName'), prefix + '1');
  assert.strictEqual(part.metadata.instanceId, id(11), 'Connection identity remains the original GUID');
  const configuration = part.getPropertyPaneConfiguration();
  const groups = configuration.pages[0].groups;
  const fields = groups.reduce((all, group) => all.concat(group.groupFields), []);
  const editableFields = fields.filter(item => item.kind !== 'label' && item.kind !== 'custom');
  assert.strictEqual(editableFields[0].targetProperty, 'instanceName', 'Web part name is the first editable setting');
  const nameFields = configuration.pages.reduce((all, page) =>
    all.concat(...page.groups.map(group => group.groupFields)), []).filter(item => item.targetProperty === 'instanceName');
  assert.strictEqual(nameFields.length, 1, 'Name field appears once, not again in the data/form settings');
  assert.strictEqual(nameFields[0].properties.label, 'Web part name', 'Uses the same label in all four controls');
  assert.strictEqual(nameFields[0].properties.value, prefix + '1', 'Property pane displays the automatically allocated name');
  const versionIndex = groups.findIndex(group => group.groupFields.some(item => item.targetProperty === 'propertyPaneVersionInfo'));
  const nameIndex = groups.findIndex(group => group.groupFields.includes(nameFields[0]));
  assert.strictEqual(nameIndex, versionIndex + 1, 'Name section follows the existing version banner');
  assert.strictEqual(groups[nameIndex].groupName, 'Web part');
  assert.strictEqual(part.getPropertyDefinitions().find(item => item.id === 'instanceName').title, 'Web part name');
  part.onPropertyPaneFieldChanged('instanceName', part.properties.instanceName, prefix + '25');
  await new Promise(resolve => setImmediate(resolve));
  assert.strictEqual(part.properties.instanceName, prefix + '25');
  assert.strictEqual(part.properties.instanceNameGeneratedFor, undefined, 'User edits clear generated-name provenance');
  part.onPropertyPaneFieldChanged('instanceName', prefix + '25', '');
  await new Promise(resolve => setImmediate(resolve));
  assert.strictEqual(part.properties.instanceName, prefix + '1', 'Clearing the field restores a default');
  part.properties = {};
  part.displayMode = 1;
  await part.onInit();
  assert.strictEqual(part.properties.instanceName, undefined);
  part.displayMode = 2;
  part.onDisplayModeChanged(1);
  await new Promise(resolve => setImmediate(resolve));
  assert.strictEqual(part.properties.instanceName, prefix + '1', 'Existing blanks are repaired when entering page edit mode');
  assert(renders > 0);
  if (name === 'ListControl' || name === 'GridControl') {
    const navigation = {};
    const source = fs.readFileSync(path.join(root, 'src', 'webparts', folder, 'components', 'PageNavigation.ts'), 'utf8');
    vm.runInNewContext(ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS } }).outputText, {
      exports: navigation, console, require: () => ({})
    });
    const target = navigation.getPageNavigationTargets(live.document, id(999)).find(item => item.instanceId === id(11));
    assert.strictEqual(target.title, prefix + '1', 'Navigation target pickers use the live generated name without waiting for a page save');
  }
  console.log(name + ': numbering, persistence, hidden/copy collisions, failures, lifecycle and property-pane label/order checks passed');
}

(async () => {
  const helper = fs.readFileSync(path.join(__dirname, '..', 'ListControl', 'src', 'webparts', 'shared', 'instanceNames.ts'), 'utf8');
  for (const project of projects) {
    assert.strictEqual(fs.readFileSync(path.join(__dirname, '..', project[0], 'src', 'webparts', 'shared', 'instanceNames.ts'), 'utf8'), helper);
    await run(project);
  }
  const document = require(path.join(__dirname, '..', 'ListControl', 'node_modules', 'jsdom')).jsdom('<body></body>');
  let queries = 0;
  const context = number => ({
    instanceId: id(number), pageContext: {
      web: { absoluteUrl: 'https://example.test' }, list: { id: { toString: () => id(90) } }, listItem: { id: 8 }
    }, spHttpClient: { get: async () => { queries++; return response({ CanvasContent1: '[]' }); } }
  });
  const element = () => { const host = document.createElement('div'); document.body.appendChild(host); return host; };
  await require('./loadInstanceNames')(path.join(__dirname, '..', 'ListControl')).ensureInstanceName(
    { instanceName: 'DynamicForm1' }, 'ListControl', context(100), element(), () => 2);
  const names = projects.map(() => ({}));
  await Promise.all(projects.map((project, index) => require('./loadInstanceNames')(path.join(__dirname, '..', project[0])).ensureInstanceName(
    names[index], project[3], context(101 + index), element(), () => 2)));
  assert.deepStrictEqual(names.map(item => item.instanceName), ['ListControl1', 'GridControl1', 'DynamicForm2', 'ReportForm1']);
  assert.strictEqual(queries, 1, 'Independent project helpers share page reservations and one canvas discovery');
  console.log('Cross-project name reservations and live navigation labels passed');
})().catch(error => { console.error(error.stack || error); process.exitCode = 1; });
