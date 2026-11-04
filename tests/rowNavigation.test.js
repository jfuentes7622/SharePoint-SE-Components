'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

let passed = 0;
let failed = 0;

['ListControl', 'GridControl'].forEach(function(project) {
  const root = path.join(__dirname, '..', project);
  const base = path.join(root, 'src', 'webparts', project === 'ListControl' ? 'listControl' : 'gridControl');
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const React = require(path.join(root, 'node_modules', 'react'));
  const jsdom = require(path.join(root, 'node_modules', 'jsdom'));
  const document = jsdom.jsdom('<!doctype html><html><body></body></html>');
  const window = document.defaultView;
  const sourceId = '11111111-1111-4111-8111-111111111111';
  const tabId = '22222222-2222-4222-8222-222222222222';
  const formId = '33333333-3333-4333-8333-333333333333';
  const missingId = '44444444-4444-4444-8444-444444444444';
  const errors = [];
  let strings;
  vm.runInNewContext(fs.readFileSync(path.join(base, 'loc', 'en-us.js'), 'utf8'), {
    define: function(dependencies, factory) { strings = factory(); }
  });
  function field(kind) { return function(name, properties) { return { kind: kind, name: name, properties: properties }; }; }
  function Base() {}
  Base.prototype.onPropertyPaneFieldChanged = function(name, oldValue, newValue) { this.properties[name] = newValue; };
  const modules = {};
  function load(relative) {
    if (modules[relative]) { return modules[relative]; }
    const exports = {};
    const compiled = ts.transpileModule(fs.readFileSync(path.join(base, relative), 'utf8'), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
    }).outputText;
    const imports = {
      react: React, 'react-dom': { render: function(element, host) { host.reactElement = element; } },
      '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } },
      '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } },
      '@microsoft/sp-webpart-base': {
        BaseClientSideWebPart: Base, PropertyPaneButtonType: { Normal: 0, Primary: 1 },
        PropertyPaneButton: field('button'), PropertyPaneDropdown: field('dropdown'),
        PropertyPaneCheckbox: field('checkbox'), PropertyPaneLabel: field('label'),
        PropertyPaneTextField: field('text'), PropertyPaneSlider: field('slider')
      },
      '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField': {
        PropertyPaneCustomField: function(properties) { return { properties: properties }; }
      },
      '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker': {
        PropertyFieldColorPicker: field('color'), PropertyFieldColorPickerStyle: { Inline: 0 }
      },
      './components/ListDesigner': {}, './components/GridDesigner': {},
      './DateTime24HourInput': {}, './GridRichTextEditor': {}, './GridValidation': {},
      '../shared/deterministicFullWidth': { updateResponsiveOptionalFullWidth: function() {} },
      '../shared/overrideCss': { applyOverrideCss: function() {}, PropertyPaneOverrideCss: field('override') },
      '../shared/instanceNames': require('./loadInstanceNames')(root),
      '../../../config/package-solution.json': require(path.join(root, 'config', 'package-solution.json'))
    };
    imports[project + 'WebPartStrings'] = strings;
    vm.runInNewContext(compiled, {
      exports: exports, document: document, window: window, setTimeout: setTimeout, clearTimeout: clearTimeout,
      console: { log: function() {}, error: function() { errors.push(Array.from(arguments)); } },
      require: function(name) {
        if (name === './LibraryBrowser') { return require('./loadLibraryBrowser')(root, window); }
        if (name === './components/PageNavigation') { return load(path.join('components', 'PageNavigation.ts')); }
        if (name === './components/' + project) { return load(path.join('components', project + '.tsx')); }
        if (/\.css$/.test(name)) { return {}; }
        assert.ok(Object.prototype.hasOwnProperty.call(imports, name), 'Unexpected import: ' + name);
        return imports[name];
      }
    }, { filename: relative });
    modules[relative] = exports;
    return exports;
  }
  const navigation = load(path.join('components', 'PageNavigation.ts'));
  const Control = load(path.join('components', project + '.tsx'))[project];
  const WebPart = load(project + 'WebPart.ts').default;
  function test(name, run) {
    try { run(); passed += 1; console.log('PASS ' + project + ': ' + name); }
    catch (error) { failed += 1; console.error('FAIL ' + project + ': ' + name + '\n' + error.stack); }
  }
  function elements(element, predicate, result) {
    result = result || [];
    if (!element || !element.props) { return result; }
    if (predicate(element)) { result.push(element); }
    React.Children.forEach(element.props.children, function(child) { elements(child, predicate, result); });
    return result;
  }
  function classElements(element, name) {
    return elements(element, function(candidate) { return String(candidate.props.className || '').split(' ').indexOf(name) >= 0; });
  }
  function rows(instance) { return classElements(instance.render(), 'lc-row'); }
  function buttons(instance) { return classElements(instance.render(), 'lc-navigation-button'); }
  function visibleButtons(instance) { return buttons(instance).filter(function(button) { return button.props.style.visibility === 'visible'; }); }
  function metadata(instanceId, title) {
    return JSON.stringify({ controlType: 3, id: instanceId, webPartData: {
      instanceId: instanceId, title: title, properties: { instanceName: title }
    } });
  }
  function fixture(targetId) {
    document.body.innerHTML = '<button id="stay">Stay here</button><div id="source"></div>'
      + '<div id="tab-host"><div id="tabs" data-spse-tab-type="1"><ul role="tablist">'
      + '<li role="tab" class="selected">General</li><li role="tab">Other</li></ul></div>'
      + '<section id="active-panel">Active</section><section id="hidden-panel" style="display:none">Hidden</section></div>'
      + '<div id="form-host"><div id="form"><h2>Details form</h2><input id="editable" /></div></div>';
    const source = document.getElementById('source');
    const tabs = document.getElementById('tabs');
    const form = document.getElementById('form');
    source.setAttribute('data-sp-webpart-instance-id', sourceId);
    document.getElementById('tab-host').setAttribute('data-sp-controldata', metadata(tabId, 'Record tabs'));
    document.getElementById('form-host').setAttribute('data-sp-controldata', metadata(formId, 'Details form'));
    form.setAttribute('data-spse-form-instance-id', formId);
    const scrolls = [];
    [tabs, tabs.querySelector('[role="tablist"]'), form].forEach(function(element) {
      element.scrollIntoView = function(options) { scrolls.push({ element: element, options: options }); };
    });
    const notifications = [];
    const part = new WebPart();
    part.domElement = source;
    part.context = { instanceId: sourceId, manifest: { id: sourceId },
      pageContext: { web: { absoluteUrl: 'https://example.test/site' } },
      propertyPane: { refresh: function() {} } };
    part.properties = { enableDiagnostics: false, listName: 'Items', navigationTargetInstanceId: targetId || '' };
    part._dynamicDataSourceManager = { notifyPropertyChanged: function(name) { notifications.push(name); } };
    part.render();
    const instance = new Control(Object.assign({
      gridSchemaJson: '', pageSize: 50, fetchBatchSize: 500, timeMinuteIncrement: 5,
      showDelete: false, showHistory: false, onFilteredCountChange: function() {}
    }, source.reactElement.props));
    instance.setState = function(update) {
      this.state = Object.assign({}, this.state, typeof update === 'function' ? update(this.state, this.props) : update);
    };
    Object.assign(instance.state, {
      loading: false, accessLoading: false, hasReadAccess: true, hasEditAccess: true,
      fields: [{ Name: 'Title', TypeAsString: 'Text', DisplayName: 'Title' }],
      rows: [{ ID: 1, Title: 'First' }, { ID: 2, Title: 'Second' }]
    });
    document.getElementById('stay').focus();
    return { instance: instance, part: part, source: source, tabs: tabs, form: form,
      scrolls: scrolls, notifications: notifications };
  }

  test('navigation defaults off without adding a column or changing row selection', function() {
    const page = fixture();
    assert.strictEqual(buttons(page.instance).length, 0);
    assert.strictEqual(classElements(page.instance.render(), 'lc-navigation-header').length, 0);
    assert.strictEqual(rows(page.instance)[0].props.tabIndex, undefined);
    rows(page.instance)[0].props.onClick();
    assert.strictEqual(page.part.getPropertyValue('selectedItemId'), 1);
    assert.strictEqual(page.part.getPropertyValue('selectedMode'), 'view');
    assert.strictEqual(page.scrolls.length, 0);
  });

  test('selecting or repeatedly clicking a row only highlights and publishes selection', function() {
    const page = fixture(tabId);
    assert.strictEqual(visibleButtons(page.instance).length, 0);
    rows(page.instance)[0].props.onClick();
    rows(page.instance)[0].props.onClick();
    assert.strictEqual(page.instance.state.selectedItemId, 1);
    assert.strictEqual(classElements(page.instance.render(), 'lc-row-selected').length, 1);
    assert.strictEqual(visibleButtons(page.instance).length, 1);
    assert.strictEqual(page.part.getPropertyValue('selectedItemId'), 1);
    assert.strictEqual(page.scrolls.length, 0);
    assert.strictEqual(document.activeElement.id, 'stay');
  });

  test('the explicit action focuses the tab strip without switching tabs or republishing selection', function() {
    const page = fixture(tabId);
    rows(page.instance)[0].props.onClick();
    const count = page.notifications.length;
    let stopped = 0;
    visibleButtons(page.instance)[0].props.onClick({ stopPropagation: function() { stopped += 1; } });
    assert.strictEqual(stopped, 1);
    assert.strictEqual(document.activeElement, page.tabs.querySelector('[role="tablist"]'));
    assert.strictEqual(page.scrolls.length, 1);
    assert.strictEqual(page.scrolls[0].options.block, 'center');
    assert.strictEqual(page.notifications.length, count);
    assert.strictEqual(page.part.getPropertyValue('selectedMode'), 'view');
    assert.strictEqual(page.tabs.querySelector('.selected').textContent, 'General');
    assert.strictEqual(document.getElementById('hidden-panel').style.display, 'none');
    assert.strictEqual(page.instance.state.selectedItemId, 1);
  });

  test('direct form navigation focuses its container, never the first editable field', function() {
    const page = fixture(formId);
    rows(page.instance)[0].props.onClick();
    visibleButtons(page.instance)[0].props.onClick({ stopPropagation: function() {} });
    assert.strictEqual(document.activeElement, page.form);
    assert.notStrictEqual(document.activeElement.id, 'editable');
    assert.strictEqual(page.form.getAttribute('tabindex'), '-1');
    document.getElementById('stay').focus();
    assert.strictEqual(page.form.hasAttribute('tabindex'), false);
  });

  test('existing destination tabindex is preserved', function() {
    const page = fixture(formId);
    page.form.setAttribute('tabindex', '0');
    navigation.navigateToPageTarget(document, formId, sourceId);
    document.getElementById('stay').focus();
    assert.strictEqual(page.form.getAttribute('tabindex'), '0');
  });

  test('choosing another row moves the visible action without changing table structure', function() {
    const page = fixture(tabId);
    const cells = rows(page.instance).map(function(row) { return elements(row, function(element) { return element.type === 'td'; }).length; });
    rows(page.instance)[0].props.onClick();
    assert.strictEqual(buttons(page.instance)[0].props.style.visibility, 'visible');
    rows(page.instance)[1].props.onClick();
    assert.strictEqual(buttons(page.instance)[0].props.style.visibility, 'hidden');
    assert.strictEqual(buttons(page.instance)[1].props.style.visibility, 'visible');
    assert.strictEqual(page.instance.state.selectedItemId, 2);
    assert.deepStrictEqual(rows(page.instance).map(function(row) {
      return elements(row, function(element) { return element.type === 'td'; }).length;
    }), cells);
  });

  ['Enter', ' ', 'Spacebar'].forEach(function(key) {
    test('keyboard ' + JSON.stringify(key) + ' selects without navigating, while child events are ignored', function() {
      const page = fixture(tabId);
      const row = rows(page.instance)[0];
      const current = {};
      let prevented = 0;
      assert.strictEqual(row.props.tabIndex, 0);
      row.props.onKeyDown({ key: key, target: {}, currentTarget: current, preventDefault: function() { prevented += 1; } });
      assert.strictEqual(page.instance.state.selectedItemId, 0);
      row.props.onKeyDown({ key: key, target: current, currentTarget: current, preventDefault: function() { prevented += 1; } });
      assert.strictEqual(prevented, 1);
      assert.strictEqual(page.instance.state.selectedItemId, 1);
      assert.strictEqual(page.scrolls.length, 0);
      assert.strictEqual(document.activeElement.id, 'stay');
    });
  });

  test('a missing target shows an error on the selected row and logs the failure', function() {
    const page = fixture(missingId);
    const before = errors.length;
    rows(page.instance)[0].props.onClick();
    visibleButtons(page.instance)[0].props.onClick({ stopPropagation: function() {} });
    assert.strictEqual(page.instance.state.navigationError, strings.RuntimeNavigationMissing);
    assert.strictEqual(classElements(page.instance.render(), 'lc-navigation-error')[0].props.role, 'alert');
    assert.strictEqual(errors.length, before + 1);
    assert.strictEqual(page.scrolls.length, 0);
    assert.strictEqual(document.activeElement.id, 'stay');
    rows(page.instance)[1].props.onClick();
    assert.strictEqual(page.instance.state.navigationError, '');
  });

  ['display', 'visibility', 'hidden'].forEach(function(property) {
    test('a target hidden by ancestor ' + property + ' is not focused or silently activated', function() {
      const page = fixture(formId);
      const host = document.getElementById('form-host');
      if (property === 'hidden') { host.hidden = true; }
      else { host.style[property] = property === 'display' ? 'none' : 'hidden'; }
      rows(page.instance)[0].props.onClick();
      visibleButtons(page.instance)[0].props.onClick({ stopPropagation: function() {} });
      assert.strictEqual(page.instance.state.navigationError, strings.RuntimeNavigationHidden);
      assert.strictEqual(page.scrolls.length, 0);
      assert.strictEqual(document.activeElement.id, 'stay');
    });
  });

  test('target discovery includes Tabs without Dynamic Data, forms and unique instance IDs, but excludes self', function() {
    const page = fixture(tabId);
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    assert.strictEqual(targets.length, 2);
    const tabs = targets.filter(function(target) { return target.instanceId === tabId; })[0];
    assert.strictEqual(tabs.title, 'Record tabs');
    assert.strictEqual(tabs.element, page.tabs);
    assert.strictEqual(targets.filter(function(target) { return target.instanceId === formId; })[0].element, page.form);
    assert.strictEqual(targets.filter(function(target) { return target.instanceId === sourceId; }).length, 0);
  });

  test('two controls with the same title remain separately selectable by instance ID', function() {
    const page = fixture(tabId);
    document.getElementById('form-host').setAttribute('data-sp-controldata', metadata(formId, 'Record tabs'));
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    assert.strictEqual(targets.length, 2);
    assert.strictEqual(targets[0].title, targets[1].title);
    const options = page.part.getNavigationTargetOptions();
    assert.ok(options.some(function(option) { return option.key === tabId && option.text.indexOf(tabId) >= 0; }));
    assert.ok(options.some(function(option) { return option.key === formId && option.text.indexOf(formId) >= 0; }));
  });

  test('page-control identity works with feature-instance attributes and legacy web-part IDs', function() {
    fixture();
    document.body.innerHTML = '<div data-sp-feature-instance-id="' + tabId + '"></div>'
      + '<div id="WebPart.' + formId + '"><div data-spse-tab-type="2"></div></div>'
      + '<div id="div_' + missingId + '"></div>';
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    assert.strictEqual(targets.length, 3);
    assert.deepStrictEqual(Array.from(targets, function(target) { return target.instanceId; }).sort(), [tabId, formId, missingId].sort());
  });

  test('malformed canvas metadata is logged without preventing other targets from being listed', function() {
    fixture();
    const bad = document.createElement('div');
    bad.setAttribute('data-sp-controldata', '{bad json');
    document.body.appendChild(bad);
    const before = errors.length;
    assert.strictEqual(navigation.getPageNavigationTargets(document, sourceId).length, 2);
    assert.ok(errors.length > before);
  });

  test('the property pane exposes None, refresh and an unavailable saved target without silently replacing it', function() {
    const page = fixture(missingId);
    const configuration = page.part.getPropertyPaneConfiguration();
    const groups = configuration.pages.reduce(function(all, item) { return all.concat(item.groups || []); }, []);
    const group = groups.filter(function(item) { return item.groupName === strings.PropertyGroupNavigation; })[0];
    const dropdown = group.groupFields.filter(function(item) { return item.name === 'navigationTargetInstanceId'; })[0];
    assert.strictEqual(dropdown.properties.selectedKey, missingId);
    assert.strictEqual(dropdown.properties.options[0].key, '');
    assert.ok(dropdown.properties.options.some(function(option) { return option.key === missingId; }));
    assert.ok(group.groupFields.some(function(item) { return item.name === 'refreshNavigationTargets'; }));
    assert.strictEqual(page.part.properties.navigationTargetInstanceId, missingId);
  });

  test('group headers span the reserved navigation column as well as data/action columns', function() {
    const page = fixture(tabId);
    const grouping = { enabled: true, field1: 'Title', showCount: true };
    if (project === 'GridControl') { page.instance.props.gridSchemaJson = JSON.stringify({ grouping: grouping }); }
    else { page.instance.props.groupingJson = JSON.stringify(grouping); }
    const tree = page.instance.render();
    const headers = elements(tree, function(element) { return element.type === 'th'; });
    const spans = elements(tree, function(element) { return element.type === 'td' && !!element.props.colSpan; });
    assert.ok(spans.length > 0);
    spans.forEach(function(cell) { assert.strictEqual(cell.props.colSpan, headers.length); });
  });

  ['text', 'iconText', 'icon'].forEach(function(mode) {
    test('navigation honors existing ' + mode + ' command-button styling and accessible labels', function() {
      const page = fixture(tabId);
      page.instance.props.buttonDisplayMode = mode;
      rows(page.instance)[0].props.onClick();
      const button = visibleButtons(page.instance)[0];
      assert.strictEqual(button.props['aria-label'], strings.RuntimeGoToDetails);
      assert.ok(button.props.className.indexOf('gc-command-button') >= 0);
      assert.strictEqual(button.props.className.indexOf('gc-command-button-icon-only') >= 0, mode === 'icon');
    });
  });

  if (project === 'GridControl') {
    test('bulk-delete checkboxes remain separate from the active item linked to Dynamic Forms', function() {
      const page = fixture(tabId);
      rows(page.instance)[0].props.onClick();
      page.instance.toggleItemChecked(2);
      assert.strictEqual(page.part.getPropertyValue('selectedItemId'), 1);
      assert.strictEqual(page.instance.state.selectedItemId, 1);
      assert.strictEqual(page.instance.state.selectedItemIds[0], 2);
      assert.strictEqual(page.scrolls.length, 0);
    });

    test('navigation is disabled during inline editing and creation, preserving editing values', function() {
      [0, 2].forEach(function(editingId) {
        const page = fixture(tabId);
        rows(page.instance)[0].props.onClick();
        page.instance.state.editingItemId = editingId;
        page.instance.state.editingValues = { Title: 'Unsaved value' };
        const button = visibleButtons(page.instance)[0];
        assert.strictEqual(button.props.disabled, true);
        assert.strictEqual(button.props.title, strings.RuntimeNavigationFinishEditing);
        button.props.onClick({ stopPropagation: function() {} });
        assert.strictEqual(page.scrolls.length, 0);
        assert.strictEqual(page.instance.state.editingValues.Title, 'Unsaved value');
        const headers = elements(page.instance.render(), function(element) { return element.type === 'th'; });
        const editing = classElements(page.instance.render(), 'gc-row-editing')[0];
        assert.strictEqual(elements(editing, function(element) { return element.type === 'td'; }).length, headers.length);
      });
    });

    test('read-only grids retain explicit navigation without requiring edit permissions', function() {
      const page = fixture(tabId);
      page.instance.state.hasEditAccess = false;
      rows(page.instance)[0].props.onClick();
      assert.strictEqual(visibleButtons(page.instance)[0].props.disabled, false);
      visibleButtons(page.instance)[0].props.onClick({ stopPropagation: function() {} });
      assert.strictEqual(page.scrolls.length, 1);
    });
  }
});

console.log(passed + ' passed, ' + failed + ' failed');
process.exitCode = failed ? 1 : 0;
