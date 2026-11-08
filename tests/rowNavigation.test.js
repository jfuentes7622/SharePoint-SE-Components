'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

let passed = 0;
let failed = 0;

function loadRegistrations(document) {
  const repositoryRoot = path.join(__dirname, '..');
  const registrations = [];
  function findWebParts(directory) {
    return fs.readdirSync(directory).reduce(function(files, name) {
      const file = path.join(directory, name);
      return files.concat(fs.statSync(file).isDirectory() ? findWebParts(file)
        : /WebPart\.ts$/.test(name) ? [file] : []);
    }, []);
  }
  fs.readdirSync(repositoryRoot).forEach(function(project) {
    const projectRoot = path.join(repositoryRoot, project);
    const sourceRoot = path.join(projectRoot, 'src', 'webparts');
    if (!fs.existsSync(path.join(projectRoot, 'config', 'package-solution.json')) || !fs.existsSync(sourceRoot)) { return; }
    const ts = require(path.join(projectRoot, 'node_modules', 'typescript'));
    findWebParts(sourceRoot).forEach(function(file) {
      const source = ts.createSourceFile(file, fs.readFileSync(file, 'utf8'), ts.ScriptTarget.ES5, true);
      const declaration = source.statements.filter(function(statement) { return statement.kind === ts.SyntaxKind.ClassDeclaration; })[0];
      const methodName = project === 'ListControl' ? 'renderContent' : 'render';
      const method = declaration.members.filter(function(member) { return member.name && member.name.text === methodName; })[0];
      assert.ok(method && method.body, project + ' has a rendering entry point');
      if (project === 'ListControl') {
        const render = declaration.members.filter(function(member) { return member.name && member.name.text === 'render'; })[0];
        assert.ok(render.getText(source).indexOf('this.renderContent()') >= 0);
      }
      const registration = method.body.statements.slice(0, 2).map(function(statement) { return statement.getText(source); }).join('\n');
      assert.ok(registration.indexOf('data-spse-instance-id') >= 0, project + ' registers an instance ID before rendering content');
      assert.ok(registration.indexOf('data-spse-control-title') >= 0, project + ' registers its framework-provided title');
      const manifestFile = file.replace(/\.ts$/, '.manifest.json');
      const parsed = ts.parseConfigFileTextToJson(manifestFile, fs.readFileSync(manifestFile, 'utf8'));
      assert.strictEqual(parsed.error, undefined);
      const manifest = parsed.config;
      const instanceId = ('00000000' + (registrations.length + 1).toString(16)).slice(-8) + '-abcd-4abc-8abc-123456789012';
      const element = document.createElement('div');
      const exports = {};
      vm.runInNewContext(ts.transpileModule('exports.register = function() {' + registration + '};', {
        compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5 }
      }).outputText, { exports: exports }, { filename: file });
      exports.register.call({
        domElement: element, context: { instanceId: instanceId },
        title: manifest.preconfiguredEntries[0].title.default
      });
      assert.strictEqual(element.getAttribute('data-spse-instance-id'), instanceId);
      assert.strictEqual(element.getAttribute('data-spse-control-title'), manifest.preconfiguredEntries[0].title.default);
      assert.ok(element.getAttribute('data-spse-control-title'));
      registrations.push({ project: project, instanceId: instanceId, element: element, title: manifest.preconfiguredEntries[0].title.default });
    });
  });
  return registrations;
}

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

  test('all 18 repository page web parts register discoverable targets without SharePoint wrapper metadata', function() {
    document.body.innerHTML = '';
    const registrations = loadRegistrations(document);
    assert.strictEqual(registrations.length, 18);
    registrations.forEach(function(registration) { document.body.appendChild(registration.element); });
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    assert.strictEqual(targets.length, registrations.length);
    registrations.forEach(function(registration) {
      const target = targets.filter(function(candidate) { return candidate.instanceId === registration.instanceId; })[0];
      assert.ok(target, registration.project + ' is a navigation target');
      assert.strictEqual(target.element, registration.element);
      assert.strictEqual(target.title, registration.title);
    });
  });

  test('metadata-free Tab is selectable and receives focus without changing the active tab', function() {
    const page = fixture(tabId);
    document.getElementById('tab-host').removeAttribute('data-sp-controldata');
    page.tabs.setAttribute('data-spse-instance-id', tabId);
    page.tabs.setAttribute('data-spse-control-title', 'SPS Tabs');
    const options = page.part.getNavigationTargetOptions();
    assert.ok(options.some(function(option) { return option.key === tabId && option.text.indexOf('SPS Tabs') >= 0; }));
    rows(page.instance)[0].props.onClick();
    visibleButtons(page.instance)[0].props.onClick({ stopPropagation: function() {} });
    assert.strictEqual(document.activeElement, page.tabs.querySelector('[role="tablist"]'));
    assert.strictEqual(page.tabs.querySelector('[role="tab"].selected').textContent, 'General');
    assert.strictEqual(page.scrolls.length, 1);
    assert.strictEqual(page.instance.state.navigationError, '');
  });

  test('manifest fallback titles preserve custom names and deduplicate canvas wrappers', function() {
    fixture(tabId);
    const form = document.getElementById('form');
    form.setAttribute('data-spse-instance-id', formId);
    form.setAttribute('data-spse-control-title', 'SPS Report Forms');
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    assert.strictEqual(targets.length, 2);
    assert.strictEqual(targets.filter(function(target) { return target.instanceId === formId; })[0].title, 'Details form');
    form.setAttribute('data-spse-instance-name', 'Custom report');
    assert.strictEqual(navigation.getPageNavigationTargets(document, sourceId)
      .filter(function(target) { return target.instanceId === formId; })[0].title, 'Custom report');
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

  test('built-in and third-party web-part metadata uses page instance IDs, not shared component IDs', function() {
    const page = fixture();
    const componentId = '55555555-5555-4555-8555-555555555555';
    const secondId = '66666666-6666-4666-8666-666666666666';
    const vendorId = '77777777-7777-4777-8777-777777777777';
    [missingId, secondId, vendorId].forEach(function(id, index) {
      const wrapper = document.createElement('section');
      wrapper.setAttribute('data-sp-canvascontrol', '');
      const host = document.createElement('div');
      host.setAttribute('data-sp-webpart', '');
      host.setAttribute('data-sp-webpartdata', JSON.stringify({
        id: componentId, instanceId: id, title: index < 2 ? 'Image' : 'Vendor dashboard'
      }));
      wrapper.appendChild(host);
      document.body.appendChild(wrapper);
    });
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    assert.strictEqual(targets.length, 5);
    assert.strictEqual(targets.filter(function(target) { return target.title === 'Image'; }).length, 2);
    assert.ok(!targets.some(function(target) { return target.instanceId === componentId; }));
    const options = page.part.getNavigationTargetOptions();
    [missingId, secondId, vendorId].forEach(function(id) {
      assert.ok(options.some(function(option) { return option.key === id && option.text.indexOf(id) >= 0; }));
    });
    const unidentified = document.createElement('div');
    unidentified.setAttribute('data-sp-webpartdata', JSON.stringify({ id: componentId, title: 'No instance ID' }));
    document.body.appendChild(unidentified);
    assert.strictEqual(navigation.getPageNavigationTargets(document, sourceId).length, 5,
      'Never use the reusable component GUID as a page-instance navigation key.');
  });

  test('Text canvas controls appear in the dropdown and receive focus through Go to details', function() {
    const page = fixture(missingId);
    const text = document.createElement('section');
    text.setAttribute('data-sp-canvascontrol', '');
    text.setAttribute('data-sp-controldata', JSON.stringify({ id: missingId, controlType: 4 }));
    text.innerHTML = '<h2>Customer instructions</h2><p>Read these instructions before entering a record.</p>';
    let scrolls = 0;
    text.scrollIntoView = function() { scrolls++; };
    document.body.appendChild(text);
    const option = page.part.getNavigationTargetOptions().filter(function(candidate) { return candidate.key === missingId; })[0];
    assert.ok(option && option.text.indexOf('Customer instructions') >= 0);
    rows(page.instance)[0].props.onClick();
    visibleButtons(page.instance)[0].props.onClick({ stopPropagation: function() {} });
    assert.strictEqual(document.activeElement, text);
    assert.strictEqual(scrolls, 1);
    assert.strictEqual(text.getAttribute('tabindex'), '-1');
    assert.strictEqual(page.part.getPropertyValue('selectedItemId'), 1);
    assert.strictEqual(page.part.properties.navigationTargetInstanceId, missingId);
    document.getElementById('stay').focus();
    assert.strictEqual(text.hasAttribute('tabindex'), false);
  });

  test('Text labels use bounded rendered previews, with a localized fallback for empty blocks', function() {
    fixture();
    const text = document.createElement('div');
    text.setAttribute('data-sp-controldata', JSON.stringify({ id: missingId, controlType: 4 }));
    text.textContent = 'Customer instructions '.repeat(20);
    document.body.appendChild(text);
    let target = navigation.getPageNavigationTargets(document, sourceId).filter(function(candidate) {
      return candidate.instanceId === missingId;
    })[0];
    assert.strictEqual(target.title.length, 80);
    assert.ok(target.title.indexOf('Customer instructions') === 0);
    text.textContent = '';
    target = navigation.getPageNavigationTargets(document, sourceId).filter(function(candidate) {
      return candidate.instanceId === missingId;
    })[0];
    assert.strictEqual(target.title, strings.PropNavigationText);
  });

  test('standard host labels identify native controls without repository markers or Dynamic Data', function() {
    fixture();
    const identities = [
      [missingId, 'data-sp-feature-instance-id', 'data-sp-feature-tag', 'Quick links'],
      ['55555555-5555-4555-8555-555555555555', 'data-sp-webpart-instance-id', 'aria-label', 'Customer contacts']
    ];
    identities.forEach(function(entry) {
      const host = document.createElement('div');
      host.setAttribute(entry[1], entry[0]);
      host.setAttribute(entry[2], entry[3]);
      document.body.appendChild(host);
    });
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    identities.forEach(function(entry) {
      assert.strictEqual(targets.filter(function(target) { return target.instanceId === entry[0]; })[0].title, entry[3]);
    });
  });

  test('legacy underscore IDs and classic WebPartID hosts are discoverable and focusable', function() {
    fixture();
    const identities = [missingId, '55555555-5555-4555-8555-555555555555', '66666666-6666-4666-8666-666666666666'];
    const elements = identities.map(function(id, index) {
      const host = document.createElement('div');
      host.id = index === 0 ? 'WebPart_' + id : index === 1 ? 'webpart_' + id : 'WebPartWPQ2';
      if (index === 2) { host.setAttribute('WebPartID', '{' + id.toUpperCase() + '}'); }
      host.setAttribute('aria-label', 'Native control ' + index);
      host.scrollIntoView = function() {};
      document.body.appendChild(host);
      return host;
    });
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    identities.forEach(function(id, index) {
      const target = targets.filter(function(candidate) { return candidate.instanceId === id; })[0];
      assert.strictEqual(target.title, 'Native control ' + index);
      navigation.navigateToPageTarget(document, id, sourceId);
      assert.strictEqual(document.activeElement, elements[index]);
    });
  });

  test('a canvas control and its nested native host appear once, preserving the existing canvas target key', function() {
    const page = fixture();
    const wrapper = document.createElement('section');
    wrapper.setAttribute('data-sp-canvascontrol', '');
    wrapper.setAttribute('data-sp-controldata', JSON.stringify({ id: missingId, controlType: 3 }));
    wrapper.innerHTML = '<h2>Native news</h2><div data-sp-webpart="" data-sp-webpartdata=\''
      + JSON.stringify({ instanceId: '55555555-5555-4555-8555-555555555555', title: 'News' }) + '\'></div>';
    wrapper.scrollIntoView = function() {};
    document.body.appendChild(wrapper);
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    assert.strictEqual(targets.length, 3);
    assert.strictEqual(targets.filter(function(target) { return target.instanceId === missingId; })[0].title, 'News');
    navigation.navigateToPageTarget(document, missingId, sourceId);
    assert.strictEqual(document.activeElement, wrapper);
    const alias = '55555555-5555-4555-8555-555555555555';
    navigation.navigateToPageTarget(document, alias, sourceId);
    assert.strictEqual(document.activeElement, wrapper, 'Previously saved native-host IDs remain navigable after deduplication.');
    page.part.properties.navigationTargetInstanceId = alias;
    const options = page.part.getNavigationTargetOptions();
    const selected = options.filter(function(option) { return option.key === alias; })[0];
    assert.ok(selected && selected.text.indexOf('News') >= 0);
    assert.ok(selected.text.indexOf(strings.PropNavigationUnavailable) < 0);
    assert.strictEqual(options.filter(function(option) { return option.key === missingId || option.key === alias; }).length, 1);
    assert.strictEqual(page.part.properties.navigationTargetInstanceId, alias, 'Do not rewrite a saved target to its wrapper ID.');
  });

  test('source exclusion covers the complete canvas control even if the wrapper has a different canvas ID', function() {
    const page = fixture();
    const wrapper = document.createElement('div');
    wrapper.setAttribute('data-sp-canvascontrol', '');
    wrapper.setAttribute('data-sp-controldata', JSON.stringify({ id: missingId, controlType: 3 }));
    document.body.appendChild(wrapper);
    wrapper.appendChild(page.source);
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    assert.strictEqual(targets.length, 2);
    assert.ok(!targets.some(function(target) { return target.instanceId === missingId || target.instanceId === sourceId; }));
  });

  test('invalid metadata is logged and a valid host identity is still usable', function() {
    fixture();
    const host = document.createElement('div');
    host.setAttribute('data-spse-instance-id', 'not-an-instance-id');
    host.setAttribute('data-sp-webpart-instance-id', missingId);
    host.setAttribute('data-sp-webpartdata', 'null');
    host.setAttribute('aria-label', 'Native control');
    document.body.appendChild(host);
    const before = errors.length;
    const targets = navigation.getPageNavigationTargets(document, sourceId);
    assert.strictEqual(targets.filter(function(target) { return target.instanceId === missingId; })[0].title, 'Native control');
    assert.ok(errors.length > before);
  });

  test('layout sections and ordinary HTML buttons are not fabricated into navigation targets', function() {
    fixture();
    const section = document.createElement('section');
    section.className = 'CanvasSection';
    section.innerHTML = '<h2>Page layout</h2><button id="ordinary-button">Not a page control</button>';
    document.body.appendChild(section);
    assert.strictEqual(navigation.getPageNavigationTargets(document, sourceId).length, 2);
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
