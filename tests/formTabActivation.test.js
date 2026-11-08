'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const test = require('node:test');

const root = path.join(__dirname, '..');
const activationEvent = 'spse:form-activation-changed';

function createHost() {
  const timers = new Map();
  const listeners = new Map();
  const observers = [];
  const errors = [];
  let nextTimer = 0;
  const window = {
    setTimeout: function(callback) { timers.set(++nextTimer, callback); return nextTimer; },
    clearTimeout: function(id) { timers.delete(id); },
    addEventListener: function(name, callback) {
      if (!listeners.has(name)) { listeners.set(name, []); }
      listeners.get(name).push(callback);
    },
    removeEventListener: function(name, callback) {
      listeners.set(name, (listeners.get(name) || []).filter(function(item) { return item !== callback; }));
    },
    dispatchEvent: function(event) {
      (listeners.get(event.type) || []).slice().forEach(function(callback) { callback(event); });
    },
    MutationObserver: function(callback) {
      this.records = [];
      this.observe = function(target, options) { this.target = target; this.options = options; };
      this.disconnect = function() { this.target = undefined; };
      this.callback = callback;
      observers.push(this);
    }
  };
  function createElement() {
    const attributes = {};
    const classes = new Set();
    const element = {
      parentElement: undefined,
      ownerDocument: document,
      children: [],
      writes: {},
      style: {
        setProperty: function(name, value) { this[name] = value; },
        removeProperty: function(name) { delete this[name]; }
      },
      classList: {
        add: function(name) { classes.add(name); },
        remove: function(name) { classes.delete(name); },
        contains: function(name) { return classes.has(name); }
      },
      getAttribute: function(name) {
        return Object.prototype.hasOwnProperty.call(attributes, name) ? attributes[name] : null;
      },
      setAttribute: function(name, value) {
        const oldValue = this.getAttribute(name);
        attributes[name] = String(value);
        this.writes[name] = (this.writes[name] || 0) + 1;
        observers.forEach(function(observer) {
          if (!observer.target || !observer.target.contains(element)
            || observer.options.attributeFilter.indexOf(name) < 0) { return; }
          observer.records.push({
            target: element,
            oldValue: observer.options.attributeOldValue ? oldValue : null,
            attributeName: name
          });
          if (observer.timer) { return; }
          observer.timer = window.setTimeout(function() {
            observer.timer = 0;
            const records = observer.records.splice(0);
            observer.callback(records);
          });
        });
      },
      removeAttribute: function(name) { delete attributes[name]; },
      appendChild: function(child) { child.parentElement = this; this.children.push(child); },
      contains: function(child) {
        while (child) {
          if (child === this) { return true; }
          child = child.parentElement;
        }
        return false;
      },
      querySelector: function(selector) {
        if (selector === '[data-spse-exclude-from-tabs="true"]') {
          return this.children.find(function(child) {
            return child.getAttribute('data-spse-exclude-from-tabs') === 'true'
              || child.querySelector(selector);
          }) || null;
        }
        return null;
      },
      querySelectorAll: function(selector) {
        return selector === '.CanvasZone' || selector === '.ControlZone' ? (this.zones || []) : [];
      }
    };
    Object.defineProperty(element, 'innerHTML', {
      set: function() { this.children = []; }
    });
    return element;
  }
  const document = {
    readyState: 'loading',
    createElement: createElement,
    querySelector: function() { return { parentElement: host.canvas }; }
  };
  const host = {
    window: window,
    document: document,
    createElement: createElement,
    errors: errors,
    renderTabCount: 0,
    reactDom: {
      render: function(element, container) {
        if (container.isTab) {
          host.renderTabCount += 1;
          let control = container.control;
          if (!control) {
            control = container.control = new element.type(element.props);
            control.render();
            control.componentDidMount();
          } else {
            const previousProps = control.props;
            control.props = element.props;
            control.render();
            control.componentDidUpdate(previousProps, control.state);
          }
        } else {
          container.reactElement = element;
          if (!container.formInstance) {
            container.formInstance = { values: {} };
            container.mounts = (container.mounts || 0) + 1;
          }
        }
      },
      unmountComponentAtNode: function(container) {
        container.formInstance = undefined;
        container.unmounts = (container.unmounts || 0) + 1;
      }
    },
    selectTab: function(index) {
      const control = host.tab.domElement.control;
      const previousState = control.state;
      control.state = { SelectedTab: index };
      control.render();
      control.componentDidUpdate(control.props, previousState);
    },
    settle: async function() {
      for (let pass = 0; pass < 20; pass += 1) {
        let count = 0;
        while (timers.size) {
          assert.ok(++count < 50, 'Tab/form render feedback did not settle');
          const entry = timers.entries().next().value;
          timers.delete(entry[0]);
          entry[1]();
        }
        await Promise.resolve();
      }
    }
  };
  return host;
}

function loadClass(host, project, relativeFile, extraImports) {
  const projectRoot = path.join(root, project);
  const ts = require(path.join(projectRoot, 'node_modules', 'typescript'));
  const source = fs.readFileSync(path.join(projectRoot, 'src', 'webparts', relativeFile), 'utf8');
  const output = ts.transpileModule(source, {
    fileName: relativeFile,
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
  }).outputText;
  const imports = Object.assign({
    '@microsoft/sp-webpart-base': { BaseClientSideWebPart: function() {} },
    '@microsoft/sp-core-library': {
      DisplayMode: { Read: 1, Edit: 2 },
      Environment: { type: 0 },
      EnvironmentType: { ClassicSharePoint: 1 }
    },
    '@microsoft/sp-http': {},
    '@microsoft/sp-component-base/lib/DynamicProperty': {},
    '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField': {},
    '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCheckbox/PropertyPaneCheckbox': {},
    '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker': {},
    '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData': {},
    'react': require(path.join(projectRoot, 'node_modules', 'react')),
    'react-dom': host.reactDom,
    'SharePointDynamicFormWebPartStrings': {},
    'TabComponentWebPartStrings': {},
    './components/SharePointDynamicForm': { SharePointDynamicFormContainer: function() {} },
    './components/RepeatedReportForms': { RepeatedReportForms: function() {} },
    './components/ErrorComponent': {},
    '../shared/deterministicFullWidth': { updateResponsiveOptionalFullWidth: function() {} },
    '../shared/overrideCss': { applyOverrideCss: function() {} },
    '../shared/instanceNames': {},
    '../../../config/package-solution.json': require(path.join(projectRoot, 'config', 'package-solution.json')),
    'sp-client-custom-fields/lib/PropertyFieldPicturePickerHost': {},
    '../tabComponent/assets/TabStyles-round.css': {}
  }, extraImports);
  const exports = {};
  vm.runInNewContext(output, {
    exports: exports,
    window: host.window,
    document: host.document,
    CustomEvent: function(name, options) { this.type = name; this.detail = options.detail; },
    console: {
      log: function() {},
      warn: function() {},
      error: function() { host.errors.push(Array.from(arguments)); }
    },
    require: function(name) {
      assert.ok(Object.prototype.hasOwnProperty.call(imports, name), 'Unexpected import: ' + name);
      return imports[name];
    }
  }, { filename: relativeFile });
  return exports.default;
}

function createScenario(project, tabType) {
  const host = createHost();
  const TabControl = loadClass(host, 'Tab', 'tabComponent\\components\\TabComponent.tsx', {
    './TabComponent.module.scss': { default: {} },
    './Tab': { default: function() {} }
  });
  const TabPart = loadClass(host, 'Tab', 'tabComponent\\TabComponentWebPart.ts', {
    './components/TabComponent': { default: TabControl }
  });
  const FormPart = loadClass(host, project, 'sharePointDynamicForm\\SharePointDynamicFormWebPart.ts');
  host.canvas = host.createElement();
  host.canvas.classList.add('CanvasSection');
  host.canvas.zones = [host.createElement(), host.createElement()];
  host.canvas.zones.forEach(function(zone) {
    zone.classList.add(tabType === '1' ? 'CanvasZone' : 'ControlZone');
    host.canvas.appendChild(zone);
  });
  host.tab = new TabPart();
  host.tab.properties = { TabType: tabType, enableDiagnostics: false };
  host.tab.context = { instanceId: 'tab', pageContext: {}, propertyPane: { refresh: function() {} } };
  host.tab.domElement = host.createElement();
  host.tab.domElement.isTab = true;
  host.canvas.appendChild(host.tab.domElement);
  host.tab.displayMode = 1;
  host.tab.findInAncestors = function() { return false; };
  host.tab.onInit();
  host.forms = host.canvas.zones.map(function(zone, index) {
    const form = new FormPart();
    form.properties = { listName: 'Test list', enableDynamicDiagnostics: false };
    form.context = { instanceId: 'form-' + index, propertyPane: { refresh: function() {} } };
    form.displayMode = 1;
    form.domElement = host.createElement();
    zone.appendChild(form.domElement);
    form.tryRebindDynamicReferences = function() {};
    form.ensureDynamicValueChangeHandler = function() {};
    form.getRuntimeDynamicItemProperty = function() {};
    form.getRuntimeDynamicItemModeProperty = function() {};
    form._getEnvironmentMessage = function() { return Promise.resolve(); };
    form.loadCount = 0;
    form.loadLists = function() {
      this.loadCount += 1;
      return new Promise(function(resolve, reject) { form.finishLoad = resolve; form.failLoad = reject; });
    };
    form.loadSitePages = function() { return Promise.resolve(); };
    form.loadListFields = function() { return Promise.resolve(); };
    form.refreshAvailableListControlSources = function() {};
    form.refreshAvailableDynamicFormSources = function() {};
    form.registerFormActivation();
    return form;
  });
  return host;
}

['DynamicForms', 'ReportForms'].forEach(function(project) {
  ['1', '2'].forEach(function(tabType) {
    [true, false].forEach(function(tabFirst) {
      test(project + ' first tab settles and preserves forms (mode ' + tabType + ', Tab first ' + tabFirst + ')', async function() {
        const host = createScenario(project, tabType);
        if (tabFirst) { host.tab.render(); }
        host.forms.forEach(function(form) { form.resolveInitialActivation(); form.render(); });
        if (!tabFirst) { host.tab.render(); }
        await host.settle();
        const first = host.forms[0];
        const second = host.forms[1];
        assert.strictEqual(first.loadCount, 1);
        assert.strictEqual(second.loadCount, 0, 'Hidden forms must remain lazy');
        assert.strictEqual(first.domElement.children[0].textContent,
          project === 'DynamicForms' ? 'Loading form...' : 'Loading report form...');
        assert.strictEqual(host.canvas.zones[0].classList.contains('sectionHidden'), false);
        first.finishLoad();
        await host.settle();
        assert.strictEqual(first._runtimeInitializationComplete, true);
        assert.strictEqual(first.domElement.mounts, 1, 'The first tab loads without clicking');
        const instance = first.domElement.formInstance;
        instance.values.note = 'Unsaved value';
        const unmounts = first.domElement.unmounts;
        host.document.readyState = 'complete';
        host.selectTab(1);
        await host.settle();
        assert.strictEqual(second.loadCount, 1);
        assert.strictEqual(first.domElement.style.display, 'none');
        second.finishLoad();
        await host.settle();
        host.selectTab(0);
        await host.settle();
        assert.strictEqual(first.domElement.style.display, '');
        assert.strictEqual(first.loadCount, 1);
        assert.strictEqual(first.domElement.formInstance, instance);
        assert.strictEqual(first.domElement.formInstance.values.note, 'Unsaved value');
        assert.strictEqual(first.domElement.unmounts, unmounts);
        first.renderedOnce = true;
        const writes = first.domElement.writes['data-spse-exclude-from-tabs'];
        const tabRenders = host.renderTabCount;
        first.onThemeChanged({ semanticColors: { bodyText: '#123456' } });
        first.onThemeChanged({ semanticColors: { bodyText: '#654321' } });
        await host.settle();
        assert.strictEqual(first.domElement.style['--bodyText'], '#654321');
        assert.strictEqual(first.domElement.writes['data-spse-exclude-from-tabs'], writes);
        assert.strictEqual(host.renderTabCount, tabRenders, 'Theme updates must not restart Tab rendering');
        assert.deepStrictEqual(host.errors, []);
      });
    });
  });

  test(project + ' ignores redundant activation but handles real changes during initialization', async function() {
    const host = createScenario(project, '2');
    host.tab.render();
    await host.settle();
    const first = host.forms[0];
    const unmounts = first.domElement.unmounts;
    const writes = first.domElement.writes['data-spse-exclude-from-tabs'];
    for (let index = 0; index < 10; index += 1) {
      host.tab.publishFormActivation(0, host.canvas.zones);
    }
    await host.settle();
    assert.strictEqual(first.domElement.unmounts, unmounts, 'Repeated activation must not reset the loading DOM');
    assert.strictEqual(first.domElement.writes['data-spse-exclude-from-tabs'], writes);
    host.selectTab(1);
    host.selectTab(0);
    await host.settle();
    assert.strictEqual(first.loadCount, 1, 'Tab switches during loading must not duplicate initialization');
    first.finishLoad();
    host.forms[1].finishLoad();
    await host.settle();
    assert.strictEqual(first._runtimeInitializationComplete, true);
    const instance = first.domElement.formInstance;
    host.window.dispatchEvent({ type: activationEvent, detail: {
      controllerType: 'tab', controlledZones: [host.createElement()], activeZone: undefined
    } });
    assert.strictEqual(first.domElement.formInstance, instance, 'Unrelated Tabs must not change activation');
  });

  test(project + ' exclusion changes remain reactive but unchanged writes do not refresh Tab', async function() {
    const host = createScenario(project, '2');
    host.tab.render();
    await host.settle();
    const form = host.forms[0];
    const initialRenders = host.renderTabCount;
    form.domElement.setAttribute('data-spse-exclude-from-tabs', 'false');
    await host.settle();
    assert.strictEqual(host.renderTabCount, initialRenders);
    form.displayMode = 2;
    form.properties.excludeFromTabs = true;
    form.render();
    await host.settle();
    assert.strictEqual(host.renderTabCount, initialRenders + 1);
    assert.strictEqual(host.tab.controlledZones.length, 1);
    form.render();
    await host.settle();
    assert.strictEqual(host.renderTabCount, initialRenders + 1);
    form.properties.excludeFromTabs = false;
    form.render();
    await host.settle();
    assert.strictEqual(host.renderTabCount, initialRenders + 2);
    assert.strictEqual(host.tab.controlledZones.length, 2);
  });

  test(project + ' still displays initialization failures and supports Retry', async function() {
    const host = createScenario(project, '2');
    host.tab.render();
    await host.settle();
    const form = host.forms[0];
    form.failLoad(new Error('Test initialization failure'));
    await host.settle();
    assert.strictEqual(host.errors.length, 1);
    assert.strictEqual(form.domElement.children[0].getAttribute('role'), 'alert');
    assert.ok(form.domElement.children[0].textContent.indexOf('Test initialization failure') >= 0);
    form.domElement.children[0].children[0].onclick();
    await host.settle();
    assert.strictEqual(form.loadCount, 2);
    form.finishLoad();
    await host.settle();
    assert.strictEqual(form._runtimeInitializationComplete, true);
  });
});
