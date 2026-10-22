'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const projects = [
  { name: 'PrintControl', file: 'printControl\\PrintControlWebPart.ts', theme: '_siteTheme' },
  { name: 'ReportForms', file: 'sharePointDynamicForm\\SharePointDynamicFormWebPart.ts', theme: '_currentTheme' },
  { name: 'ListControl', file: 'listControl\\ListControlWebPart.ts', theme: '_currentTheme' }
];

function createElement() {
  const element = {
    style: {
      setProperty: function(name, value) { this[name] = value; }
    },
    children: [],
    attributes: {},
    setAttribute: function(name, value) { this.attributes[name] = value; },
    appendChild: function(child) { this.children.push(child); },
    addEventListener: function() {}
  };
  Object.defineProperty(element, 'innerHTML', {
    set: function() { this.children = []; }
  });
  return element;
}

function loadWebPart(project) {
  const root = path.join(__dirname, '..', project.name);
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const source = fs.readFileSync(path.join(root, 'src', 'webparts', project.file), 'utf8');
  const output = ts.transpileModule(source, {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5 }
  }).outputText;
  const exports = {};
  const errors = [];
  const imports = {
    '@microsoft/sp-webpart-base': { BaseClientSideWebPart: function() {} },
    '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } },
    '@microsoft/sp-http': {},
    '@microsoft/sp-component-base/lib/DynamicProperty': {},
    '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField': {},
    '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCheckbox/PropertyPaneCheckbox': {},
    '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker': {},
    'react': require(path.join(root, 'node_modules', 'react')),
    'react-dom': {
      render: function(element, container) {
        if (container.renderError) { throw container.renderError; }
        container.reactElement = element;
        container.reactRenders = (container.reactRenders || 0) + 1;
      },
      unmountComponentAtNode: function() {}
    },
    'SharePointDynamicFormWebPartStrings': {},
    './components/SharePointDynamicForm': {},
    './components/RepeatedReportForms': {},
    './components/ListControl': { ListControl: function() {} },
    './components/ListDesigner': { ListDesigner: function() {} },
    'ListControlWebPartStrings': {},
    '../shared/deterministicFullWidth': { updateResponsiveOptionalFullWidth: function() {} },
    '../shared/overrideCss': { applyOverrideCss: function() {} },
    '../../../config/package-solution.json': require(path.join(root, 'config', 'package-solution.json'))
  };
  vm.runInNewContext(output, {
    exports: exports,
    console: { log: function() {}, error: function() { errors.push(Array.from(arguments)); } },
    document: { createElement: createElement },
    require: function(name) {
      assert.ok(Object.prototype.hasOwnProperty.call(imports, name), 'Unexpected import: ' + name);
      return imports[name];
    }
  }, { filename: project.file });
  exports.default.testErrors = errors;
  return exports.default;
}

let failures = 0;
let passed = 0;
const pending = [];
function test(name, run) {
  try {
    const result = run();
    if (result && typeof result.then === 'function') {
      pending.push(result.then(function() {
        passed += 1;
        console.log('PASS ' + name);
      }, function(error) {
        failures += 1;
        console.error('FAIL ' + name + ': ' + error.message);
      }));
    } else {
      passed += 1;
      console.log('PASS ' + name);
    }
  } catch (error) {
    failures += 1;
    console.error('FAIL ' + name + ': ' + error.message);
  }
}

projects.forEach(function(project) {
  const WebPart = loadWebPart(project);
  [false, true].forEach(function(hasDomElement) {
    test(project.name + ' caches the initial theme before properties exist (DOM ' + hasDomElement + ')', function() {
      const part = new WebPart();
      if (hasDomElement) { part.domElement = createElement(); }
      const theme = { isInverted: true, semanticColors: { bodyText: '#eeeeee' } };
      part.onThemeChanged(theme);
      assert.strictEqual(part[project.theme], theme);
      assert.strictEqual(part.properties, undefined);
      if (hasDomElement) { assert.strictEqual(part.domElement.children.length, 0); }
    });
  });

  test(project.name + ' uses the cached theme on first render and updates after rendering', function() {
    const part = new WebPart();
    const initial = {
      isInverted: true,
      semanticColors: { primaryButtonBackground: '#123456', bodyText: '#eeeeee', link: '#112233' }
    };
    part.onThemeChanged(initial);
    part.properties = { appearanceMode: 'siteTheme' };
    part.domElement = createElement();
    let refreshes = 0;
    part.context = {
      instanceId: 'test-instance',
      propertyPane: { refresh: function() { refreshes += 1; } }
    };
    if (project.name === 'ReportForms') {
      part.refreshFullWidthLayout = function() {};
      part.tryRebindDynamicReferences = function() {};
      part.ensureDynamicValueChangeHandler = function() {};
    }
    part.render();
    if (project.name === 'PrintControl') {
      assert.strictEqual(part.domElement.children[0].children[0].style.backgroundColor, '#123456');
    } else if (project.name === 'ReportForms') {
      assert.strictEqual(part.domElement.style['--bodyText'], '#eeeeee');
      assert.strictEqual(part.domElement.style.display, 'none');
      assert.strictEqual(part._isDarkTheme, true);
    } else {
      assert.strictEqual(part.domElement.reactElement.props.bodyTextColor, '#eeeeee');
    }
    let renders = 0;
    const render = part.render;
    part.render = function() {
      renders += 1;
      render.call(this);
    };
    const updated = {
      isInverted: false,
      semanticColors: { primaryButtonBackground: '#654321', bodyText: '#111111' }
    };
    part.onThemeChanged(updated);
    assert.strictEqual(part[project.theme], updated);
    assert.strictEqual(renders, 1);
    if (project.name === 'PrintControl') {
      assert.strictEqual(refreshes, 1);
      assert.strictEqual(part.domElement.children[0].children[0].style.backgroundColor, '#654321');
    } else if (project.name === 'ReportForms') {
      assert.strictEqual(part._isDarkTheme, false);
      assert.strictEqual(part.domElement.style['--bodyText'], '#111111');
    } else {
      assert.strictEqual(part.domElement.reactElement.props.bodyTextColor, '#111111');
    }
  });

  if (project.name === 'ListControl') {
    test('ListControl does not mount before onInit even when properties and DOM already exist', function() {
      const part = new WebPart();
      part.properties = {};
      part.domElement = createElement();
      let renders = 0;
      part.render = function() { renders += 1; };
      part.onThemeChanged({ semanticColors: { bodyText: '#123456' } });
      part.onThemeChanged({ semanticColors: { bodyText: '#654321' } });
      assert.strictEqual(renders, 0);
      assert.strictEqual(part._currentTheme.semanticColors.bodyText, '#654321');
    });

    test('ListControl theme updates still render after the designer is first displayed', function() {
      const part = new WebPart();
      part.properties = {};
      part.domElement = createElement();
      part.context = { instanceId: 'designer-test' };
      part._isListDesignerOpen = true;
      part.render();
      assert.strictEqual(part.domElement.reactRenders, 1);
      part.onThemeChanged({ semanticColors: { bodyText: '#123456' } });
      assert.strictEqual(part.domElement.reactRenders, 2);
    });

    test('ListControl logs actual render failures and does not treat a failed render as mounted', function() {
      const part = new WebPart();
      part.properties = { enableDiagnostics: false };
      part.domElement = createElement();
      part.context = { instanceId: 'failed-render-test' };
      const error = new Error('React mount failed');
      part.domElement.renderError = error;
      assert.throws(function() { part.render(); }, function(actual) { return actual === error; });
      assert.strictEqual(WebPart.testErrors[WebPart.testErrors.length - 1][1], error);
      assert.strictEqual(part._hasRenderedOnce, false);
      part.onThemeChanged({ semanticColors: { bodyText: '#abcdef' } });
      delete part.domElement.renderError;
      part.render();
      assert.strictEqual(part._hasRenderedOnce, true);
    });

    test('ListControl logs and rejects synchronous initialization failures with diagnostics disabled', function() {
      const part = new WebPart();
      const error = new Error('Dynamic data registration failed');
      part.properties = { enableDiagnostics: false };
      part.context = { dynamicDataSourceManager: { initializeSource: function() { throw error; } } };
      return part.onInit().then(function() {
        assert.fail('Initialization must reject');
      }, function(actual) {
        assert.strictEqual(actual, error);
        assert.ok(WebPart.testErrors.some(function(entry) {
          return entry[0] === '[ListControlWebPart] Initialization failed.' && entry[1] === error;
        }));
      });
    });

    test('ListControl logs and rejects asynchronous initialization failures with diagnostics disabled', function() {
      const part = new WebPart();
      const error = new Error('Metadata initialization failed');
      part.properties = { enableDiagnostics: false };
      part.context = {};
      part.discoverSearchSources = function() {};
      part.loadLists = function() { return Promise.reject(error); };
      part.loadSitePages = function() { return Promise.resolve(); };
      return part.onInit().then(function() {
        assert.fail('Initialization must reject');
      }, function(actual) {
        assert.strictEqual(actual, error);
        assert.ok(WebPart.testErrors.some(function(entry) {
          return entry[0] === '[ListControlWebPart] Initialization failed.' && entry[1] === error;
        }));
      });
    });

    test('ListControl finishes initialization in order before its first normal render uses the cached theme', function() {
      const part = new WebPart();
      const calls = [];
      part.properties = {
        listName: 'Items', viewId: 'view-id', viewColumns: [], linkTargetIdParam: 'ID',
        enableDiagnostics: false, appearanceMode: 'siteTheme'
      };
      part.context = { instanceId: 'initialized-test' };
      part.domElement = createElement();
      part.initializeDynamicDataSource = function() { calls.push('dynamic-data'); };
      part.discoverSearchSources = function() { calls.push('search'); };
      ['loadLists', 'loadSitePages', 'loadViews', 'loadListFields', 'loadViewColumns'].forEach(function(method) {
        part[method] = function() { calls.push(method); return Promise.resolve(); };
      });
      const initialization = part.onInit();
      part.onThemeChanged({ semanticColors: { bodyText: '#102030' } });
      return initialization.then(function() {
        assert.deepStrictEqual(calls, ['dynamic-data', 'search', 'loadLists', 'loadSitePages', 'loadViews', 'loadListFields', 'loadViewColumns']);
        assert.strictEqual(part.properties.linkTargetIdParam, 'itemid');
        assert.strictEqual(part.domElement.reactRenders, undefined);
        part.render();
        assert.strictEqual(part.domElement.reactElement.props.bodyTextColor, '#102030');
        assert.strictEqual(part.domElement.reactRenders, 1);
      });
    });
  }
});

Promise.all(pending).then(function() {
  console.log(passed + ' passed, ' + failures + ' failed');
  process.exitCode = failures ? 1 : 0;
});
