'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const projects = [
  { name: 'PrintControl', file: 'printControl\\PrintControlWebPart.ts', theme: '_siteTheme' },
  { name: 'ReportForms', file: 'sharePointDynamicForm\\SharePointDynamicFormWebPart.ts', theme: '_currentTheme' }
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
  const imports = {
    '@microsoft/sp-webpart-base': { BaseClientSideWebPart: function() {} },
    '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } },
    '@microsoft/sp-http': {},
    '@microsoft/sp-component-base/lib/DynamicProperty': {},
    '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField': {},
    '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCheckbox/PropertyPaneCheckbox': {},
    '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker': {},
    'react': {},
    'react-dom': {},
    'SharePointDynamicFormWebPartStrings': {},
    './components/SharePointDynamicForm': {},
    './components/RepeatedReportForms': {},
    '../shared/deterministicFullWidth': {},
    '../shared/overrideCss': { applyOverrideCss: function() {} },
    '../../../config/package-solution.json': require(path.join(root, 'config', 'package-solution.json'))
  };
  vm.runInNewContext(output, {
    exports: exports,
    document: { createElement: createElement },
    require: function(name) {
      assert.ok(Object.prototype.hasOwnProperty.call(imports, name), 'Unexpected import: ' + name);
      return imports[name];
    }
  }, { filename: project.file });
  return exports.default;
}

let failures = 0;
let passed = 0;
function test(name, run) {
  try {
    run();
    passed += 1;
    console.log('PASS ' + name);
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
    } else {
      assert.strictEqual(part.domElement.style['--bodyText'], '#eeeeee');
      assert.strictEqual(part.domElement.style.display, 'none');
      assert.strictEqual(part._isDarkTheme, true);
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
    } else {
      assert.strictEqual(part._isDarkTheme, false);
      assert.strictEqual(part.domElement.style['--bodyText'], '#111111');
    }
  });
});

console.log(passed + ' passed, ' + failures + ' failed');
process.exitCode = failures ? 1 : 0;
