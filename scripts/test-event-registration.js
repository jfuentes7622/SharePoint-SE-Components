'use strict';

// Run the legacy component tests without Karma's unavailable @ms/sp-client-shared.
// Only SharePoint/PnP host controls are mocked; application modules and React run normally.
const fs = require('fs');
const path = require('path');
const Module = require('module');
const root = path.resolve(__dirname, '..');
const dependencies = path.join(root, 'EventRegistration', 'node_modules');
const ts = require(path.join(dependencies, 'typescript'));
const Mocha = require(path.join(dependencies, 'mocha'));
const jsdom = require(path.join(dependencies, 'jsdom'));
const document = jsdom.jsdom('', { url: 'https://example.test/sites/events/SitePages/Register.aspx' });

global.window = document.defaultView;
global.document = document;
global.navigator = window.navigator;
global.DOMParser = window.DOMParser;

const mocks = {
  '@microsoft/sp-webpart-base': { BaseClientSideWebPart: function BaseClientSideWebPart() {} },
  '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } },
  '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker': {},
  '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData': {},
  '@pnp/spfx-controls-react/lib/WebPartTitle': {},
  '@pnp/spfx-controls-react/lib/Placeholder': {},
  CalendarWebPartStrings: {}
};
const originalLoad = Module._load;
Module._load = function(request, parent, isMain) {
  return Object.prototype.hasOwnProperty.call(mocks, request)
    ? mocks[request] : originalLoad.call(this, request, parent, isMain);
};
require.extensions['.scss'] = function(module) { module.exports = { default: {} }; };
require.extensions['.ts'] = require.extensions['.tsx'] = function(module, filename) {
  const output = ts.transpileModule(fs.readFileSync(filename, 'utf8'), {
    fileName: filename,
    compilerOptions: { target: ts.ScriptTarget.ES5, module: ts.ModuleKind.CommonJS, jsx: ts.JsxEmit.React }
  });
  module._compile(output.outputText, filename);
};

const mocha = new Mocha({ reporter: 'spec' });
[
  path.join(root, 'EventRegistration', 'src', 'webparts', 'eventRegistration', 'tests'),
  path.join(root, 'Calendar', 'src', 'webparts', 'calendar', 'tests')
].forEach(function(directory) {
  fs.readdirSync(directory).filter(function(file) { return /\.test\.ts$/.test(file); })
    .sort().forEach(function(file) { mocha.addFile(path.join(directory, file)); });
});
mocha.run(function(failures) {
  Module._load = originalLoad;
  window.close();
  process.exitCode = failures ? 1 : 0;
});
