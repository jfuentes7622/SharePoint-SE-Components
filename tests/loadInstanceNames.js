'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

module.exports = function(root) {
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const file = path.join(root, 'src', 'webparts', 'shared', 'instanceNames.ts');
  const exports = {};
  vm.runInNewContext(ts.transpileModule(fs.readFileSync(file, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5 }
  }).outputText, {
    exports,
    require(name) {
      if (name === '@microsoft/sp-core-library') { return { DisplayMode: { Read: 1, Edit: 2 } }; }
      if (name === '@microsoft/sp-http') { return { SPHttpClient: { configurations: { v1: {} } } }; }
      throw new Error('Unexpected naming import: ' + name);
    }
  }, { filename: file });
  return exports;
};
