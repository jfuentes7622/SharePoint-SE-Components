'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

function serializeXml(node) {
  function escape(value) { return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); }
  if (node.nodeType === 9) { return serializeXml(node.documentElement); }
  if (node.nodeType === 3) { return escape(node.nodeValue); }
  return '<' + node.nodeName + Array.from(node.attributes).map(attribute =>
    ' ' + attribute.name + '="' + escape(attribute.value) + '"').join('') + '>'
    + Array.from(node.childNodes).map(serializeXml).join('') + '</' + node.nodeName + '>';
}

module.exports = function(root, window) {
  const project = path.basename(root);
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const source = path.join(root, 'src', 'webparts', project === 'ListControl' ? 'listControl' : 'gridControl', 'components', 'LibraryBrowser.ts');
  const exports = {};
  vm.runInNewContext(ts.transpileModule(fs.readFileSync(source, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5 }
  }).outputText, { exports, DOMParser: window.DOMParser,
    XMLSerializer: function() { this.serializeToString = serializeXml; }, Date, Promise });
  return exports;
};

module.exports.serializeXml = serializeXml;
