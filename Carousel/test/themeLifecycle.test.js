'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ts = require('typescript');

const source = fs.readFileSync(path.join(
  __dirname, '..', 'src', 'webparts', 'carousel', 'CarouselWebPart.ts'
), 'utf8');
const compiled = ts.transpileModule(source, {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5 }
}).outputText;

function BaseClientSideWebPart() {}

const context = {
  exports: {},
  require: (name) => name === '@microsoft/sp-webpart-base'
    ? { BaseClientSideWebPart: BaseClientSideWebPart }
    : {}
};
vm.runInNewContext(compiled, context, { filename: 'CarouselWebPart.js' });
const CarouselWebPart = context.exports.default;
const webPart = new CarouselWebPart();
let renderCalls = 0;
webPart.render = () => { renderCalls += 1; };
Object.defineProperty(webPart, 'properties', {
  get: () => { throw new Error('Properties are not initialized'); }
});
Object.defineProperty(webPart, 'domElement', {
  get: () => { throw new Error('DOM is not initialized'); }
});

const initialTheme = { semanticColors: { bodyBackground: '#ffffff' } };
webPart.renderedOnce = false;
assert.doesNotThrow(() => webPart.onThemeChanged(initialTheme));
assert.strictEqual(renderCalls, 0, 'Early theme notification must not render');
assert.strictEqual(webPart.currentTheme, initialTheme, 'First render must retain the early theme');

const pendingTheme = { semanticColors: { bodyBackground: '#eeeeee' } };
webPart.onThemeChanged(pendingTheme);
assert.strictEqual(renderCalls, 0, 'Repeated early notifications must not render');
assert.strictEqual(webPart.currentTheme, pendingTheme, 'Latest early theme must be retained');

webPart.renderedOnce = true;
const updatedTheme = { semanticColors: { bodyBackground: '#222222' } };
webPart.onThemeChanged(updatedTheme);
assert.strictEqual(renderCalls, 1, 'Theme changes after first render must refresh the web part');
assert.strictEqual(webPart.currentTheme, updatedTheme);
webPart.onThemeChanged(undefined);
assert.strictEqual(renderCalls, 2, 'Clearing the theme must also refresh an initialized web part');
assert.strictEqual(webPart.currentTheme, undefined);

console.log('Carousel theme lifecycle tests passed.');
