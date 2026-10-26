'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = path.join(__dirname, '..', 'Tiles');
const base = path.join(root, 'src', 'webparts', 'tiles');
const ts = require(path.join(root, 'node_modules', 'typescript'));
const React = require(path.join(root, 'node_modules', 'react'));
const jsdom = require(path.join(root, 'node_modules', 'jsdom'));
const document = jsdom.jsdom('<!doctype html><html><body></body></html>');
const DisplayMode = { Read: 1, Edit: 2 };
let nextGuid = 0;
const collectionFieldTypes = { string: 1, number: 2, boolean: 3, dropdown: 4, fabricIcon: 5, url: 6, custom: 7 };
function DefaultButton() {}
function PrimaryButton() {}
function TextField() {}
function Dropdown() {}
function Checkbox() {}
function Panel() {}
let strings;
vm.runInNewContext(fs.readFileSync(path.join(base, 'loc', 'en-us.js'), 'utf8'), {
  define: function(dependencies, factory) { strings = factory(); }
});

function field(type) {
  return function(targetProperty, properties) {
    return { type: type, targetProperty: targetProperty, properties: properties };
  };
}
const spWebPart = { BaseClientSideWebPart: function() {}, PropertyPaneFieldType: { Custom: 1 } };
spWebPart.BaseClientSideWebPart.prototype.onPropertyPaneFieldChanged = function() {};
['Checkbox', 'Dropdown', 'Label', 'Slider', 'Toggle', 'TextField'].forEach(function(type) {
  spWebPart['PropertyPane' + type] = field(type);
});

function load(file, extraImports) {
  const imports = Object.assign({
    react: React,
    'react-dom': {
      render: function(element, container) {
        container.rendered = element;
        container.renderCount = (container.renderCount || 0) + 1;
      }
    },
    '@microsoft/sp-core-library': {
      DisplayMode: DisplayMode, Guid: { newGuid: function() { nextGuid += 1; return 'tile-' + nextGuid; } }
    },
    '@microsoft/sp-lodash-subset': { assign: Object.assign },
    '@microsoft/sp-webpart-base': spWebPart,
    '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker': {
      PropertyFieldColorPicker: field('ColorPicker'), PropertyFieldColorPickerStyle: { Inline: 1 }
    },
    '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData': {
      PropertyFieldCollectionData: field('Collection'), CustomCollectionFieldType: collectionFieldTypes
    },
    './components/TileCollectionEditor': { PropertyPaneTileCollection: field('Collection') },
    'office-ui-fabric-react/lib/Button': { DefaultButton: DefaultButton, PrimaryButton: PrimaryButton },
    'office-ui-fabric-react/lib/TextField': { TextField: TextField },
    'office-ui-fabric-react/lib/Dropdown': { Dropdown: Dropdown },
    'office-ui-fabric-react/lib/Checkbox': { Checkbox: Checkbox },
    'office-ui-fabric-react/lib/Panel': { Panel: Panel, PanelType: { large: 3 } },
    'sp-client-custom-fields/lib/PropertyFieldPicturePickerHost': { default: function PicturePickerHost() {} },
    TilesWebPartStrings: strings,
    './Tiles.module.scss': { default: { tiles: 'tiles', webPartTitle: 'webPartTitle', tilesList: 'tilesList' } },
    './Tile/Tile': { Tile: function() { return React.createElement('div'); } },
    '@pnp/spfx-controls-react/lib/Placeholder': { Placeholder: function() { return React.createElement('div'); } },
    '../shared/overrideCss': { applyOverrideCss: function() {}, PropertyPaneOverrideCss: field('OverrideCss') },
    '../../../config/package-solution.json': require(path.join(root, 'config', 'package-solution.json'))
  }, extraImports || {});
  const exports = {};
  const output = ts.transpileModule(fs.readFileSync(file, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
  }).outputText;
  vm.runInNewContext(output, {
    exports: exports, document: document, window: document.defaultView, console: console,
    require: function(name) {
      assert.ok(Object.prototype.hasOwnProperty.call(imports, name), 'Unexpected import: ' + name);
      return imports[name];
    }
  }, { filename: file });
  return exports;
}

const Tile = load(path.join(base, 'components', 'Tile', 'Tile.tsx'), {
  './Tile.module.scss': { default: { tile: 'tile', tileLink: 'tileLink', tileLinkWithImage: 'tileLinkWithImage' } }
}).Tile;
const Tiles = load(path.join(base, 'components', 'Tiles.tsx'), { './Tile/Tile': { Tile: Tile } }).Tiles;
const WebPart = load(path.join(base, 'TilesWebPart.ts'), { './components/Tiles': { Tiles: Tiles } }).default;
const editorModule = load(path.join(base, 'components', 'TileCollectionEditor.tsx'));
const defaults = require(path.join(base, 'TilesWebPart.manifest.json')).preconfiguredEntries[0].properties;
function fixture() {
  const part = new WebPart();
  part.properties = Object.assign({}, defaults, {
    showTitle: true, title: 'Original title', collectionData: [{ title: 'Tile label', url: '/destination' }]
  });
  part.displayMode = DisplayMode.Edit;
  part.domElement = document.createElement('div');
  part.refreshes = 0;
  part.context = {
    instanceId: 'tiles-title-test',
    propertyPane: { open: function() {}, refresh: function() { part.refreshes += 1; } }
  };
  part.render();
  return part;
}
function find(element, predicate) {
  if (!element || !element.props) { return undefined; }
  if (predicate(element)) { return element; }
  let match;
  React.Children.forEach(element.props.children, function(child) {
    if (!match) { match = find(child, predicate); }
  });
  return match;
}
function tree(part) {
  return new Tiles(part.domElement.rendered.props).render();
}
function input(part) {
  return find(tree(part), function(element) { return element.type === 'input'; });
}
function setting(part, targetProperty) {
  const fields = [];
  part.getPropertyPaneConfiguration().pages.forEach(function(page) {
    page.groups.forEach(function(group) {
      group.groupFields.forEach(function(item) {
        if (item.targetProperty === targetProperty) { fields.push(item); }
      });
    });
  });
  assert.strictEqual(fields.length, 1, 'A setting must be available for ' + targetProperty);
  return fields[0];
}
function titleField(part) { return setting(part, 'title'); }
function renderedTiles(part) {
  const list = find(tree(part), function(element) { return element.props.className === 'tilesList'; });
  return React.Children.toArray(list.props.children).map(function(element) { return new Tile(element.props).render(); });
}
let passed = 0;
let failed = 0;
function test(name, run) {
  try {
    run();
    passed += 1;
    console.log('PASS ' + name);
  } catch (error) {
    failed += 1;
    console.error('FAIL ' + name + ': ' + error.message);
  }
}

test('the initial theme is cached before host initialization and used on the first render', function() {
  const part = new WebPart();
  const initial = { semanticColors: { primaryButtonBackground: '#123456' } };
  part.onThemeChanged(initial);
  assert.strictEqual(part._siteTheme, initial);
  part.properties = Object.assign({}, defaults, { appearanceMode: 'siteTheme' });
  part.domElement = document.createElement('div');
  let refreshes = 0;
  part.context = {
    instanceId: 'tiles-theme-test',
    propertyPane: { refresh: function() { refreshes += 1; } }
  };
  part.onThemeChanged(initial);
  assert.strictEqual(part.domElement.renderCount, undefined);
  assert.strictEqual(refreshes, 0);
  part.render();
  assert.strictEqual(part.domElement.rendered.props.backgroundColor, '#123456');
  part.renderedOnce = true;
  part.onThemeChanged({ semanticColors: { primaryButtonBackground: '#654321' } });
  assert.strictEqual(part.domElement.renderCount, 2);
  assert.strictEqual(part.domElement.rendered.props.backgroundColor, '#654321');
  assert.strictEqual(refreshes, 1);
});

test('typing consecutive characters updates the controlled inline title immediately', function() {
  const part = fixture();
  const originalTiles = part.properties.collectionData;
  ['N', 'Ne', 'New title'].forEach(function(value) {
    input(part).props.onChange({ currentTarget: { value: value } });
    assert.strictEqual(part.properties.title, value);
    assert.strictEqual(input(part).props.value, value);
  });
  assert.strictEqual(part.properties.collectionData, originalTiles);
  assert.strictEqual(part.properties.collectionData[0].title, 'Tile label');
});

test('clearing the inline title does not restore the previous text', function() {
  const part = fixture();
  input(part).props.onChange({ currentTarget: { value: '' } });
  assert.strictEqual(part.properties.title, '');
  assert.strictEqual(input(part).props.value, '');
});

test('the property pane exposes an editable title field with the saved value', function() {
  const part = fixture();
  const title = titleField(part);
  assert.strictEqual(title.type, 'TextField');
  assert.strictEqual(title.properties.label, strings.titleField);
  assert.strictEqual(title.properties.value, 'Original title');
  assert.notStrictEqual(title.properties.disabled, true);
});

test('inline changes refresh the title property-pane field', function() {
  const part = fixture();
  input(part).props.onChange({ currentTarget: { value: 'Edited inline' } });
  assert.ok(part.refreshes > 0, 'The open property pane must receive the changed title');
  assert.strictEqual(titleField(part).properties.value, 'Edited inline');
});

test('property-pane title changes are reflected in the inline editor', function() {
  const part = fixture();
  assert.strictEqual(titleField(part).targetProperty, 'title');
  part.properties.title = 'Changed in settings';
  part.onPropertyPaneFieldChanged('title', 'Original title', part.properties.title);
  part.render();
  assert.strictEqual(input(part).props.value, 'Changed in settings');
});

test('hiding and showing the title preserves its edited value', function() {
  const part = fixture();
  input(part).props.onChange({ currentTarget: { value: 'Keep this title' } });
  part.properties.showTitle = false;
  part.render();
  assert.strictEqual(input(part), undefined);
  assert.strictEqual(part.properties.title, 'Keep this title');
  part.properties.showTitle = true;
  part.render();
  assert.strictEqual(input(part).props.value, 'Keep this title');
});

test('published and restored properties display the edited title as read-only text', function() {
  const part = fixture();
  input(part).props.onChange({ currentTarget: { value: 'Published title' } });
  const restored = fixture();
  restored.properties = JSON.parse(JSON.stringify(part.properties));
  restored.displayMode = DisplayMode.Read;
  restored.render();
  assert.strictEqual(input(restored), undefined);
  const title = find(tree(restored), function(element) { return element.type === 'span'; });
  assert.strictEqual(title.props.children, 'Published title');
});

test('the title remains editable with site-theme styling or while hidden', function() {
  const part = fixture();
  part.properties.appearanceMode = 'siteTheme';
  part.properties.showTitle = false;
  assert.notStrictEqual(titleField(part).properties.disabled, true);
  part.properties.title = 'Preset hidden title';
  part.properties.showTitle = true;
  part.render();
  assert.strictEqual(input(part).props.value, 'Preset hidden title');
});

test('legacy tiles retain their background-colored 2px solid borders and shape radii', function() {
  const part = fixture();
  ['tileBorderColor', 'tileBorderWidth', 'tileBorderStyle', 'tileCornerRadius'].forEach(function(name) {
    delete part.properties[name];
  });
  part.properties.collectionData[0].color = '#123456';
  ['rounded', 'squared', 'round'].forEach(function(shape, index) {
    part.properties.tileShape = shape;
    part.render();
    const style = renderedTiles(part)[0].props.style;
    assert.strictEqual(style['--tile-border-color'], '#123456');
    assert.strictEqual(style['--tile-border-width'], '2px');
    assert.strictEqual(style['--tile-border-style'], 'solid');
    assert.strictEqual(style['--border-radius'], ['10px', '0px', '9999px'][index]);
  });
});

test('all tile variants receive the configured border without changing the outer web-part border', function() {
  const part = fixture();
  Object.assign(part.properties, {
    tileBorderColor: '#0078d4', tileBorderWidth: 5, tileBorderStyle: 'dashed', tileCornerRadius: 24,
    webPartBorderColor: '#ff0000', webPartBorderWidth: 1, webPartBorderStyle: 'dotted',
    collectionData: [
      { title: 'Normal', url: '#normal' },
      { title: 'Image only', url: '#image', imageUrl: 'image.png', imageOnly: true },
      { title: 'Background image', url: '#background', imageUrl: 'image.png', imageAsBackground: true }
    ]
  });
  part.render();
  renderedTiles(part).forEach(function(tile) {
    assert.strictEqual(tile.props.style['--tile-border-color'], '#0078d4');
    assert.strictEqual(tile.props.style['--tile-border-width'], '5px');
    assert.strictEqual(tile.props.style['--tile-border-style'], 'dashed');
    assert.strictEqual(tile.props.style['--border-radius'], '24px');
    assert.strictEqual(tile.props.style.height, '140px');
  });
  assert.strictEqual(tree(part).props.style.borderColor, '#ff0000');
  assert.strictEqual(tree(part).props.style.borderWidth, '1px');
  assert.strictEqual(tree(part).props.style.borderStyle, 'dotted');
});

test('zero-width and none-style borders are preserved along with every supported style', function() {
  const part = fixture();
  part.properties.tileBorderWidth = 0;
  ['none', 'solid', 'dashed', 'dotted', 'double'].forEach(function(style) {
    part.properties.tileBorderStyle = style;
    part.render();
    const tile = renderedTiles(part)[0];
    assert.strictEqual(tile.props.style['--tile-border-width'], '0px');
    assert.strictEqual(tile.props.style['--tile-border-style'], style);
  });
});

test('corner radius accepts zero and custom values without overriding round or square shapes', function() {
  const part = fixture();
  part.properties.tileCornerRadius = 0;
  part.render();
  assert.strictEqual(renderedTiles(part)[0].props.style['--border-radius'], '0px');
  part.properties.tileCornerRadius = 37;
  part.render();
  assert.strictEqual(renderedTiles(part)[0].props.style['--border-radius'], '37px');
  part.properties.tileShape = 'round';
  part.render();
  assert.strictEqual(renderedTiles(part)[0].props.style['--border-radius'], '9999px');
  part.properties.tileShape = 'squared';
  part.render();
  assert.strictEqual(renderedTiles(part)[0].props.style['--border-radius'], '0px');
});

test('tile border settings are exposed with independent color, width, style and corner controls', function() {
  const part = fixture();
  assert.strictEqual(setting(part, 'tileBorderColor').type, 'ColorPicker');
  assert.strictEqual(setting(part, 'tileBorderWidth').properties.value, 2);
  assert.strictEqual(setting(part, 'tileBorderWidth').properties.min, 0);
  assert.strictEqual(setting(part, 'tileBorderWidth').properties.max, 20);
  assert.deepStrictEqual(Array.from(setting(part, 'tileBorderStyle').properties.options, function(option) { return option.key; }),
    ['none', 'solid', 'dashed', 'dotted', 'double']);
  assert.strictEqual(setting(part, 'tileCornerRadius').properties.value, 10);
  assert.notStrictEqual(setting(part, 'tileCornerRadius').properties.disabled, true);
  part.properties.tileShape = 'round';
  part.onPropertyPaneFieldChanged('tileShape', 'rounded', 'round');
  assert.ok(part.refreshes > 0, 'Changing shape must refresh the radius setting');
  assert.strictEqual(setting(part, 'tileCornerRadius').properties.disabled, true);
});

test('explicit border colors respect manual, site-theme and site-theme-with-overrides modes', function() {
  const part = fixture();
  part.properties.tileBorderColor = '#123456';
  part._siteTheme = { semanticColors: { variantBorder: '#abcdef' } };
  part.render();
  assert.strictEqual(renderedTiles(part)[0].props.style['--tile-border-color'], '#123456');
  part.properties.appearanceMode = 'siteTheme';
  part.render();
  assert.strictEqual(renderedTiles(part)[0].props.style['--tile-border-color'], '#abcdef');
  assert.strictEqual(setting(part, 'tileBorderColor').properties.disabled, true);
  part.properties.appearanceMode = 'siteThemeOverrides';
  part.render();
  assert.strictEqual(renderedTiles(part)[0].props.style['--tile-border-color'], '#123456');
});

test('border settings survive restored read-mode properties', function() {
  const part = fixture();
  Object.assign(part.properties, { tileBorderColor: '#008800', tileBorderWidth: 6, tileBorderStyle: 'double', tileCornerRadius: 18 });
  const restored = fixture();
  restored.properties = JSON.parse(JSON.stringify(part.properties));
  restored.displayMode = DisplayMode.Read;
  restored.render();
  assert.strictEqual(renderedTiles(restored)[0].props.style['--tile-border-color'], '#008800');
  assert.strictEqual(renderedTiles(restored)[0].props.style['--tile-border-width'], '6px');
  assert.strictEqual(renderedTiles(restored)[0].props.style['--tile-border-style'], 'double');
  assert.strictEqual(renderedTiles(restored)[0].props.style['--border-radius'], '18px');
});

test('border thickness and corner radius enforce their exact supported pixel limits', function() {
  const part = fixture();
  [
    { width: -1, radius: -1, expectedWidth: '0px', expectedRadius: '0px' },
    { width: 20, radius: 200, expectedWidth: '20px', expectedRadius: '200px' },
    { width: 21, radius: 201, expectedWidth: '20px', expectedRadius: '200px' }
  ].forEach(function(values) {
    part.properties.tileBorderWidth = values.width;
    part.properties.tileCornerRadius = values.radius;
    part.render();
    const style = renderedTiles(part)[0].props.style;
    assert.strictEqual(style['--tile-border-width'], values.expectedWidth);
    assert.strictEqual(style['--border-radius'], values.expectedRadius);
  });
});

function editorFixture(values) {
  const part = fixture();
  const properties = setting(part, 'collectionData').properties;
  const saved = [];
  const editor = new editorModule.TileCollectionEditor(Object.assign({}, properties, {
    value: values || [{ title: 'Existing', url: '/existing' }],
    onChanged: function(items) { saved.push(items); }
  }));
  editor.setState = function(update) {
    const next = typeof update === 'function' ? update(this.state, this.props) : update;
    this.state = Object.assign({}, this.state, next);
  };
  editor.openPanel();
  return { editor: editor, saved: saved };
}
function saveButton(editor) {
  return find(editor.render(), function(element) { return element.type === PrimaryButton; });
}
function draftField(editor, name) {
  const field = editor.props.fields.filter(function(item) { return item.id === name; })[0];
  return editor.renderField(editor.state.draft, field, true);
}
function changeDraft(editor, name, value) {
  editor.updateField(editor.state.draft.uniqueId, name, value);
}

test('the tile editor retains both required columns and all optional editor fields', function() {
  const editor = editorFixture().editor;
  assert.deepStrictEqual(Array.from(editor.props.fields.filter(function(item) { return item.required; }),
    function(item) { return item.id; }), ['title', 'url']);
  assert.strictEqual(editor.props.fields.length, 14);
  ['imageUrl', 'imageOpacity', 'hoverImageUrl', 'color'].forEach(function(id) {
    assert.strictEqual(typeof editor.props.fields.filter(function(item) { return item.id === id; })[0].onCustomRender, 'function');
  });
});

test('an untouched blank draft does not block saving existing tiles or create a phantom tile', function() {
  const result = editorFixture();
  const editor = result.editor;
  assert.strictEqual(saveButton(editor).props.disabled, false);
  assert.strictEqual(saveButton(editor).props.text, strings.tileSaveButton);
  saveButton(editor).props.onClick();
  assert.strictEqual(result.saved.length, 1);
  assert.strictEqual(result.saved[0].length, 1);
  assert.strictEqual(result.saved[0][0].title, 'Existing');
  assert.strictEqual(editor.state.open, false);
});

test('a title-only draft disables Save and shows URL is required without dropping the draft', function() {
  const result = editorFixture([]);
  const editor = result.editor;
  changeDraft(editor, 'title', 'New tile');
  assert.strictEqual(saveButton(editor).props.disabled, true);
  assert.strictEqual(draftField(editor, 'url').props.errorMessage, 'URL is required.');
  assert.strictEqual(draftField(editor, 'title').props.errorMessage, '');
  editor.save();
  assert.strictEqual(result.saved.length, 0);
  assert.strictEqual(editor.state.open, true);
  assert.strictEqual(editor.state.draft.title, 'New tile');
  assert.ok(find(editor.render(), function(element) {
    return element.props.role === 'alert' && element.props.children === strings.tileRequiredFieldsHelp;
  }));
});

test('finishing required fields enables Save and add and persists the unfinished row exactly once', function() {
  const result = editorFixture();
  const editor = result.editor;
  changeDraft(editor, 'url', '/new');
  assert.strictEqual(draftField(editor, 'title').props.errorMessage, 'Title is required.');
  assert.strictEqual(saveButton(editor).props.disabled, true);
  changeDraft(editor, 'title', 'New');
  assert.strictEqual(saveButton(editor).props.disabled, false);
  assert.strictEqual(saveButton(editor).props.text, strings.tileSaveAndAddButton);
  saveButton(editor).props.onClick();
  assert.strictEqual(result.saved[0].length, 2);
  assert.strictEqual(result.saved[0][1].title, 'New');
  assert.strictEqual(result.saved[0][1].url, '/new');
  assert.strictEqual(result.saved[0][1].imageOpacity, 100);
  assert.strictEqual(result.saved[0][1].imageAsBackground, false);
});

test('optional-only and whitespace-only drafts cannot silently disappear through Save', function() {
  const editor = editorFixture().editor;
  changeDraft(editor, 'description', 'Keep this unfinished description');
  assert.strictEqual(saveButton(editor).props.disabled, true);
  assert.strictEqual(draftField(editor, 'url').props.errorMessage, 'URL is required.');
  changeDraft(editor, 'title', '   ');
  changeDraft(editor, 'url', '/target');
  assert.strictEqual(saveButton(editor).props.disabled, true);
  assert.strictEqual(draftField(editor, 'title').props.errorMessage, 'Title is required.');
  changeDraft(editor, 'title', 'Ready');
  assert.strictEqual(saveButton(editor).props.disabled, false);
  changeDraft(editor, 'url', '\t ');
  assert.strictEqual(saveButton(editor).props.disabled, true);
});

test('invalid edits to existing tiles block Save immediately and retain every typed value', function() {
  const result = editorFixture();
  const editor = result.editor;
  const id = editor.state.items[0].uniqueId;
  editor.updateField(id, 'url', '');
  editor.updateField(id, 'description', 'Unsaved edit');
  assert.strictEqual(saveButton(editor).props.disabled, true);
  editor.save();
  assert.strictEqual(editor.state.open, true);
  assert.strictEqual(result.saved.length, 0);
  assert.strictEqual(editor.state.items[0].description, 'Unsaved edit');
  editor.updateField(id, 'url', '#restored');
  saveButton(editor).props.onClick();
  assert.strictEqual(result.saved[0][0].description, 'Unsaved edit');
  assert.strictEqual(result.saved[0][0].url, '#restored');
});

test('adding a valid draft resets only the draft and allows saving without duplicating it', function() {
  const result = editorFixture([]);
  const editor = result.editor;
  changeDraft(editor, 'title', 'Added');
  changeDraft(editor, 'url', '/added');
  editor.addDraft();
  assert.strictEqual(editor.state.items.length, 1);
  assert.strictEqual(editor.hasDraftChanges(), false);
  saveButton(editor).props.onClick();
  assert.strictEqual(result.saved[0].length, 1);
});

test('sorting and deleting retain row validation and can save an intentionally empty collection', function() {
  const result = editorFixture([{ title: 'First', url: '/first' }, { title: 'Second', url: '/second' }]);
  const editor = result.editor;
  const firstId = editor.state.items[0].uniqueId;
  editor.updateField(firstId, 'url', '');
  editor.moveItem(firstId, 1);
  assert.strictEqual(saveButton(editor).props.disabled, true);
  editor.deleteItem(firstId);
  assert.strictEqual(saveButton(editor).props.disabled, false);
  editor.deleteItem(editor.state.items[0].uniqueId);
  saveButton(editor).props.onClick();
  assert.strictEqual(result.saved[0].length, 0);
});

test('save serializes sorting order and preserves image, hover, color and target settings', function() {
  const original = [
    { title: 'First', url: '/first', imageUrl: '/photo.png', imageOpacity: 0, imageOnly: true,
      hoverImageUrl: '/hover.png', hoverImagePosition: 'left', color: '#008800', target: '_blank' },
    { title: 'Second', url: '/second' }
  ];
  const result = editorFixture(original);
  const editor = result.editor;
  editor.moveItem(editor.state.items[1].uniqueId, 0);
  saveButton(editor).props.onClick();
  assert.strictEqual(result.saved[0][0].title, 'Second');
  assert.strictEqual(result.saved[0][0].sortIdx, 1);
  assert.strictEqual(result.saved[0][1].sortIdx, 2);
  Object.keys(original[0]).forEach(function(key) {
    assert.strictEqual(result.saved[0][1][key], original[0][key]);
  });
  assert.strictEqual(original[0].uniqueId, undefined);
  assert.strictEqual(original[0].sortIdx, undefined);
});

test('Cancel and panel dismissal leave saved values unchanged and reopening discards edits', function() {
  const original = [{ title: 'Original', url: '/original' }];
  const result = editorFixture(original);
  const editor = result.editor;
  editor.updateField(editor.state.items[0].uniqueId, 'title', 'Unsaved');
  changeDraft(editor, 'title', 'Partial');
  find(editor.render(), function(element) { return element.type === Panel; }).props.onDismiss();
  assert.strictEqual(result.saved.length, 0);
  assert.strictEqual(original[0].title, 'Original');
  editor.openPanel();
  assert.strictEqual(editor.state.items[0].title, 'Original');
  assert.strictEqual(editor.hasDraftChanges(), false);
});

test('custom image controls keep their update callbacks and custom errors also disable Save', function() {
  const result = editorFixture();
  const editor = result.editor;
  const field = editor.props.fields.filter(function(item) { return item.id === 'imageUrl'; })[0];
  const wrapper = editor.renderField(editor.state.draft, field, true);
  const picker = React.Children.toArray(wrapper.props.children)[0];
  picker.props.onPropertyChange('imageUrl', '', '/selected.png');
  assert.strictEqual(editor.state.draft.imageUrl, '/selected.png');
  assert.strictEqual(saveButton(editor).props.disabled, true);
  changeDraft(editor, 'title', 'Picture tile');
  changeDraft(editor, 'url', '/target');
  editor.validateCustomField(editor.state.draft.uniqueId, 'imageUrl', 'Invalid image');
  assert.strictEqual(saveButton(editor).props.disabled, true);
  editor.validateCustomField(editor.state.draft.uniqueId, 'imageUrl', '');
  assert.strictEqual(saveButton(editor).props.disabled, false);
});

test('save failures show an error and keep the panel and edited values intact', function() {
  const editor = editorFixture().editor;
  editor.props.onChanged = function() { throw new Error('Simulated property update failure'); };
  const errors = [];
  const original = console.error;
  console.error = function() { errors.push(Array.from(arguments)); };
  try {
    editor.save();
  } finally {
    console.error = original;
  }
  assert.strictEqual(editor.state.open, true);
  assert.strictEqual(editor.state.saveError, strings.tileSaveError);
  assert.strictEqual(errors.length, 1);
});

test('the custom property field wires Save into the SharePoint property change callback', function() {
  const part = fixture();
  const properties = setting(part, 'collectionData').properties;
  const field = editorModule.PropertyPaneTileCollection('collectionData', properties);
  const container = document.createElement('div');
  const changes = [];
  field.properties.onRender(container, {}, function(property, value) { changes.push({ property: property, value: value }); });
  const values = [{ title: 'Saved', url: '/saved' }];
  container.rendered.props.onChanged(values);
  assert.strictEqual(changes.length, 1);
  assert.strictEqual(changes[0].property, 'collectionData');
  assert.strictEqual(changes[0].value, values);
});

document.defaultView.close();
console.log(passed + ' passed, ' + failed + ' failed');
process.exitCode = failed ? 1 : 0;
