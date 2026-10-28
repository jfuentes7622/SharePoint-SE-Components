'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

async function runProject(project) {
  const root = path.join(__dirname, '..', project);
  const folder = project === 'ListControl' ? 'listControl' : 'gridControl';
  const base = path.join(root, 'src', 'webparts', folder);
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const document = require(path.join(root, 'node_modules', 'jsdom')).jsdom('<!doctype html><html><body></body></html>');
  let strings;
  vm.runInNewContext(fs.readFileSync(path.join(base, 'loc', 'en-us.js'), 'utf8'), {
    define: function(deps, factory) { strings = factory(); }
  });
  function field(kind) {
    return function(targetProperty, properties) {
      return { kind: kind, targetProperty: targetProperty, properties: properties };
    };
  }
  function Base() {}
  Base.prototype.onPropertyPaneFieldChanged = function(property, oldValue, newValue) {
    this.properties[property] = newValue;
  };
  const imports = {
    react: {}, 'react-dom': {},
    '@microsoft/sp-core-library': {},
    '@microsoft/sp-http': {},
    '@microsoft/sp-webpart-base': {
      BaseClientSideWebPart: Base,
      PropertyPaneButtonType: { Normal: 0, Primary: 1 },
      PropertyPaneButton: field('button'), PropertyPaneCheckbox: field('checkbox'),
      PropertyPaneDropdown: field('dropdown'), PropertyPaneLabel: field('label'),
      PropertyPaneTextField: field('text'), PropertyPaneSlider: field('slider')
    },
    '@microsoft/sp-component-base/lib/DynamicProperty': {},
    '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField': {
      PropertyPaneCustomField: function(properties) { return { kind: 'custom', properties: properties }; }
    },
    '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker': {
      PropertyFieldColorPicker: field('color'), PropertyFieldColorPickerStyle: { Inline: 0 }
    },
    '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData': {
      PropertyFieldCollectionData: field('collection'), CustomCollectionFieldType: {}
    },
    './components/ListControl': {}, './components/GridControl': {},
    './components/ListDesigner': {}, './components/GridDesigner': {},
    '../shared/deterministicFullWidth': {},
    '../shared/overrideCss': { PropertyPaneOverrideCss: field('override') },
    '../../../config/package-solution.json': require(path.join(root, 'config', 'package-solution.json'))
  };
  imports[project + 'WebPartStrings'] = strings;
  const exports = {};
  const loggedErrors = [];
  const source = fs.readFileSync(path.join(base, project + 'WebPart.ts'), 'utf8');
  const compiled = ts.transpileModule(source, {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5 }
  }).outputText;
  vm.runInNewContext(compiled, {
    exports: exports, document: document,
    console: { log: console.log, error: function() { loggedErrors.push(Array.from(arguments)); } },
    require: function(name) {
      assert.ok(Object.prototype.hasOwnProperty.call(imports, name), 'Unexpected import: ' + name);
      return imports[name];
    }
  });
  const part = new exports.default();
  part.properties = { enableDiagnostics: false };
  let refreshes = 0;
  part.context = {
    propertyPane: { refresh: function() { refreshes += 1; } },
    pageContext: { web: { absoluteUrl: 'https://example.test' } }
  };
  part.render = function() {};
  function metadata(verbose) {
    const fields = [
      { InternalName: 'Status', Title: 'Status', TypeAsString: 'Choice', Choices: ['Open', 'Closed', '<Review> & approve'] },
      { InternalName: 'Tags', Title: 'Tags', TypeAsString: 'MultiChoice', Choices: ['First', 'Second'] },
      { InternalName: 'Title', Title: 'Title', TypeAsString: 'Text' },
      { InternalName: 'Empty', Title: 'Empty', TypeAsString: 'Choice', Choices: [] },
      { InternalName: 'Owner', Title: 'Owner', TypeAsString: 'Lookup', LookupList: 'owner-list', LookupField: 'Code' },
      { InternalName: 'Owners', Title: 'Owners', TypeAsString: 'LookupMulti', LookupList: 'owner-list', LookupField: 'Code' },
      { InternalName: 'Person', Title: 'Person', TypeAsString: 'User' },
      { InternalName: 'People', Title: 'People', TypeAsString: 'UserMulti' },
      { InternalName: 'Date', Title: 'Date', TypeAsString: 'DateTime', DisplayFormat: 0 },
      { InternalName: 'Timestamp', Title: 'Timestamp', TypeAsString: 'DateTime', DisplayFormat: 1 },
      { InternalName: 'Amount', Title: 'Amount', TypeAsString: 'Number' },
      { InternalName: 'Cost', Title: 'Cost', TypeAsString: 'Currency' },
      { InternalName: 'Count', Title: 'Count', TypeAsString: 'Integer' },
      { InternalName: 'Enabled', Title: 'Enabled', TypeAsString: 'Boolean' }
    ];
    if (!verbose) { return { value: fields }; }
    return { d: { results: fields.map(function(item) {
      return Object.assign({}, item, item.Choices ? { Choices: { results: item.Choices } } : {});
    }) } };
  }
  function fields() {
    return part.getPropertyPaneConfiguration().pages.reduce(function(all, page) {
      return all.concat.apply(all, page.groups.map(function(group) { return group.groupFields; }));
    }, []);
  }
  function picker(valueProperty) {
    return fields().filter(function(item) {
      return item.kind === 'custom' && item.properties.key === valueProperty + 'TypedValuePicker';
    })[0];
  }
  function renderPicker(valueProperty) {
    const host = document.createElement('div');
    const control = picker(valueProperty);
    assert.ok(control, 'Typed value picker is present');
    control.properties.onRender(host);
    return { host: host, select: host.querySelector('select'), control: control };
  }
  function choose(select, value) {
    select.value = value;
    const event = document.createEvent('HTMLEvents');
    event.initEvent('change', true, true);
    select.dispatchEvent(event);
  }

  for (const verbose of [false, true]) {
    part.getJsonWithAcceptFallback = function() { return Promise.resolve(metadata(verbose)); };
    await part.loadListFields('Items');
    for (const conditional of [false, true]) {
      const fieldProperty = conditional ? 'conditionalStyleConditionField' : 'filterDesignerField';
      const valueProperty = conditional ? 'conditionalStyleConditionValue' : 'filterDesignerValue';
      const typeProperty = conditional ? 'conditionalStyleConditionValueType' : 'filterDesignerValueType';
      part.properties[fieldProperty] = 'Status';
      delete part.properties[typeProperty];
      part.properties[valueProperty] = '';
      const rendered = renderPicker(valueProperty);
      assert.deepStrictEqual(Array.from(rendered.select.options).map(function(option) { return option.textContent; }),
        ['Select a choice', 'Open', 'Closed', '<Review> & approve']);
      assert.strictEqual(rendered.select.value, '', 'No choice is selected silently');
      assert.strictEqual(rendered.host.querySelector('label').contains(rendered.select), true, 'Picker has an accessible label');
      const before = refreshes;
      choose(rendered.select, '<Review> & approve');
      assert.strictEqual(part.properties[valueProperty], '<Review> & approve');
      assert.strictEqual(refreshes, before + 1, 'Selection refreshes the value editor');
      assert.strictEqual(renderPicker(valueProperty).select.value, '<Review> & approve');
      const text = fields().filter(function(item) { return item.kind === 'text' && item.targetProperty === valueProperty; })[0];
      assert.strictEqual(text.properties.value, '<Review> & approve');
      part.properties[valueProperty] = 'partial or legacy value';
      assert.strictEqual(renderPicker(valueProperty).select.value, '', 'Manual values are not replaced');
      assert.strictEqual(part.properties[valueProperty], 'partial or legacy value');
      part.properties[typeProperty] = 'expression';
      assert.strictEqual(picker(valueProperty), undefined, 'Expressions remain text-only');
      part.properties[typeProperty] = 'static';
      ['Title', 'Owner', 'Empty'].forEach(function(name) {
        part.properties[fieldProperty] = name;
        assert.strictEqual(picker(valueProperty), undefined, name + ' has no choice picker');
      });
      part.properties[fieldProperty] = 'Tags';
      assert.strictEqual(renderPicker(valueProperty).select.options.length, 3, 'MultiChoice static conditions can pick one option');
      part.properties[fieldProperty] = 'Status';
      choose(renderPicker(valueProperty).select, 'Closed');
      if (conditional) {
        part.properties.conditionalStyleRuleName = 'Closed items';
        part.properties.conditionalStyleJson = '';
        part.handleAddConditionalStyleCondition();
        const saved = JSON.parse(part.properties.conditionalStyleJson)[0];
        assert.strictEqual(saved.value, 'Closed');
        assert.strictEqual(saved.valueType, 'static');
        assert.strictEqual(saved.conditionField, 'Status');
        part.properties.conditionalStyleSelectedIndex = '0';
        part.handleLoadSelectedConditionalStyleCondition();
      } else {
        part.properties.filterJson = '';
        part.handleAddFilterDesignerCondition();
        const saved = JSON.parse(part.properties.filterJson)[0];
        assert.strictEqual(saved.value, 'Closed');
        assert.strictEqual(saved.valueType, 'static');
        assert.strictEqual(saved.field, 'Status');
        part.properties.filterDesignerSelectedIndex = '0';
        part.handleLoadSelectedFilterCondition();
      }
      assert.strictEqual(renderPicker(valueProperty).select.value, 'Closed', 'Loading saved conditions restores the selection');
      console.log('PASS ' + project + ' ' + (conditional ? 'conditional rules' : 'preset filters') + ' (verbose=' + verbose + ')');
    }
  }

  part.properties.filterDesignerValueType = 'fieldValue';
  part.properties.filterDesignerField = 'Tags';
  part.properties.filterDesignerValue = '["Second"]';
  assert.strictEqual(picker('filterDesignerValue'), undefined, 'No duplicate picker in Field value mode');
  const multi = fields().filter(function(item) {
    return item.kind === 'custom' && item.properties.key === 'filterDesignerFieldValueMultiPicker';
  })[0];
  const host = document.createElement('div');
  multi.properties.onRender(host);
  assert.strictEqual(host.querySelectorAll('input')[1].checked, true, 'Existing multi-selection is retained');
  part.properties.filterJson = '';
  part.handleAddFilterDesignerCondition();
  assert.deepStrictEqual(JSON.parse(part.properties.filterJson)[0].value, ['Second']);
  part.properties.filterDesignerField = 'Status';
  assert.ok(fields().some(function(item) { return item.kind === 'dropdown' && item.targetProperty === 'filterDesignerValue'; }),
    'Existing single-choice Field value dropdown remains available');
  part._fieldChoicesByInternalName.Status = ['Updated option'];
  part.properties.filterDesignerValueType = 'static';
  assert.strictEqual(renderPicker('filterDesignerValue').select.options[1].value, 'Updated option', 'Uses current metadata');
  console.log('PASS ' + project + ' existing Field value modes and metadata updates');

  for (const conditional of [false, true]) {
    const fieldProperty = conditional ? 'conditionalStyleConditionField' : 'filterDesignerField';
    const valueProperty = conditional ? 'conditionalStyleConditionValue' : 'filterDesignerValue';
    const typeProperty = conditional ? 'conditionalStyleConditionValueType' : 'filterDesignerValueType';
    const samples = [
      ['Date', 'date', '2026-10-28'],
      ['Timestamp', 'datetime-local', '2026-10-28T15:45'],
      ['Amount', 'number', '0'],
      ['Cost', 'number', '12.75'],
      ['Count', 'number', '3'],
      ['Enabled', 'select', strings.RuntimeBooleanNo]
    ];
    for (const sample of samples) {
      part.properties[fieldProperty] = sample[0];
      part.properties[valueProperty] = '';
      part.properties[typeProperty] = 'static';
      const rendered = renderPicker(valueProperty);
      const input = rendered.host.querySelector(sample[1] === 'select' ? 'select' : 'input');
      assert.ok(input, sample[0] + ' input is present');
      if (sample[1] !== 'select') { assert.strictEqual(input.getAttribute('type'), sample[1]); }
      choose(input, sample[2]);
      assert.strictEqual(part.properties[valueProperty], sample[2]);
      if (conditional) {
        part.properties.conditionalStyleRuleName = 'Typed value';
        part.properties.conditionalStyleJson = '';
        part.handleAddConditionalStyleCondition();
        assert.strictEqual(JSON.parse(part.properties.conditionalStyleJson)[0].value, sample[2]);
        part.properties.conditionalStyleSelectedIndex = '0';
        part.handleLoadSelectedConditionalStyleCondition();
      } else {
        part.properties.filterJson = '';
        part.handleAddFilterDesignerCondition();
        assert.strictEqual(JSON.parse(part.properties.filterJson)[0].value, sample[2]);
        part.properties.filterDesignerSelectedIndex = '0';
        part.handleLoadSelectedFilterCondition();
      }
      assert.strictEqual(renderPicker(valueProperty).host.querySelector('input, select').value, sample[2], 'Saved typed value reopens');
      part.properties[typeProperty] = 'expression';
      assert.strictEqual(picker(valueProperty), undefined, 'Typed picker is hidden for expressions');
      console.log('PASS ' + project + ' ' + (conditional ? 'rules' : 'filters') + ' ' + sample[0]);
    }
  }

  const requests = [];
  part.getJsonWithAcceptFallback = function(url) {
    requests.push(url);
    if (url.indexOf('/siteusers?') >= 0) {
      return Promise.resolve({ value: [
        { Id: 7, Title: 'Person One', PrincipalType: 1 },
        { Id: 8, Title: 'Site Group', PrincipalType: 8 },
        { Id: 9, Title: 'Invalid principal', PrincipalType: 0 }
      ] });
    }
    if (url === 'https://example.test/next') {
      return Promise.resolve({ d: { results: [{ Id: 12, Code: 'Second code', Title: 'Not the configured lookup field' }] } });
    }
    assert.ok(decodeURIComponent(url).indexOf("lists/getByTitle('owner-list')/items?$select=Id,Code") >= 0, url);
    return Promise.resolve({ value: [{ Id: 11, Code: 'First code' }], 'odata.nextLink': 'https://example.test/next' });
  };
  for (const conditional of [false, true]) {
    const fieldProperty = conditional ? 'conditionalStyleConditionField' : 'filterDesignerField';
    const valueProperty = conditional ? 'conditionalStyleConditionValue' : 'filterDesignerValue';
    const typeProperty = conditional ? 'conditionalStyleConditionValueType' : 'filterDesignerValueType';
    const pickProperty = conditional ? 'conditionalStyleLookupPick' : 'filterDesignerLookupPick';
    const loadValues = conditional ? 'handleLoadConditionalStyleLookupItems' : 'handleLoadFilterLookupItems';
    const optionsProperty = conditional ? '_conditionalStyleLookupItemOptions' : '_filterLookupItemOptions';
    const messageProperty = conditional ? '_conditionalStyleLookupMessage' : '_filterLookupMessage';
    part.properties[typeProperty] = 'static';
    for (const name of ['Person', 'People', 'Owner', 'Owners']) {
      part.properties[fieldProperty] = name;
      await part[loadValues]();
      const dropdown = fields().filter(function(item) { return item.targetProperty === pickProperty; })[0];
      assert.ok(dropdown, name + ' picker is available without changing value type');
      assert.strictEqual(dropdown.properties.options.length, 2);
      assert.ok(dropdown.properties.options[0].text.indexOf(name.indexOf('Owner') === 0 ? 'code' : 'Person') >= 0);
      const id = String(dropdown.properties.options[0].key);
      part.onPropertyPaneFieldChanged(pickProperty, '', id);
      assert.strictEqual(part.properties[valueProperty], id, 'Selecting an option immediately writes its ID');
      assert.strictEqual(fields().filter(function(item) { return item.targetProperty === valueProperty && item.kind === 'text'; })[0].properties.value, id);
      if (conditional) {
        part.properties.conditionalStyleRuleName = 'Reference';
        assert.strictEqual(part.createConditionalStyleEntryFromDesigner().value, id);
      } else {
        part.properties.filterJson = '';
        part.handleAddFilterDesignerCondition();
        assert.strictEqual(JSON.parse(part.properties.filterJson)[0].value, id);
      }
      part.properties[typeProperty] = 'expression';
      assert.strictEqual(fields().some(function(item) { return item.targetProperty === pickProperty; }), false);
      part.properties[typeProperty] = 'static';
    }
    part.properties[fieldProperty] = 'Person';
    const originalLoader = part.getJsonWithAcceptFallback;
    let resolveOld;
    part.getJsonWithAcceptFallback = function() { return new Promise(function(resolve) { resolveOld = resolve; }); };
    const pending = part[loadValues]();
    part.properties[fieldProperty] = 'Title';
    resolveOld({ value: [{ Id: 99, Title: 'Stale person', PrincipalType: 1 }] });
    await pending;
    assert.strictEqual(part[optionsProperty].length, 0, 'Stale requests do not fill a different field');
    part.properties[fieldProperty] = 'Person';
    part.getJsonWithAcceptFallback = function() { return Promise.reject(new Error('HTTP 403 Forbidden')); };
    await part[loadValues]();
    assert.ok(part[messageProperty].indexOf('HTTP 403') >= 0, 'Failure is visible, not reported as an empty list');
    assert.strictEqual(loggedErrors.length, conditional ? 2 : 1, 'Failures are logged even with diagnostics off');
    part.getJsonWithAcceptFallback = originalLoader;
    console.log('PASS ' + project + ' ' + (conditional ? 'rules' : 'filters') + ' people, lookups, pagination, stale responses, errors');
  }
  part.properties.filterDesignerField = 'People';
  part.properties.filterDesignerValueType = 'fieldValue';
  part.properties.filterDesignerValue = '["7"]';
  await part.handleLoadFilterLookupItems();
  const peopleMulti = fields().filter(function(item) {
    return item.kind === 'custom' && item.properties.key === 'filterDesignerFieldValueMultiPicker';
  })[0];
  const peopleHost = document.createElement('div');
  peopleMulti.properties.onRender(peopleHost);
  assert.strictEqual(peopleHost.querySelector('input').checked, true);
  assert.strictEqual(fields().some(function(item) { return item.targetProperty === 'filterDesignerLookupPick'; }), false);
  part.properties.filterJson = '';
  part.handleAddFilterDesignerCondition();
  assert.deepStrictEqual(JSON.parse(part.properties.filterJson)[0].value, ['7']);
  const requestCount = requests.length;
  part.onPropertyPaneFieldChanged('filterDesignerField', 'People', 'Person');
  await new Promise(setImmediate);
  assert.ok(requests.length > requestCount, 'Selecting a reference column loads its values automatically');
  assert.strictEqual(part._filterLookupItemOptions.length, 2);
  console.log('PASS ' + project + ' multi-person Field value and automatic loading');

  imports.react = require(path.join(root, 'node_modules', 'react'));
  imports['./DateTime24HourInput'] = {};
  imports['./GridValidation'] = {};
  imports['./GridRichTextEditor'] = {};
  imports['./' + project + '.css'] = {};
  const runtimeExports = {};
  const runtimeSource = fs.readFileSync(path.join(base, 'components', project + '.tsx'), 'utf8');
  vm.runInNewContext(ts.transpileModule(runtimeSource, {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
  }).outputText, {
    exports: runtimeExports, document: document, window: document.defaultView, console: console,
    require: function(name) {
      assert.ok(Object.prototype.hasOwnProperty.call(imports, name), 'Unexpected runtime import: ' + name);
      return imports[name];
    }
  });
  const Control = runtimeExports[project];
  const runtime = new Control({
    views: [], viewColumns: [], gridSchemaJson: '', defaultViewId: '', listName: 'Items',
    enableDiagnostics: false, context: part.context
  });
  [
    ['Person', 'User', [{ lookupId: '7', lookupValue: 'Person One' }], '7'],
    ['Person', 'User', [{ lookupId: '7', lookupValue: 'Person One' }], 'Person One'],
    ['People', 'UserMulti', [{ lookupId: '7', lookupValue: 'Person One' }, { lookupId: '8', lookupValue: 'Site Group' }], '8'],
    ['Person', 'User', '[{"lookupId":"7","lookupValue":"Person One"}]', '7'],
    ['Owner', 'Lookup', { LookupId: 11, LookupValue: 'First code' }, '11'],
    ['Owner', 'Lookup', { LookupId: 11, LookupValue: 'First code' }, 'First code'],
    ['Owners', 'LookupMulti', [{ LookupId: 11, LookupValue: 'First code' }, { LookupId: 12, LookupValue: 'Second code' }], '12'],
    ['Tags', 'MultiChoice', ['First', 'Second'], 'Second'],
    ['Amount', 'Number', 12.75, '12.750'],
    ['Cost', 'Currency', 12, '12.00'],
    ['Count', 'Integer', 0, '0'],
    ['Enabled', 'Boolean', false, strings.RuntimeBooleanNo]
  ].forEach(function(sample) {
    const definition = { Name: sample[0], TypeAsString: sample[1] };
    const definitions = {};
    definitions[sample[0]] = definition;
    const row = {};
    row[sample[0]] = sample[2];
    const condition = { field: sample[0], operator: 'eq', logical: 'and', valueType: 'static', value: sample[3] };
    assert.strictEqual(runtime.rowMatchesPresetConditions(row, definitions, [condition]), true, sample[0] + ' preset rule');
    assert.strictEqual(runtime.evaluateConditionalStyleRule(row, { conditions: [condition] }, definitions), true, sample[0] + ' style rule');
    condition.operator = 'ne';
    assert.strictEqual(runtime.rowMatchesPresetConditions(row, definitions, [condition]), false, sample[0] + ' negative rule');
    console.log('PASS ' + project + ' runtime ' + sample[1] + ' comparison');
  });
}

(async function() {
  await runProject('ListControl');
  await runProject('GridControl');
  console.log('All type-aware rule value designer checks passed.');
})().catch(function(error) {
  console.error(error.stack || error);
  process.exitCode = 1;
});
