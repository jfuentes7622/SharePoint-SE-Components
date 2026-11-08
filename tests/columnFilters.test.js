'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const tests = [];

['ListControl', 'GridControl'].forEach(function(project) {
  const root = path.join(__dirname, '..', project);
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const React = require(path.join(root, 'node_modules', 'react'));
  const jsdom = require(path.join(root, 'node_modules', 'jsdom'));
  const document = jsdom.jsdom('<!doctype html><html><body><button id="anchor"></button></body></html>');
  const directory = path.join(root, 'src', 'webparts', project === 'ListControl' ? 'listControl' : 'gridControl');
  let strings;
  vm.runInNewContext(fs.readFileSync(path.join(directory, 'loc', 'en-us.js'), 'utf8'), {
    define: function(dependencies, factory) { strings = factory(); }
  });
  const errors = [];
  const exports = {};
  const output = ts.transpileModule(fs.readFileSync(path.join(directory, 'components', project + '.tsx'), 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
  }).outputText;
  vm.runInNewContext(output, {
    exports: exports, document: document, window: document.defaultView,
    console: { log: console.log, error: function() { errors.push(Array.from(arguments)); } },
    require: function(name) {
      if (name === './LibraryBrowser') { return require('./loadLibraryBrowser')(root, document.defaultView); }
      if (name === 'react') { return React; }
      if (name === '@microsoft/sp-http') { return { SPHttpClient: { configurations: { v1: {} } } }; }
      if (name === project + 'WebPartStrings') { return strings; }
      if (/\.css$/.test(name) || name === './GridRichTextEditor' || name === './GridValidation'
        || name === './DateTime24HourInput') { return {}; }
      throw new Error('Unexpected import: ' + name);
    }
  });
  const Control = exports[project];
  function test(name, run) { tests.push({ name: project + ': ' + name, run: run }); }
  function plain(value) { return JSON.parse(JSON.stringify(value)); }
  function settle() { return new Promise(function(resolve) { setImmediate(resolve); }); }
  function response(data, status) {
    return { ok: !status || status === 200, status: status || 200, statusText: status ? 'Forbidden' : 'OK',
      json: function() { return Promise.resolve(data); } };
  }
  function field(type) {
    return { Name: 'Category', DisplayName: 'Category', TypeAsString: type || 'Choice',
      Choices: ['Red', 'Blue', 'Green'], LookupList: "Owners' list", LookupField: 'Code' };
  }
  function control(definition) {
    const instance = new Control({
      views: [], viewColumns: [], gridSchemaJson: '', defaultViewId: '', listName: 'Items',
      enableDiagnostics: false, isEditMode: false, pageSize: 50, fetchBatchSize: 500,
      timeDisplayFormat: '24hour', timeMinuteIncrement: 5, dateDisplayFormat: 'mdy',
      context: { pageContext: { web: { absoluteUrl: 'https://example.test/site' } } },
      onFilteredCountChange: function() {}, onSelectionChange: function() {}
    });
    instance.setState = function(update) {
      this.state = Object.assign({}, this.state, typeof update === 'function' ? update(this.state, this.props) : update);
    };
    Object.assign(instance.state, {
      loading: false, accessLoading: false, hasReadAccess: true, hasEditAccess: true,
      fields: [definition], rows: [{ ID: 1, Category: 'Red' }, { ID: 2, Category: 'Blue' }, { ID: 3, Category: 'Green' }]
    });
    instance.state.fieldMetadataByName = {};
    instance.state.fieldMetadataByName[String(definition.RealFieldName || definition.Name).toLowerCase()] = {
      internalName: definition.RealFieldName || definition.Name, typeAsString: definition.TypeAsString,
      choices: definition.Choices || [], lookupList: definition.LookupList, lookupField: definition.LookupField
    };
    instance.getJsonWithFallback = function() {
      return Promise.resolve(response({ value: [{ Id: 7, Code: 'Same label' }, { Id: 8, Code: 'Same label' }, { Id: 17, Code: 'Other' }] }));
    };
    return instance;
  }
  async function open(instance, definition) {
    instance.openFilter(definition, document.getElementById('anchor'));
    await settle();
  }
  function elements(element, predicate, result) {
    result = result || [];
    if (!element || !element.props) { return result; }
    if (predicate(element)) { result.push(element); }
    React.Children.forEach(element.props.children, function(child) { elements(child, predicate, result); });
    return result;
  }
  function popup(instance) {
    const matches = elements(instance.render(), function(element) { return element.props.className === 'lc-filter-popover'; });
    assert.strictEqual(matches.length, 1, 'Expected one active column popup');
    return matches[0];
  }
  function options(instance) {
    return elements(popup(instance), function(element) { return element.props.className === 'lc-filter-option'; });
  }
  function button(instance, label) {
    return elements(popup(instance), function(element) { return element.type === 'button' && element.props['aria-label'] === label; })[0];
  }
  function clickOption(instance, index, checked) {
    const checkbox = elements(options(instance)[index], function(element) { return element.type === 'input'; })[0];
    checkbox.props.onChange({ currentTarget: { checked: checked } });
  }

  ['Choice', 'MultiChoice'].forEach(function(type) {
    test(type + ' renders labeled checkbox options, commits arrays and matches any selected value', async function() {
      const definition = field(type);
      const instance = control(definition);
      await open(instance, definition);
      assert.strictEqual(options(instance).length, 3);
      options(instance).forEach(function(option) {
        assert.strictEqual(option.type, 'label');
        const input = elements(option, function(element) { return element.type === 'input'; })[0];
        assert.strictEqual(input.props.type, 'checkbox');
        assert.strictEqual(input.props.checked, false);
      });
      assert.deepStrictEqual(plain(instance.getFilterOperatorOptions(definition).map(function(option) { return option.key; })), ['eq', 'ne']);
      clickOption(instance, 0, true);
      clickOption(instance, 1, true);
      assert.strictEqual(options(instance)[0].props.children[0].props.checked, true);
      assert.strictEqual(instance.state.columnFilters.Category, undefined, 'Changes remain drafts until Apply');
      instance.state.currentPage = 4;
      button(instance, strings.RuntimeFilterApply).props.onClick();
      assert.deepStrictEqual(plain(instance.state.columnFilters.Category), { operator: 'eq', value: ['Red', 'Blue'] });
      assert.deepStrictEqual(plain(instance.getProcessedRows().map(function(row) { return row.ID; })), [1, 2]);
      assert.strictEqual(instance.state.currentPage, 0);
      assert.strictEqual(instance.state.activeFilterFieldName, '');
    });
  });

  test('Not equals excludes every checked value, including multi-valued cells', async function() {
    const definition = field('MultiChoice');
    const instance = control(definition);
    await open(instance, definition);
    instance.state.draftFilterOperator = 'ne';
    clickOption(instance, 0, true);
    clickOption(instance, 1, true);
    instance.applyActiveFilter();
    const filter = instance.state.columnFilters.Category;
    ['Red', 'Blue', ['Green', 'Blue'], ';#Red;#Green;#'].forEach(function(value) {
      assert.strictEqual(instance.rowMatchesFilter({ Category: value }, definition, filter), false, JSON.stringify(value));
    });
    ['Green', '', ['Green']].forEach(function(value) {
      assert.strictEqual(instance.rowMatchesFilter({ Category: value }, definition, filter), true, JSON.stringify(value));
    });
  });

  test('multi-choice arrays and SharePoint-delimited values match discrete whole choices', async function() {
    const definition = field('MultiChoice');
    const instance = control(definition);
    await open(instance, definition);
    instance.toggleFilterValue('Red', true);
    instance.toggleFilterValue('Blue', true);
    instance.applyActiveFilter();
    const filter = instance.state.columnFilters.Category;
    [['Green', 'Blue'], { results: ['Green', 'Blue'] }, ';#Red;#Green;#', 'Green; Blue'].forEach(function(value) {
      assert.strictEqual(instance.rowMatchesFilter({ Category: value }, definition, filter), true, JSON.stringify(value));
    });
    assert.strictEqual(instance.rowMatchesFilter({ Category: ['Redwood', 'Green'] }, definition, filter), false);
    assert.strictEqual(instance.rowMatchesFilter({ Category: ['Red; Blue'] }, definition, filter), false);
    assert.strictEqual(instance.rowMatchesFilter({ Category: ['Red; Blue'] }, definition, { operator: 'eq', value: ['Red; Blue'] }), true);
  });

  ['Lookup', 'LookupMulti'].forEach(function(type) {
    test(type + ' selects by ID even with duplicate labels and matches multiple selections', async function() {
      const definition = field(type);
      const instance = control(definition);
      await open(instance, definition);
      assert.deepStrictEqual(plain(instance.state.activeFilterOptions.map(function(option) { return option.value; })).sort(), ['17', '7', '8']);
      assert.strictEqual(instance.state.activeFilterOptions.filter(function(option) { return option.text === 'Same label'; }).length, 2);
      instance.toggleFilterValue('7', true);
      instance.applyActiveFilter();
      let filter = instance.state.columnFilters.Category;
      assert.strictEqual(instance.rowMatchesFilter({ Category: [{ lookupId: '7', lookupValue: 'Same label' }] }, definition, filter), true);
      assert.strictEqual(instance.rowMatchesFilter({ Category: [{ lookupId: '8', lookupValue: 'Same label' }] }, definition, filter), false);
      assert.strictEqual(instance.rowMatchesFilter({ Category: [{ lookupId: '17', lookupValue: 'Other' }] }, definition, filter), false);
      await open(instance, definition);
      instance.toggleFilterValue('8', true);
      instance.applyActiveFilter();
      filter = instance.state.columnFilters.Category;
      assert.strictEqual(instance.rowMatchesFilter({ Category: [{ lookupId: '8', lookupValue: 'Same label' }] }, definition, filter), true);
      assert.strictEqual(instance.rowMatchesFilter({ CategoryId: [17, 8] }, definition, filter), true);
      filter.operator = 'ne';
      assert.strictEqual(instance.rowMatchesFilter({ CategoryId: [17, 8] }, definition, filter), false);
      assert.strictEqual(instance.rowMatchesFilter({ CategoryId: [17] }, definition, filter), true);
    });
  });

  test('multiple column filters remain ANDed with OR selections within one column', async function() {
    const definition = field();
    const instance = control(definition);
    instance.state.fields.push({ Name: 'Title', TypeAsString: 'Text' });
    instance.state.rows = [{ ID: 1, Category: 'Red', Title: 'Keep' }, { ID: 2, Category: 'Blue', Title: 'Drop' },
      { ID: 3, Category: 'Green', Title: 'Keep' }, { ID: 4, Category: 'Blue', Title: 'Keep' }];
    instance.state.columnFilters.Title = { operator: 'eq', value: 'Keep' };
    await open(instance, definition);
    instance.toggleFilterValue('Red', true);
    instance.toggleFilterValue('Blue', true);
    instance.applyActiveFilter();
    assert.deepStrictEqual(plain(instance.getProcessedRows().map(function(row) { return row.ID; })), [1, 4]);
  });

  test('reopening restores checks; Close discards drafts; scalar filters stay compatible', async function() {
    const definition = field();
    const instance = control(definition);
    instance.state.columnFilters.Category = { operator: 'ne', value: 'Red' };
    await open(instance, definition);
    assert.deepStrictEqual(plain(instance.state.draftFilterValues), ['Red']);
    assert.strictEqual(instance.state.draftFilterOperator, 'ne');
    assert.strictEqual(options(instance)[0].props.children[0].props.checked, true);
    clickOption(instance, 1, true);
    button(instance, strings.RuntimeFilterClose).props.onClick();
    assert.strictEqual(instance.state.columnFilters.Category.value, 'Red');
    await open(instance, definition);
    assert.deepStrictEqual(plain(instance.state.draftFilterValues), ['Red']);
    clickOption(instance, 1, true);
    instance.applyActiveFilter();
    await open(instance, definition);
    assert.deepStrictEqual(plain(instance.state.draftFilterValues), ['Red', 'Blue']);
  });

  test('unchecking everything and Clear both remove only the active filter and reset paging', async function() {
    const definition = field();
    const instance = control(definition);
    instance.state.columnFilters = { Category: { operator: 'eq', value: ['Red'] }, Title: { operator: 'contains', value: 'Keep' } };
    await open(instance, definition);
    clickOption(instance, 0, false);
    instance.state.currentPage = 3;
    instance.applyActiveFilter();
    assert.deepStrictEqual(plain(instance.state.columnFilters), { Title: { operator: 'contains', value: 'Keep' } });
    assert.strictEqual(instance.state.currentPage, 0);
    instance.state.columnFilters.Category = { operator: 'eq', value: ['Blue'] };
    await open(instance, definition);
    button(instance, strings.RuntimeFilterClear).props.onClick();
    assert.strictEqual(instance.state.columnFilters.Category, undefined);
    assert.strictEqual(instance.state.columnFilters.Title.value, 'Keep');
  });

  test('an applied selection removed from metadata remains visible and can be unchecked', async function() {
    const definition = field();
    const instance = control(definition);
    instance.state.columnFilters.Category = { operator: 'eq', value: ['Removed choice'] };
    await open(instance, definition);
    assert.strictEqual(options(instance).length, 4);
    assert.strictEqual(options(instance)[3].props.children[0].props.checked, true);
    clickOption(instance, 3, false);
    instance.applyActiveFilter();
    assert.strictEqual(instance.state.columnFilters.Category, undefined);
  });

  test('an empty choice list shows an explicit empty state instead of a text editor', async function() {
    const definition = field();
    definition.Choices = [];
    const instance = control(definition);
    await open(instance, definition);
    assert.strictEqual(options(instance).length, 0);
    assert.strictEqual(elements(popup(instance), function(element) { return element.type === 'input'; }).length, 0);
    assert.strictEqual(elements(popup(instance), function(element) { return element.props.children === strings.RuntimeFilterNoOptions; }).length, 1);
    instance.applyActiveFilter();
    assert.strictEqual(instance.state.columnFilters.Category, undefined);
  });

  test('lookup loading disables Apply without removing the current filter', async function() {
    const definition = field('Lookup');
    const instance = control(definition);
    let finish;
    instance.getJsonWithFallback = function() { return new Promise(function(resolve) { finish = resolve; }); };
    instance.state.columnFilters.Category = { operator: 'eq', value: ['7'] };
    instance.openFilter(definition, document.getElementById('anchor'));
    assert.strictEqual(button(instance, strings.RuntimeFilterApply).props.disabled, true);
    assert.strictEqual(elements(popup(instance), function(element) { return element.props.role === 'status'; }).length, 1);
    instance.applyActiveFilter();
    assert.strictEqual(instance.state.activeFilterFieldName, 'Category');
    assert.deepStrictEqual(plain(instance.state.columnFilters.Category.value), ['7']);
    finish(response({ value: [{ Id: 7, Code: 'Owner' }] }));
    await settle();
    assert.strictEqual(button(instance, strings.RuntimeFilterApply).props.disabled, false);
  });

  test('lookup HTTP failures are displayed and logged; Clear remains available', async function() {
    const definition = field('Lookup');
    const instance = control(definition);
    const before = errors.length;
    instance.getJsonWithFallback = function() { return Promise.resolve(response({}, 403)); };
    instance.state.columnFilters.Category = { operator: 'eq', value: ['7'] };
    await open(instance, definition);
    assert.ok(instance.state.activeFilterOptionsError.indexOf('403') >= 0);
    assert.strictEqual(errors.length, before + 1);
    assert.strictEqual(elements(popup(instance), function(element) { return element.props.role === 'alert'; }).length, 1);
    assert.strictEqual(button(instance, strings.RuntimeFilterApply).props.disabled, true);
    instance.applyActiveFilter();
    assert.deepStrictEqual(plain(instance.state.columnFilters.Category.value), ['7']);
    button(instance, strings.RuntimeFilterClear).props.onClick();
    assert.strictEqual(instance.state.columnFilters.Category, undefined);
  });

  test('older lookup responses cannot overwrite a closed and reopened filter', async function() {
    const definition = field('Lookup');
    const instance = control(definition);
    const pending = [];
    instance.getJsonWithFallback = function() { return new Promise(function(resolve) { pending.push(resolve); }); };
    instance.openFilter(definition, document.getElementById('anchor'));
    instance.closeFilter();
    instance.openFilter(definition, document.getElementById('anchor'));
    pending[1](response({ value: [{ Id: 8, Code: 'Current' }] }));
    await settle();
    pending[0](response({ value: [{ Id: 7, Code: 'Outdated' }] }));
    await settle();
    assert.deepStrictEqual(plain(instance.state.activeFilterOptions), [{ value: '8', text: 'Current' }]);
  });

  ['@odata.nextLink', 'odata.nextLink', '__next'].forEach(function(nextProperty) {
    test('lookup options include paged ' + nextProperty + ' results and escape list titles', async function() {
      const definition = field('Lookup');
      const instance = control(definition);
      const urls = [];
      instance.getJsonWithFallback = function(url) {
        urls.push(url);
        if (urls.length > 1) { return Promise.resolve(response({ d: { results: [{ Id: 8, Code: 'Second' }] } })); }
        const data = nextProperty === '__next' ? { d: { results: [{ Id: 7, Code: 'First' }], __next: 'https://example.test/site/next' } }
          : { value: [{ Id: 7, Code: 'First' }] };
        if (nextProperty !== '__next') { data[nextProperty] = 'https://example.test/site/next'; }
        return Promise.resolve(response(data));
      };
      await open(instance, definition);
      assert.strictEqual(urls.length, 2);
      assert.ok(urls[0].indexOf("getByTitle('Owners'' list')") >= 0);
      assert.ok(urls[0].indexOf('Id%2CCode') >= 0);
      assert.deepStrictEqual(plain(instance.state.activeFilterOptions.map(function(option) { return option.value; })), ['7', '8']);
    });
  });

  test('reference options resolve real internal names for aliased columns', async function() {
    const definition = field('Lookup');
    definition.RealFieldName = 'Owner';
    const instance = control(definition);
    await open(instance, definition);
    assert.strictEqual(instance.state.activeFilterOptionsError, '');
    assert.strictEqual(options(instance).length, 3);
  });

  ['Text', 'Number', 'Boolean', 'User'].forEach(function(type) {
    test(type + ' retains single-value filter behavior', async function() {
      const definition = field(type);
      const instance = control(definition);
      await open(instance, definition);
      assert.strictEqual(instance.state.activeFilterIsMultiSelect, false);
      assert.strictEqual(options(instance).length, 0);
      instance.state.draftFilterValue = type === 'Number' ? '12' : type === 'User' ? '7' : 'Red';
      instance.applyActiveFilter();
      assert.strictEqual(typeof instance.state.columnFilters.Category.value, 'string');
    });
  });

  test('checkbox options use a bounded scroll area and do not become sticky-header proxy controls', function() {
    const css = fs.readFileSync(path.join(directory, 'components', project + '.css'), 'utf8');
    assert.ok(/\.lc-filter-options\s*\{[^}]*max-height:\s*180px;[^}]*overflow-y:\s*auto;/.test(css));
    const head = document.createElement('thead');
    head.innerHTML = '<tr><th><button>Filter</button><div class="lc-filter-popover"><label><input type="checkbox" /></label></div></th></tr>';
    assert.strictEqual(control(field()).getStickyHeaderControls(head).length, 1);
  });

  if (project === 'GridControl') {
    test('missing runtime metadata is an explicit filter-options failure', async function() {
      const definition = field();
      const instance = control(definition);
      instance.state.fieldMetadataByName = {};
      await open(instance, definition);
      assert.strictEqual(instance.state.activeFilterOptionsError, strings.RuntimeFilterOptionsLoadFailed);
      assert.strictEqual(button(instance, strings.RuntimeFilterApply).props.disabled, true);
    });
  }
});

(async function() {
  let failed = 0;
  for (const test of tests) {
    try {
      await test.run();
      console.log('PASS ' + test.name);
    } catch (error) {
      failed += 1;
      console.error('FAIL ' + test.name + '\n' + error.stack);
    }
  }
  console.log((tests.length - failed) + ' passed, ' + failed + ' failed');
  process.exitCode = failed ? 1 : 0;
})();
