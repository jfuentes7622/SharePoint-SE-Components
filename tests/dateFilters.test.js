'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

let passed = 0;
let failed = 0;

['ListControl', 'GridControl'].forEach(function(project) {
  const root = path.join(__dirname, '..', project);
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const React = require(path.join(root, 'node_modules', 'react'));
  const jsdom = require(path.join(root, 'node_modules', 'jsdom'));
  const document = jsdom.jsdom('<!doctype html><html><body><button id="anchor"></button></body></html>');
  const window = document.defaultView;
  const webpart = project === 'ListControl' ? 'listControl' : 'gridControl';
  const directory = path.join(root, 'src', 'webparts', webpart);
  const stylesheet = document.createElement('style');
  stylesheet.textContent = fs.readFileSync(path.join(directory, 'components', project + '.css'), 'utf8');
  document.head.appendChild(stylesheet);
  let strings;
  vm.runInNewContext(fs.readFileSync(path.join(directory, 'loc', 'en-us.js'), 'utf8'), {
    define: function(dependencies, factory) { strings = factory(); }
  });
  function load(name) {
    const exports = {};
    const output = ts.transpileModule(fs.readFileSync(path.join(directory, 'components', name + '.tsx'), 'utf8'), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
    }).outputText;
    vm.runInNewContext(output, {
      exports: exports, window: window, document: document, console: console,
      require: function(module) {
        if (module === 'react') { return React; }
        if (module === '@microsoft/sp-http') { return { SPHttpClient: { configurations: { v1: {} } } }; }
        if (module === project + 'WebPartStrings') { return strings; }
        if (module === './DateTime24HourInput') { return load('DateTime24HourInput'); }
        if (module === './GridRichTextEditor' || module === './GridValidation' || /\.css$/.test(module)) { return {}; }
        throw new Error('Unexpected import: ' + module);
      }
    });
    return exports;
  }
  const Control = load(project)[project];
  function control(displayFormat) {
    const instance = new Control({
      views: [], viewColumns: [], gridSchemaJson: '', defaultViewId: '', listName: 'Items',
      enableDiagnostics: false, isEditMode: false, pageSize: 50, fetchBatchSize: 500,
      timeDisplayFormat: '24hour', timeMinuteIncrement: 5, dateDisplayFormat: 'mdy',
      context: { pageContext: { web: { absoluteUrl: 'https://example.test' } } },
      onFilteredCountChange: function() {}, onSelectionChange: function() {}
    });
    instance.setState = function(update) { this.state = Object.assign({}, this.state, update); };
    instance.state.accessLoading = false;
    instance.state.hasReadAccess = true;
    instance.state.hasEditAccess = true;
    instance.state.loading = false;
    const field = { Name: 'EventDate', TypeAsString: 'DateTime', DisplayFormat: displayFormat, DisplayName: 'Event date' };
    instance.state.fields = [field];
    instance.state.rows = [{ ID: 1, EventDate: '2026-10-22T12:30:00' }];
    instance.openFilter(field, document.getElementById('anchor'));
    return { instance: instance, field: field };
  }
  function test(name, run) {
    try {
      run();
      passed += 1;
      console.log('PASS ' + project + ': ' + name);
    } catch (error) {
      failed += 1;
      console.error('FAIL ' + project + ': ' + name + ': ' + error.message);
    }
  }
  function apply(instance, start, end) {
    instance.state.draftFilterValue = start;
    instance.state.draftFilterEndValue = end || '';
    instance.applyActiveFilter();
    return instance.state.columnFilters.EventDate;
  }
  function findElements(element, predicate, matches) {
    matches = matches || [];
    if (!element || !element.props) { return matches; }
    if (predicate(element)) { matches.push(element); }
    React.Children.forEach(element.props.children, function(child) { findElements(child, predicate, matches); });
    return matches;
  }
  function styleFor(selector) {
    const rules = stylesheet.sheet.cssRules;
    for (let index = 0; index < rules.length; index += 1) {
      if (rules[index].selectorText && rules[index].selectorText.replace(/\s+/g, ' ') === selector.replace(/\s+/g, ' ')) {
        return rules[index].style;
      }
    }
    assert.fail('Missing style rule: ' + selector);
  }

  [0, 1].forEach(function(displayFormat) {
    test('a single date matches the entire day (DisplayFormat ' + displayFormat + ')', function() {
      const page = control(displayFormat);
      const filter = apply(page.instance, '2026-10-22');
      assert.strictEqual(filter.compareDateOnly, true);
      ['2026-10-22T00:00:00', '2026-10-22T12:30:00', '2026-10-22T23:59:59.999'].forEach(function(value) {
        assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: value }, page.field, filter), true, value);
      });
      ['2026-10-21T23:59:59.999', '2026-10-23T00:00:00', '', 'not-a-date'].forEach(function(value) {
        assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: value }, page.field, filter), false, value);
      });
    });
  });

  test('a plain ISO date is not shifted to the previous day by UTC parsing', function() {
    const page = control(0);
    const filter = apply(page.instance, '2026-10-22');
    assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: '2026-10-22' }, page.field, filter), true);
    assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: '2026-10-23' }, page.field, filter), false);
  });

  test('one date works on a daylight-saving transition day', function() {
    const page = control(1);
    const filter = apply(page.instance, '2026-03-08');
    ['2026-03-08T00:00:00', '2026-03-08T03:30:00', '2026-03-08T23:59:59'].forEach(function(value) {
      assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: value }, page.field, filter), true);
    });
    assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: '2026-03-09T00:00:00' }, page.field, filter), false);
  });

  test('UTC timestamps match their corresponding local calendar day', function() {
    const page = control(1);
    const filter = apply(page.instance, '2026-10-22');
    const onDate = new Date(2026, 9, 22, 0, 0, 0).toISOString();
    const lateOnDate = new Date(2026, 9, 22, 23, 59, 59).toISOString();
    const nextDate = new Date(2026, 9, 23, 0, 0, 0).toISOString();
    assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: onDate }, page.field, filter), true);
    assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: lateOnDate }, page.field, filter), true);
    assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: nextDate }, page.field, filter), false);
  });

  test('an optional end date preserves inclusive ranges and reversed-endpoint normalization', function() {
    const page = control(1);
    const filter = apply(page.instance, '2026-10-24', '2026-10-22');
    assert.strictEqual(filter.value, '2026-10-22');
    assert.strictEqual(filter.endValue, '2026-10-24');
    ['2026-10-22T00:00:00', '2026-10-23T12:00:00', '2026-10-24T23:59:59'].forEach(function(value) {
      assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: value }, page.field, filter), true);
    });
    assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: '2026-10-25T00:00:00' }, page.field, filter), false);
  });

  test('applying a single date preserves other column filters, resets paging and can be cleared', function() {
    const page = control(0);
    page.instance.state.columnFilters.Title = { operator: 'eq', value: 'Keep' };
    page.instance.state.currentPage = 3;
    apply(page.instance, '2026-10-22');
    assert.strictEqual(page.instance.state.columnFilters.Title.value, 'Keep');
    assert.strictEqual(page.instance.state.currentPage, 0);
    page.instance.openFilter(page.field, document.getElementById('anchor'));
    assert.strictEqual(page.instance.state.draftFilterValue, '2026-10-22');
    assert.strictEqual(page.instance.state.draftFilterEndValue, '');
    page.instance.clearActiveFilter();
    assert.strictEqual(page.instance.state.columnFilters.EventDate, undefined);
    assert.strictEqual(page.instance.state.columnFilters.Title.value, 'Keep');
  });

  test('date input clicks retain native default behavior without triggering header sorting', function() {
    const page = control(0);
    const popover = findElements(page.instance.render(), function(element) {
      return element.props.className === 'lc-filter-popover';
    })[0];
    assert.ok(popover, 'Date filter popover must render');
    let prevented = false;
    let stopped = false;
    popover.props.onClick({
      preventDefault: function() { prevented = true; },
      stopPropagation: function() { stopped = true; }
    });
    assert.strictEqual(prevented, false);
    assert.strictEqual(stopped, true);
  });

  test('calendar buttons select start and end dates without applying the filter prematurely', function() {
    const page = control(0);
    const buttons = findElements(page.instance.render(), function(element) {
      return element.type === 'button' && /lc-date-picker-button/.test(element.props.className || '');
    });
    assert.strictEqual(buttons.length, 2);
    buttons[0].props.onClick();
    assert.strictEqual(page.instance.state.datePickerTarget, 'start');
    page.instance.state.datePickerMonth = '2026-10';
    const startDay = findElements(page.instance.renderDatePicker(), function(element) {
      return element.type === 'button' && element.props.children === 22;
    })[0];
    startDay.props.onClick();
    assert.strictEqual(page.instance.state.draftFilterValue, '2026-10-22');
    assert.strictEqual(page.instance.state.datePickerTarget, '');
    buttons[1].props.onClick();
    page.instance.state.datePickerMonth = '2026-10';
    const endDay = findElements(page.instance.renderDatePicker(), function(element) {
      return element.type === 'button' && element.props.children === 24;
    })[0];
    endDay.props.onClick();
    assert.strictEqual(page.instance.state.draftFilterEndValue, '2026-10-24');
    assert.strictEqual(page.instance.state.datePickerTarget, '');
    assert.strictEqual(page.instance.state.columnFilters.EventDate, undefined);
  });

  test('calendar month navigation crosses year boundaries and preserves the draft date', function() {
    const page = control(0);
    page.instance.state.draftFilterValue = '2026-12-22';
    page.instance.toggleDatePicker('start');
    const navigation = findElements(page.instance.renderDatePicker(), function(element) {
      return element.type === 'button' && /month$/.test(element.props['aria-label'] || '');
    });
    navigation[1].props.onClick();
    assert.strictEqual(page.instance.state.datePickerMonth, '2027-01');
    navigation[0].props.onClick();
    assert.strictEqual(page.instance.state.datePickerMonth, '2026-12');
    assert.strictEqual(page.instance.state.draftFilterValue, '2026-12-22');
  });

  test('calendar and month-arrow symbols remain visible without the SharePoint icon font', function() {
    const calendar = styleFor('.lc-date-picker-button .ms-Icon--Calendar');
    assert.strictEqual(calendar.width, '14px');
    assert.strictEqual(calendar.height, '14px');
    assert.ok(calendar.border.indexOf('solid') >= 0);
    const calendarHeader = styleFor('.lc-date-picker-button .ms-Icon--Calendar::before');
    assert.strictEqual(calendarHeader.content, "''");
    assert.strictEqual(calendarHeader.height, '1px');
    const calendarDays = styleFor('.lc-date-picker-button .ms-Icon--Calendar::after');
    assert.strictEqual(calendarDays.content, "''");
    assert.ok(calendarDays.getPropertyValue('box-shadow'));
    const arrows = styleFor('.lc-date-picker-header .ms-Icon--ChevronLeft::before,\n.lc-date-picker-header .ms-Icon--ChevronRight::before');
    assert.strictEqual(arrows.content, "''");
    assert.ok(arrows.getPropertyValue('border-left').indexOf('solid') >= 0);
    assert.ok(styleFor('.lc-date-picker-header .ms-Icon--ChevronLeft::before').transform);
    assert.ok(styleFor('.lc-date-picker-header .ms-Icon--ChevronRight::before').transform);
  });

  test('date inputs can shrink to leave room for the calendar button', function() {
    const input = styleFor('.lc-date-input-row > .lc-filter-input');
    assert.strictEqual(input.getPropertyValue('min-width'), '0');
    assert.ok(input.flex);
  });

  test('the filter explains that an end date is optional for single-date matching', function() {
    const page = control(1);
    const hints = findElements(page.instance.render(), function(element) {
      return element.props.children === strings.RuntimeFilterSingleDateHelp;
    });
    assert.ok(strings.RuntimeFilterSingleDateHelp);
    assert.ok(hints.length > 0);
  });

  if (project === 'ListControl') {
    test('explicit date/time mode is retained and compares instants instead of timestamp spelling', function() {
      const page = control(1);
      const checkbox = findElements(page.instance.render(), function(element) {
        return element.type === 'input' && element.props.type === 'checkbox'
          && element.props['aria-label'] === strings.RuntimeFilterIncludeTimeLabel;
      })[0];
      assert.ok(checkbox, 'Date/time filtering remains available as an explicit option');
      checkbox.props.onChange({ currentTarget: { checked: true } });
      assert.strictEqual(page.instance.state.activeFilterCompareDateOnly, false);
      const filter = apply(page.instance, '2026-10-22T12:30');
      assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: '2026-10-22T12:30:00' }, page.field, filter), true);
      assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: '2026-10-22T12:31:00' }, page.field, filter), false);
      page.instance.openFilter(page.field, document.getElementById('anchor'));
      assert.strictEqual(page.instance.state.activeFilterCompareDateOnly, false);
    });

    test('existing explicit date/time ranges retain their precision', function() {
      const page = control(1);
      page.instance.state.activeFilterCompareDateOnly = false;
      const filter = apply(page.instance, '2026-10-22T12:30', '2026-10-22T13:00');
      assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: '2026-10-22T12:29:59' }, page.field, filter), false);
      assert.strictEqual(page.instance.rowMatchesFilter({ EventDate: '2026-10-22T13:00:00' }, page.field, filter), true);
    });

    test('switching date/time mode preserves dates and closes the calendar picker', function() {
      const page = control(1);
      page.instance.state.draftFilterValue = '2026-10-22';
      page.instance.state.draftFilterEndValue = '2026-10-24';
      page.instance.state.datePickerTarget = 'end';
      page.instance.changeDateFilterMode(true);
      assert.strictEqual(page.instance.state.draftFilterValue, '2026-10-22T00:00');
      assert.strictEqual(page.instance.state.draftFilterEndValue, '2026-10-24T00:00');
      assert.strictEqual(page.instance.state.datePickerTarget, '');
      page.instance.changeDateFilterMode(false);
      assert.strictEqual(page.instance.state.draftFilterValue, '2026-10-22');
      assert.strictEqual(page.instance.state.draftFilterEndValue, '2026-10-24');
      assert.strictEqual(page.instance.state.activeFilterCompareDateOnly, true);
    });
  }
  window.close();
});

console.log(passed + ' passed, ' + failed + ' failed');
process.exitCode = failed ? 1 : 0;
