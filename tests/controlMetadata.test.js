'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

let passed = 0;
let failed = 0;

function serializeXml(node) {
  function escape(value) {
    return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  }
  if (node.nodeType === 9) { return serializeXml(node.documentElement); }
  if (node.nodeType === 3) { return escape(node.nodeValue); }
  const attributes = Array.from(node.attributes).map(function(attribute) {
    return ' ' + attribute.name + '="' + escape(attribute.value) + '"';
  }).join('');
  return '<' + node.nodeName + attributes + '>' + Array.from(node.childNodes).map(serializeXml).join('')
    + '</' + node.nodeName + '>';
}

async function runProject(project) {
  const root = path.join(__dirname, '..', project);
  const componentName = project === 'ListControl' ? 'listControl' : 'gridControl';
  const ts = require(path.join(root, 'node_modules', 'typescript'));
  const React = require(path.join(root, 'node_modules', 'react'));
  const jsdom = require(path.join(root, 'node_modules', 'jsdom'));
  const document = jsdom.jsdom('<!doctype html><html><body></body></html>');
  const window = document.defaultView;
  const base = path.join(root, 'src', 'webparts', componentName);

  function load(file) {
    const source = fs.readFileSync(file, 'utf8');
    const output = ts.transpileModule(source, {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
    }).outputText;
    const exports = {};
    const imports = {
      react: React, 'react-dom': {},
      '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } },
      '@microsoft/sp-webpart-base': { BaseClientSideWebPart: function() {} },
      '@microsoft/sp-core-library': {},
      '@microsoft/sp-component-base/lib/DynamicProperty': {},
      '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField': {},
      '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCollectionData/PropertyPaneCollectionData': {},
      '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker': {},
      '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData': {},
      './components/ListControl': {}, './components/GridControl': {},
      './components/ListDesigner': {}, './components/GridDesigner': {},
      './DateTime24HourInput': {}, './GridValidation': {}, './GridRichTextEditor': {},
      './ListControl.css': {}, './GridControl.css': {},
      './ListDesigner.css': {}, './GridDesigner.css': {},
      '../shared/deterministicFullWidth': {}, '../shared/overrideCss': {},
      '../../../config/package-solution.json': require(path.join(root, 'config', 'package-solution.json'))
    };
    imports[project + 'WebPartStrings'] = {};
    vm.runInNewContext(output, {
      exports: exports, document: document, window: window,
      DOMParser: window.DOMParser,
      XMLSerializer: function() {
        this.serializeToString = serializeXml;
      },
      console: console, setTimeout: setTimeout, clearTimeout: clearTimeout,
      require: function(name) {
        assert.ok(Object.prototype.hasOwnProperty.call(imports, name), 'Unexpected import: ' + name);
        return imports[name];
      }
    }, { filename: file });
    return exports;
  }

  const Control = load(path.join(base, 'components', project + '.tsx'))[project];
  const Designer = load(path.join(base, 'components', project === 'ListControl' ? 'ListDesigner.tsx' : 'GridDesigner.tsx'))[
    project === 'ListControl' ? 'ListDesigner' : 'GridDesigner'];
  const WebPart = load(path.join(base, project + 'WebPart.ts')).default;

  function schema(multi) {
    return { steps: [{ fields: [
      { id: 'title', fieldName: 'Title', type: 'text', label: 'Title' },
      { id: 'status', fieldName: 'Status', type: multi ? 'multiselect' : 'dropdown', config: { choices: ['Old choice'] } },
      { id: 'new', fieldName: 'NewColumn', type: 'text', label: 'New column', gridWidth: '160px' }
    ] }] };
  }

  function fixture(options) {
    options = options || {};
    const requests = [];
    const server = { choices: ['Current choice', 'New choice'], metadataFailure: false };
    function metadata() {
      return [
        { InternalName: 'Title', Title: 'Title', TypeAsString: 'Text' },
        { InternalName: 'Status', Title: 'Status', TypeAsString: options.multi ? 'MultiChoice' : 'Choice', Choices: server.choices },
        { InternalName: 'NewColumn', Title: 'New column', TypeAsString: 'Text' }
      ];
    }
    function response(data, status) {
      return Promise.resolve({
        ok: !status || status === 200, status: status || 200, statusText: status ? 'Forbidden' : 'OK',
        json: function() { return Promise.resolve(data); }, text: function() { return Promise.resolve(JSON.stringify(data)); }
      });
    }
    const context = {
      pageContext: { web: { absoluteUrl: 'https://example.test/site' } },
      spHttpClient: {
        get: function(url, configuration, requestOptions) {
          requests.push({ method: 'GET', url: url, options: requestOptions });
          if (url.indexOf('/ViewFields') >= 0) { return response({ value: ['Title', 'Status'] }); }
          if (url.indexOf('/views/') >= 0) {
            return response({ ViewQuery: '<Where><Eq><FieldRef Name="Status"/><Value Type="Choice">Current choice</Value></Eq></Where><OrderBy><FieldRef Name="Title" Ascending="FALSE"/></OrderBy>',
              Scope: 1, RowLimit: 50 });
          }
          if (url.indexOf('/fields?') >= 0) {
            if (server.metadataFailure) { return response({}, 403); }
            return response(options.verbose ? { d: { results: metadata().map(function(field) {
              return Object.assign({}, field, field.Choices ? { Choices: { results: field.Choices } } : {});
            }) } } : { value: metadata() });
          }
          if (url.indexOf('/items?') >= 0) {
            const itemId = decodeURIComponent(url).indexOf('ID eq 2') >= 0 ? 2 : 1;
            return response({ value: [{ ID: itemId, Title: 'Row', Status: 'Current choice', NewColumn: 'New value' }] });
          }
          throw new Error('Unexpected GET ' + url);
        },
        post: function(url, configuration, requestOptions) {
          requests.push({ method: 'POST', url: url, options: requestOptions });
          assert.ok(url.indexOf('/RenderListDataAsStream') >= 0, 'Unexpected POST ' + url);
          const body = JSON.parse(requestOptions.body);
          const xml = new window.DOMParser().parseFromString(body.parameters.ViewXml, 'text/xml');
          let names = Array.from(xml.getElementsByTagName('FieldRef')).filter(function(ref) {
            return ref.parentNode.nodeName === 'ViewFields';
          }).map(function(ref) { return ref.getAttribute('Name'); });
          if (!names.length) { names = ['Title', 'Status']; }
          const row = { ID: body.parameters.Paging ? 2 : 1 };
          const values = { Title: 'Row', Status: 'Current choice', NewColumn: 'New value' };
          names.forEach(function(name) { row[name] = values[name]; });
          return response({
            ListSchema: { Field: names.map(function(name) { return { Name: name, DisplayName: name }; }) },
            Row: [row], NextHref: options.paging && !body.parameters.Paging ? '?Paged=TRUE&p_ID=1' : ''
          });
        }
      }
    };
    const props = {
      context: context, listName: 'Items', defaultViewId: options.noView ? '' : 'view-id', views: [],
      viewColumns: [
        { fieldName: 'Title', displayName: 'Title' }, { fieldName: 'Status', displayName: 'Status' },
        { fieldName: 'NewColumn', displayName: 'New column', width: '160px' }
      ],
      gridSchemaJson: JSON.stringify(schema(options.multi)), isEditMode: false, enableDiagnostics: false,
      fetchBatchSize: 500, pageSize: 50, onFilteredCountChange: function() {}, onSelectionChange: function() {}
    };
    const control = new Control(props);
    control.setState = function(update, callback) {
      this.state = Object.assign({}, this.state, update);
      if (callback) { callback(); }
    };
    control.state.accessLoading = false;
    control.state.hasReadAccess = true;
    control.loadIsDocumentLibrary = async function() { return false; };
    control.loadLookupOptionsByField = async function() { return {}; };
    return { control: control, props: props, requests: requests, server: server, metadata: metadata };
  }

  async function test(name, run) {
    try {
      await run();
      passed += 1;
      console.log('PASS ' + project + ': ' + name);
    } catch (error) {
      failed += 1;
      console.error('FAIL ' + project + ': ' + name + ': ' + error.message);
    }
  }

  for (const noView of [false, true]) {
    await test('published columns are requested and populated (selected view ' + !noView + ')', async function() {
      const page = fixture({ noView: noView });
      await page.control.loadRows();
      assert.strictEqual(page.control.state.error, null);
      const xml = JSON.parse(page.requests.find(function(request) { return request.method === 'POST'; }).options.body).parameters.ViewXml;
      assert.ok(/Name="NewColumn"/.test(xml), 'New column must be included in RenderListDataAsStream, not only designer state');
      assert.strictEqual(page.control.getDisplayFields().some(function(field) { return field.Name === 'NewColumn'; }), true);
      assert.strictEqual(page.control.state.rows[0].NewColumn, 'New value');
      if (!noView) {
        assert.ok(xml.indexOf('<Where>') >= 0, 'Selected-view filtering is preserved');
        assert.ok(xml.indexOf('<OrderBy>') >= 0, 'Selected-view sorting is preserved');
      }
    });
  }

  for (const multi of [false, true]) {
  await test('refresh uses current ' + (multi ? 'MultiChoice' : 'Choice') + ' options, including removal of all options', async function() {
    const page = fixture({ verbose: true, multi: multi });
    for (const choices of [['New option', 'Another option'], []]) {
      page.server.choices = choices;
      await page.control.loadRows();
      assert.strictEqual(page.control.state.error, null);
      const field = page.control.state.fields.find(function(candidate) { return candidate.Name === 'Status'; });
      const actual = project === 'ListControl' ? page.control.getStaticFilterOptions(field).map(function(option) {
        return option.value;
      }) : page.control.getGridFieldMetadata(field).choices;
      assert.strictEqual(JSON.stringify(actual), JSON.stringify(choices));
    }
  });
  }

  await test('saving an added column reloads its runtime data without recreating the component', async function() {
    const page = fixture();
    const oldSchema = { steps: [{ fields: [{ fieldName: 'Title', type: 'text' }] }] };
    page.control.props = Object.assign({}, page.props, {
      viewColumns: [{ fieldName: 'Title', displayName: 'Title' }], gridSchemaJson: JSON.stringify(oldSchema)
    });
    await page.control.loadRows();
    assert.strictEqual(page.control.getDisplayFields().some(function(field) { return field.Name === 'NewColumn'; }), false);
    const previous = page.control.props;
    const loadRows = page.control.loadRows.bind(page.control);
    let pending;
    page.control.loadRows = function() { pending = loadRows(); return pending; };
    page.control.props = page.props;
    page.control.componentDidUpdate(previous, page.control.state);
    assert.ok(pending, 'Saving the design must trigger a data reload');
    await pending;
    assert.strictEqual(page.control.state.error, null);
    assert.strictEqual(page.control.getDisplayFields().some(function(field) { return field.Name === 'NewColumn'; }), true);
    assert.strictEqual(page.control.state.rows[0].NewColumn, 'New value');
  });

  await test('configured columns remain in subsequent paging requests', async function() {
    const page = fixture({ paging: true });
    await page.control.loadRows();
    assert.strictEqual(page.control.state.error, null);
    await page.control.loadNextBatch();
    const requests = page.requests.filter(function(request) { return request.method === 'POST'; });
    assert.strictEqual(requests.length, 2);
    requests.forEach(function(request) {
      assert.ok(JSON.parse(request.options.body).parameters.ViewXml.indexOf('Name="NewColumn"') >= 0);
    });
    assert.strictEqual(page.control.state.rows.length, 2);
    assert.strictEqual(page.control.state.rows[1].NewColumn, 'New value');
  });

  await test('runtime GET fallbacks bypass the browser cache', async function() {
    const page = fixture();
    const calls = [];
    page.props.context.spHttpClient.get = async function(url, configuration, options) {
      calls.push(options);
      return { ok: calls.length === 4, status: calls.length === 4 ? 200 : 406, statusText: 'Not acceptable' };
    };
    await page.control.getJsonWithFallback('https://example.test/site/_api/web/lists');
    assert.strictEqual(calls.length, 4);
    calls.forEach(function(options) {
      assert.strictEqual(options.cache, 'no-store');
      assert.strictEqual(options.headers['Cache-Control'], 'no-cache, no-store');
    });
  });

  await test('designer field/view GET fallbacks bypass the browser cache', async function() {
    const page = fixture();
    const calls = [];
    const designer = new Designer({
      context: page.props.context, listName: 'Items', viewId: 'view-id',
      schemaJson: JSON.stringify(schema()), viewColumns: page.props.viewColumns, enableDiagnostics: false
    });
    page.props.context.spHttpClient.get = async function(url, configuration, options) {
      calls.push(options);
      return { ok: calls.length === 4, status: calls.length === 4 ? 200 : 406, statusText: 'Not acceptable' };
    };
    await designer.getFieldsResponse('https://example.test/site/_api/web/lists/fields');
    assert.strictEqual(calls.length, 4);
    calls.forEach(function(options) {
      assert.strictEqual(options && options.cache, 'no-store');
      assert.strictEqual(options.headers['Cache-Control'], 'no-cache, no-store');
    });
  });

  await test('property-pane metadata GET fallbacks bypass the browser cache', async function() {
    const page = fixture();
    const calls = [];
    const webpart = new WebPart();
    webpart.context = page.props.context;
    page.props.context.spHttpClient.get = async function(url, configuration, options) {
      calls.push(options);
      return { ok: calls.length === 3, status: calls.length === 3 ? 200 : 406, statusText: 'Not acceptable',
        json: async function() { return { value: [] }; } };
    };
    await webpart.getJsonWithAcceptFallback('https://example.test/site/_api/web/lists');
    assert.strictEqual(calls.length, 3);
    calls.forEach(function(options) {
      assert.strictEqual(options.cache, 'no-store');
      assert.strictEqual(options.headers['Cache-Control'], 'no-cache, no-store');
    });
  });

  await test('field metadata errors remain visible instead of using saved choices', async function() {
    const page = fixture();
    page.server.metadataFailure = true;
    await page.control.loadRows();
    assert.ok(page.control.state.error, 'Metadata failures must surface in the runtime error state');
    assert.ok(page.control.state.error.indexOf('403') >= 0);
  });

  if (project === 'GridControl') {
    await test('saved schema columns are not hidden by the old SharePoint view column snapshot', async function() {
      const page = fixture();
      page.props.viewColumns = [{ fieldName: 'Title', displayName: 'Title' }];
      await page.control.loadRows();
      assert.strictEqual(page.control.state.error, null);
      const fields = page.control.getDisplayFields();
      assert.strictEqual(JSON.stringify(fields.map(function(field) { return field.Name; })), '["Title","Status","NewColumn"]');
      assert.strictEqual(fields[2].ConfiguredWidth, '160px');
    });

    await test('Grid Designer refreshes saved Choice options from SharePoint', async function() {
      const page = fixture();
      const designer = new Designer({
        context: page.props.context, listName: 'Items', schemaJson: JSON.stringify(schema()), enableDiagnostics: false
      });

      await test('Grid Designer marks SharePoint choices as managed by the list', async function() {
        const page = fixture();
        const designer = new Designer({
          context: page.props.context, listName: 'Items', schemaJson: JSON.stringify(schema()), enableDiagnostics: false
        });
        designer.setState = function(update) { this.state = Object.assign({}, this.state, update); };
        await designer.loadFields();
        designer.state.selectedFieldId = 'status';
        const controls = [];
        function visit(element) {
          if (!element || !element.props) { return; }
          if (element.type === 'textarea') { controls.push(element.props); }
          React.Children.forEach(element.props.children, visit);
        }
        visit(designer.renderFieldEditor());
        const choices = controls.find(function(props) { return props.value === page.server.choices.join('\n'); });
        assert.ok(choices);
        assert.strictEqual(choices.readOnly, true);
      });

      await test('legacy view-column layouts still work when no grid schema is configured', async function() {
        const page = fixture();
        page.props.gridSchemaJson = '';
        await page.control.loadRows();
        assert.strictEqual(page.control.state.error, null);
        assert.strictEqual(page.control.getDisplayFields().some(function(field) { return field.Name === 'NewColumn'; }), true);
        assert.strictEqual(page.control.state.rows[0].NewColumn, 'New value');
      });
      designer.setState = function(update) { this.state = Object.assign({}, this.state, update); };
      for (const choices of [['Current option'], []]) {
        page.server.choices = choices;
        await designer.loadFields();
        assert.strictEqual(designer.state.error, '');
        const status = designer.state.fields.find(function(field) { return field.fieldName === 'Status'; });
        assert.strictEqual(JSON.stringify(status.config.choices), JSON.stringify(choices));
      }
    });
  }
  window.close();
}

async function main() {
  await runProject('ListControl');
  await runProject('GridControl');
  console.log(passed + ' passed, ' + failed + ' failed');
  process.exitCode = failed ? 1 : 0;
}
main().catch(function(error) { console.error(error); process.exitCode = 1; });
