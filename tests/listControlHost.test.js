'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const projects = [
  { name: 'ReportForms', control: 'ListControl', componentId: '3fbfa88a-4ac3-4b8e-b7a6-eb5256cc7001',
    rootClass: 'lc-root', ownerAttribute: 'data-reportforms-listcontrol-owner',
    markerAttribute: 'data-reportforms-listcontrol-configure' },
  { name: 'DynamicForms', control: 'GridControl', componentId: 'de5f92e8-570c-4a7c-ba96-7ae89a098723',
    rootClass: 'lc-root gc-root', ownerAttribute: 'data-dynamicforms-gridcontrol-owner',
    markerAttribute: 'data-dynamicforms-gridcontrol-configure' }
];
let passed = 0;
let failed = 0;

projects.forEach(function(configuration) {
const project = path.join(__dirname, '..', configuration.name);
const ts = require(path.join(project, 'node_modules', 'typescript'));
const React = require(path.join(project, 'node_modules', 'react'));
const jsdom = require(path.join(project, 'node_modules', 'jsdom'));
const source = fs.readFileSync(path.join(project, 'src', 'webparts', 'sharePointDynamicForm',
  'components', configuration.control + 'Host.tsx'), 'utf8');
const compiled = ts.transpileModule(source, {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5, jsx: ts.JsxEmit.React }
}).outputText;
const componentId = configuration.componentId;
const instanceId = 'b65121db-8170-4661-9c01-0621064215cb';
const configEvent = 'spse:' + configuration.control.toLowerCase() + '-runtime-config';
const requestEvent = configEvent + '-request';

function fixture() {
  const document = jsdom.jsdom('<!doctype html><html><body></body></html>');
  const window = document.defaultView;
  let now = 0;
  let nextId = 0;
  const timers = new Map();
  window.setTimeout = function(callback, delay) {
    const id = ++nextId;
    timers.set(id, { callback: callback, at: now + delay });
    return id;
  };
  window.setInterval = function(callback, delay) {
    const id = ++nextId;
    timers.set(id, { callback: callback, at: now + delay, interval: delay });
    return id;
  };
  window.clearTimeout = window.clearInterval = function(id) { timers.delete(id); };
  const exports = {};
  vm.runInNewContext(compiled, {
    exports: exports, window: window, document: document,
    require: function(name) {
      if (name === 'react') { return React; }
      if (name === 'SharePointDynamicFormWebPartStrings') { return {}; }
      throw new Error('Unexpected import: ' + name);
    }
  });
  const hosts = [];
  const configs = [];
  window.addEventListener(configEvent, function(event) { configs.push(event.detail); });
  return {
    document: document,
    timers: timers,
    configs: configs,
    host: function(id, owner, editMode) {
      const host = new exports[configuration.control + 'Host']({
        fieldId: 'field-' + owner, ownerId: owner, sourceId: 'WebPart.' + componentId + '.' + id,
        sourceName: 'List ' + owner, sourceListName: 'Items', runtimeFilterJson: '[{"value":7}]',
        defaultField: 'ParentId', defaultValue: 7, readOnly: true,
        isPageEditMode: !!editMode
      });
      host.setState = function(update) { this.state = Object.assign({}, this.state, update); };
      host._mountElement = document.createElement('div');
      document.body.appendChild(host._mountElement);
      hosts.push(host);
      host.componentDidMount();
      return host;
    },
    target: function(id, mixedAttributes) {
      const wrapper = document.createElement('div');
      wrapper.id = 'WebPart.' + componentId + '.' + id;
      if (mixedAttributes) { wrapper.setAttribute('data-sp-webpart-id', componentId); }
      const target = document.createElement('div');
      target.innerHTML = '<div class="' + configuration.rootClass + '"></div>';
      wrapper.appendChild(target);
      document.body.appendChild(wrapper);
      return target;
    },
    ready: function(id) {
      window.dispatchEvent(new window.CustomEvent(requestEvent, { detail: { instanceId: id } }));
    },
    advance: function(duration) {
      const end = now + duration;
      let steps = 0;
      while (true) {
        const pending = Array.from(timers.entries()).filter(function(entry) { return entry[1].at <= end; })
          .sort(function(a, b) { return a[1].at - b[1].at || a[0] - b[0]; });
        if (!pending.length) { break; }
        assert.ok(++steps < 10000, 'Timer loop did not settle');
        const id = pending[0][0];
        const timer = pending[0][1];
        now = timer.at;
        if (timer.interval) { timer.at += timer.interval; } else { timers.delete(id); }
        timer.callback();
      }
      now = end;
    },
    dispose: function() {
      hosts.forEach(function(host) { host.componentWillUnmount(); });
      window.close();
    }
  };
}

function test(name, run) {
  const page = fixture();
  try {
    run(page);
    passed += 1;
    console.log('PASS ' + configuration.name + ': ' + name.replace(/ListControl/g, configuration.control));
  } catch (error) {
    failed += 1;
    console.error('FAIL ' + configuration.name + ': ' + name.replace(/ListControl/g, configuration.control) + ': ' + error.message);
  } finally {
    page.dispose();
  }
}

test('embeds an already mounted ListControl and restores its original position', function(page) {
  const target = page.target(instanceId);
  const parent = target.parentNode;
  const sibling = page.document.createElement('span');
  parent.appendChild(sibling);
  const host = page.host(instanceId, 'report:field', true);
  assert.strictEqual(host.state.status, 'embedded');
  assert.strictEqual(target.parentNode, host._mountElement);
  assert.ok(parent.querySelector('[' + configuration.markerAttribute + ']'));
  host.componentWillUnmount();
  assert.strictEqual(target.parentNode, parent);
  assert.strictEqual(target.nextSibling, sibling);
  assert.strictEqual(target.getAttribute(configuration.ownerAttribute), null);
  assert.strictEqual(parent.querySelector('[' + configuration.markerAttribute + ']'), null);
});

test('keeps the existing short polling discovery', function(page) {
  const host = page.host(instanceId, 'report:field');
  page.advance(1500);
  const target = page.target(instanceId);
  page.advance(250);
  assert.strictEqual(target.parentNode, host._mountElement);
  assert.strictEqual(page.timers.size, 0);
});

test('recovers from the 10-second timeout when ListControl mounts at 30 seconds', function(page) {
  const host = page.host(instanceId, 'report:field');
  page.advance(30000);
  assert.strictEqual(host.state.status, 'missing');
  assert.strictEqual(page.timers.size, 0);
  const target = page.target(instanceId);
  page.ready(instanceId.toUpperCase());
  assert.notStrictEqual(target.parentNode, host._mountElement, 'Do not move DOM during ListControl componentDidMount');
  page.advance(0);
  assert.strictEqual(host.state.status, 'embedded');
  assert.strictEqual(target.parentNode, host._mountElement);
  assert.strictEqual(page.configs[page.configs.length - 1].filterJson, host.props.runtimeFilterJson);
});

test('discovers the instance ID even when a different SharePoint ID attribute exists', function(page) {
  const target = page.target(instanceId, true);
  const host = page.host(instanceId, 'report:field');
  assert.strictEqual(host.state.status, 'embedded');
  assert.strictEqual(target.parentNode, host._mountElement);
});

test('restarts bounded discovery if readiness arrives before the host is discoverable', function(page) {
  const host = page.host(instanceId, 'report:field');
  page.advance(11000);
  page.ready(instanceId);
  page.advance(0);
  const target = page.target(instanceId);
  page.advance(250);
  assert.strictEqual(target.parentNode, host._mountElement);
  assert.strictEqual(host.state.status, 'embedded');
  assert.strictEqual(page.timers.size, 0);
});

test('recovers twelve independently linked reports without unrelated discovery scans', function(page) {
  const ids = Array.from({ length: 12 }, function(value, index) {
    return '00000000-0000-4000-8000-' + String(index + 1).padStart(12, '0');
  });
  const hosts = ids.map(function(id, index) { return page.host(id, 'report-' + index + ':field'); });
  page.advance(15000);
  const scans = hosts.map(function(host) {
    const calls = { count: 0 };
    const discover = host.tryEmbedTarget;
    host.tryEmbedTarget = function() { calls.count += 1; return discover.call(this); };
    return calls;
  });
  const targets = ids.map(function(id) { return page.target(id); });
  page.ready(ids[0]);
  page.advance(0);
  assert.strictEqual(hosts[0].state.status, 'embedded');
  scans.slice(1).forEach(function(calls) { assert.strictEqual(calls.count, 0); });
  ids.slice(1).reverse().forEach(page.ready);
  page.advance(0);
  hosts.forEach(function(host, index) {
    assert.strictEqual(host.state.status, 'embedded');
    assert.strictEqual(targets[index].parentNode, host._mountElement);
    assert.strictEqual(targets[index].getAttribute(configuration.ownerAttribute), host.props.ownerId);
  });
  assert.strictEqual(page.timers.size, 0);
});

test('coalesces readiness events and uses the latest filter', function(page) {
  const host = page.host(instanceId, 'report:field');
  page.advance(11000);
  page.target(instanceId);
  page.ready(instanceId);
  page.ready(instanceId);
  assert.strictEqual(page.timers.size, 1);
  host.props = Object.assign({}, host.props, {
    runtimeFilterJson: '[{"value":42}]', defaultField: 'RelatedId', defaultValue: 42, readOnly: false
  });
  page.advance(0);
  assert.strictEqual(page.configs.filter(function(config) { return config.active; }).length, 1);
  assert.strictEqual(page.configs[0].filterJson, '[{"value":42}]');
  if (configuration.name === 'DynamicForms') {
    assert.strictEqual(page.configs[0].defaultField, 'RelatedId');
    assert.strictEqual(page.configs[0].defaultValue, 42);
    assert.strictEqual(page.configs[0].readOnly, false);
  }
  page.ready(instanceId);
  assert.strictEqual(page.configs.length, 2, 'Already embedded controls still receive their runtime config');
});

test('unmount cancels pending discovery and removes the readiness listener', function(page) {
  const host = page.host(instanceId, 'report:field');
  page.advance(11000);
  const target = page.target(instanceId);
  const originalParent = target.parentNode;
  page.ready(instanceId);
  host.componentWillUnmount();
  page.advance(0);
  page.ready(instanceId);
  page.advance(0);
  assert.strictEqual(target.parentNode, originalParent);
  assert.strictEqual(page.timers.size, 0);
  assert.strictEqual(page.configs.filter(function(config) { return config.active; }).length, 0);
});

test('changing the configured source cancels the old deferred discovery', function(page) {
  const host = page.host(instanceId, 'report:field');
  page.advance(11000);
  const oldTarget = page.target(instanceId);
  page.ready(instanceId);
  const previous = host.props;
  const nextId = '00000000-0000-4000-8000-000000000099';
  host.props = Object.assign({}, previous, { sourceId: 'WebPart.' + componentId + '.' + nextId });
  host.componentDidUpdate(previous);
  page.advance(0);
  assert.notStrictEqual(oldTarget.parentNode, host._mountElement);
  const target = page.target(nextId);
  page.ready(nextId);
  page.advance(0);
  assert.strictEqual(target.parentNode, host._mountElement);
  assert.strictEqual(page.timers.size, 0);
});

test('missing or differently owned controls remain missing without endless polling', function(page) {
  const host = page.host(instanceId, 'report:field');
  page.advance(11000);
  page.ready('00000000-0000-4000-8000-000000000099');
  assert.strictEqual(page.timers.size, 0);
  assert.strictEqual(host.state.status, 'missing');
  const target = page.target(instanceId);
  target.setAttribute(configuration.ownerAttribute, 'another-report:field');
  page.ready(instanceId);
  page.advance(11000);
  assert.strictEqual(host.state.status, 'missing');
  assert.notStrictEqual(target.parentNode, host._mountElement);
  assert.strictEqual(target.getAttribute(configuration.ownerAttribute), 'another-report:field');
  assert.strictEqual(page.timers.size, 0);
});

if (configuration.name === 'DynamicForms') {
  test('keeps filter, default values and read-only mode synchronized after embedding', function(page) {
    page.target(instanceId);
    const host = page.host(instanceId, 'form:field');
    assert.strictEqual(page.configs[0].defaultField, 'ParentId');
    assert.strictEqual(page.configs[0].defaultValue, 7);
    assert.strictEqual(page.configs[0].readOnly, true);
    const previous = host.props;
    host.props = Object.assign({}, previous, {
      runtimeFilterJson: '[{"value":99}]', defaultField: 'RelatedId', defaultValue: 99, readOnly: false
    });
    host.componentDidUpdate(previous);
    const config = page.configs[page.configs.length - 1];
    assert.strictEqual(config.active, true);
    assert.strictEqual(config.filterJson, host.props.runtimeFilterJson);
    assert.strictEqual(config.defaultField, 'RelatedId');
    assert.strictEqual(config.defaultValue, 99);
    assert.strictEqual(config.readOnly, false);
  });
}
});

console.log(passed + ' passed, ' + failed + ' failed');
process.exitCode = failed ? 1 : 0;
