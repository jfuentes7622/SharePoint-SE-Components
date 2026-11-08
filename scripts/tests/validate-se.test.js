'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { test } = require('node:test');
const { validate, parseJson, prohibited } = require('../validate-se');
const { affected } = require('../ci-projects');

const solutionId = '11111111-1111-1111-1111-111111111111';
const componentId = '22222222-2222-2222-2222-222222222222';
const featureId = '33333333-3333-3333-3333-333333333333';

function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'se-validator-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const write = (file, value) => {
    const target = path.join(root, 'Example', file);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, typeof value === 'string' ? value : JSON.stringify(value));
  };
  const change = (file, edit) => {
    const value = JSON.parse(fs.readFileSync(path.join(root, 'Example', file), 'utf8'));
    edit(value);
    write(file, value);
  };
  write('package.json', { version: '1.2.3', dependencies: { '@microsoft/sp-core-library': '~1.5.1' } });
  write('package-lock.json', {
    lockfileVersion: 1,
    dependencies: { '@microsoft/sp-core-library': { version: '1.5.1', dependencies: { '@microsoft/sp-polyfills': { version: '1.4.0' } } } }
  });
  write('config/package-solution.json', {
    solution: { id: solutionId, version: '1.2.3.4', skipFeatureDeployment: false, features: [{ id: featureId, version: '1.2.3.3' }] },
    paths: { zippedPackage: 'solution/example.sppkg' }
  });
  write('src/Example.manifest.json', { id: componentId, version: '*', componentType: 'WebPart' });
  write('src/example.ts', 'export const value = 1;');
  write('config/config.json', { bundles: { example: { components: [{ entrypoint: './lib/example.js', manifest: './src/Example.manifest.json' }] } } });
  write('config/serve.json', { port: 4321 });
  write('config/write-manifests.json', { cdnBasePath: '<!-- PATH TO CDN -->' });
  const git = (...args) => {
    const result = spawnSync('git', args, { cwd: root, encoding: 'utf8' });
    assert.equal(result.status, 0, result.stderr);
    return result.stdout.trim();
  };
  git('init', '--quiet');
  git('add', '.');
  git('-c', 'user.name=Validation Fixture', '-c', 'user.email=fixture@example.invalid',
    '-c', 'commit.gpgsign=false', 'commit', '--quiet', '-m',
    'Create test fixture\n\nCo-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>');
  const commit = () => {
    git('add', '.');
    git('-c', 'user.name=Validation Fixture', '-c', 'user.email=fixture@example.invalid',
      '-c', 'commit.gpgsign=false', 'commit', '--quiet', '-m',
      'Update test fixture\n\nCo-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>');
    return git('rev-parse', 'HEAD');
  };
  return { root, write, change, commit, base: git('rev-parse', 'HEAD') };
}

test('JSONC preserves URLs and escaped quotes while stripping comments', () => {
  assert.deepEqual(parseJson('{"url":"https://example.test", /* comment */ "x":"a\\"b" // end\n}', 'fixture'),
    { url: 'https://example.test', x: 'a"b' });
  assert.throws(() => parseJson('{/*', 'fixture'), /unterminated/);
  assert.throws(() => parseJson('{"x":}', 'fixture'), /fixture/);
});

test('baseline passes with wildcard manifest and older bundled transitive dependencies', t => {
  const f = fixture(t);
  assert.deepEqual(validate(f.root, f.base).errors, []);
});

test('detects framework dependency and resolved-lock upgrades', t => {
  const f = fixture(t);
  f.change('package.json', pkg => { pkg.dependencies['@microsoft/sp-core-library'] = '~1.6.0'; });
  f.change('package-lock.json', lock => { lock.dependencies['@microsoft/sp-core-library'].version = '1.6.0'; });
  assert.ok(validate(f.root, f.base).errors.some(error => /must stay on 1.5.1/.test(error)));
  assert.ok(validate(f.root, f.base).errors.some(error => /must resolve to 1.5.1/.test(error)));
});

test('detects all ID changes and version regressions without demanding equal versions', t => {
  const f = fixture(t);
  f.change('config/package-solution.json', config => {
    config.solution.id = componentId;
    config.solution.version = '1.2.3.2';
    config.solution.features[0].id = componentId;
  });
  f.change('src/Example.manifest.json', manifest => { manifest.id = solutionId; });
  const errors = validate(f.root, f.base).errors;
  for (const pattern of [/solution ID changed/, /component ID changed/, /feature ID removed/, /version regressed/]) {
    assert.ok(errors.some(error => pattern.test(error)), String(pattern));
  }
});

test('deployment scope changes fail', t => {
  const f = fixture(t);
  f.change('config/package-solution.json', config => { config.solution.skipFeatureDeployment = true; });
  assert.ok(validate(f.root, f.base).errors.some(error => /skipFeatureDeployment changed/.test(error)));
});

test('duplicate component IDs fail; feature and component namespaces are separate', t => {
  const f = fixture(t);
  f.change('config/package-solution.json', config => { config.solution.features[0].id = componentId; });
  assert.deepEqual(validate(f.root).errors, []);
  f.write('src/Duplicate.manifest.json', { id: componentId.toUpperCase(), version: '*', componentType: 'WebPart' });
  assert.ok(validate(f.root).errors.some(error => /duplicate component ID/.test(error)));
});

test('deleted manifests and broken bundle references fail', t => {
  const f = fixture(t);
  fs.unlinkSync(path.join(f.root, 'Example', 'src', 'Example.manifest.json'));
  const errors = validate(f.root, f.base).errors;
  assert.ok(errors.some(error => /removed or renamed/.test(error)));
  assert.ok(errors.some(error => /missing source manifest/.test(error)));
});

test('known new Graph/modern imports and endpoints fail', t => {
  const f = fixture(t);
  f.write('src/example.ts', "import { graph } from '@pnp/graph'; fetch('https://graph.microsoft.com/v1.0/me');");
  assert.equal(validate(f.root, f.base).errors.filter(error => /prohibited/.test(error)).length, 2);
  assert.equal(prohibited("import x from '@microsoft/sp-adaptive-card-extension-base';").length, 1);
});

test('existing prohibited patterns cannot be multiplied', t => {
  const f = fixture(t);
  f.write('src/example.ts', "import x from '@pnp/graph';");
  const base = f.commit();
  assert.deepEqual(validate(f.root, base).errors, []);
  f.write('src/example.ts', "import x from '@pnp/graph';\nimport y from '@pnp/graph';");
  assert.equal(validate(f.root, base).errors.filter(error => /prohibited/.test(error)).length, 1);
});

test('affected project selection includes shared changes but skips docs-only builds', () => {
  const names = ['Accordion', 'Tiles'];
  assert.deepEqual(affected(names, ['Accordion/src/example.ts']), ['Accordion']);
  assert.deepEqual(affected(names, ['docs/Deployment.md', '.github/agents/triage.agent.md']), []);
  assert.deepEqual(affected(names, ['scripts/validate-se.js']), names);
  assert.deepEqual(affected(names, ['Tiles/src/a.ts', 'Accordion/src/b.ts']), names);
});
