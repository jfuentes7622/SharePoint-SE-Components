'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const attachPackageCopy = require('../collect-solution-package');

function fixture(t) {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'spse-package-copy-'));
  t.after(() => fs.rmSync(temporary, { recursive: true, force: true }));
  const repository = path.join(temporary, 'Components');
  const project = path.join(repository, 'Example');
  fs.mkdirSync(path.join(project, 'config'), { recursive: true });
  fs.mkdirSync(path.join(project, 'sharepoint', 'solution'), { recursive: true });
  fs.writeFileSync(path.join(project, 'package.json'), '{}');
  const config = path.join(project, 'config', 'package-solution.json');
  fs.writeFileSync(config, JSON.stringify({ paths: { zippedPackage: 'solution/example.sppkg' } }));
  const source = path.join(project, 'sharepoint', 'solution', 'example.sppkg');
  fs.writeFileSync(source, Buffer.from([0x50, 0x4b, 0, 0xff, 42]));
  const directory = path.join(temporary, 'SharePoint-SE-Packages');
  const destination = path.join(directory, 'example.sppkg');
  return { temporary, repository, project, config, source, directory, destination };
}

function buildStub(original) {
  const events = [];
  const definition = { executable: original, arguments: { custom: 'preserved' } };
  const other = { executable: { execute: async () => events.push('build') } };
  const rig = {
    getTasks() {
      assert.equal(this, rig);
      return new Map([['package-solution', definition], ['build', other], ['bundle', other]]);
    }
  };
  const build = {
    rig,
    subTask(name, callback) {
      return {
        execute: options => new Promise((resolve, reject) => {
          events.push(name);
          callback.call({ log: message => events.push(message) }, {}, options,
            error => error ? reject(error) : resolve());
        })
      };
    },
    serial(...tasks) {
      return {
        execute: async options => {
          for (const task of tasks) {
            if (!task.isEnabled || task.isEnabled(options)) { await task.execute(options); }
          }
        }
      };
    }
  };
  return { build, events, definition, other };
}

test('all independent projects attach collection before initialization and have unique filenames', () => {
  const root = path.resolve(__dirname, '..', '..');
  const projects = fs.readdirSync(root).filter(name =>
    fs.existsSync(path.join(root, name, 'package.json'))
    && fs.existsSync(path.join(root, name, 'config', 'package-solution.json')));
  assert.ok(projects.length > 0);
  const filenames = new Set();
  for (const project of projects) {
    const gulpfile = fs.readFileSync(path.join(root, project, 'gulpfile.js'), 'utf8');
    const hook = "require('../scripts/collect-solution-package')(build, __dirname);";
    assert.equal(gulpfile.split(hook).length - 1, 1, project);
    assert.ok(gulpfile.indexOf(hook) < gulpfile.lastIndexOf('build.initialize('), project);
    const config = JSON.parse(fs.readFileSync(path.join(root, project, 'config', 'package-solution.json'), 'utf8'));
    const filename = path.basename(config.paths.zippedPackage).toLowerCase();
    assert.ok(!filenames.has(filename), 'Central package filename collision: ' + filename);
    filenames.add(filename);
  }
});

test('copy preserves original bytes and atomically replaces only the matching external package', t => {
  const f = fixture(t);
  assert.equal(attachPackageCopy.copyPackage(f.project), f.destination);
  assert.deepEqual(fs.readFileSync(f.destination), fs.readFileSync(f.source));
  fs.writeFileSync(path.join(f.directory, 'another.sppkg'), 'unrelated package');
  fs.writeFileSync(f.source, 'new package');
  attachPackageCopy.copyPackage(f.project);
  assert.equal(fs.readFileSync(f.destination, 'utf8'), 'new package');
  assert.equal(fs.readFileSync(f.source, 'utf8'), 'new package');
  assert.equal(fs.readFileSync(path.join(f.directory, 'another.sppkg'), 'utf8'), 'unrelated package');
  assert.deepEqual(fs.readdirSync(f.directory).sort(), ['another.sppkg', 'example.sppkg']);
  assert.ok(!f.destination.startsWith(f.repository + path.sep));
});

test('honors configured package directory and filename rather than the working directory', t => {
  const f = fixture(t);
  fs.mkdirSync(path.join(f.project, 'output'));
  fs.writeFileSync(f.config, '\uFEFF' + JSON.stringify({
    paths: { packageDir: 'output', zippedPackage: 'Different.sppkg' }
  }));
  fs.writeFileSync(path.join(f.project, 'output', 'Different.sppkg'), 'custom package');
  const result = attachPackageCopy.copyPackage(f.project);
  assert.equal(result, path.join(f.directory, 'Different.sppkg'));
  assert.equal(fs.readFileSync(result, 'utf8'), 'custom package');
});

test('wraps the executable, preserves metadata and copies only after successful production packaging', async t => {
  const f = fixture(t);
  const s = buildStub({ execute: async () => {
    s.events.push('package');
    assert.ok(!fs.existsSync(f.destination));
    fs.writeFileSync(f.source, 'fresh production package');
  } });
  attachPackageCopy(s.build, f.project);
  const tasks = s.build.rig.getTasks();
  assert.equal(tasks.get('package-solution').arguments, s.definition.arguments);
  assert.equal(tasks.get('build'), s.other);
  assert.equal(tasks.get('bundle'), s.other);
  assert.notEqual(tasks.get('package-solution').executable, s.definition.executable);
  await tasks.get('package-solution').executable.execute({ production: true });
  assert.deepEqual(s.events.slice(0, 2), ['package', 'collect-solution-package']);
  assert.equal(fs.readFileSync(f.destination, 'utf8'), 'fresh production package');
  assert.ok(s.events.some(message => message.includes(f.destination)));
});

test('debug packaging, ordinary build/bundle and disabled packaging never refresh stale copies', async t => {
  const f = fixture(t);
  attachPackageCopy.copyPackage(f.project);
  const previous = fs.readFileSync(f.destination);
  const s = buildStub({ execute: async () => fs.writeFileSync(f.source, 'debug output') });
  attachPackageCopy(s.build, f.project);
  const tasks = s.build.rig.getTasks();
  await tasks.get('package-solution').executable.execute({ production: false });
  await tasks.get('build').executable.execute({ production: true });
  await tasks.get('bundle').executable.execute({ production: true });
  assert.deepEqual(fs.readFileSync(f.destination), previous);
  const disabled = buildStub({ isEnabled: () => false, execute: async () => assert.fail('disabled task ran') });
  attachPackageCopy(disabled.build, f.project);
  await disabled.build.rig.getTasks().get('package-solution').executable.execute({ production: true });
  assert.deepEqual(fs.readFileSync(f.destination), previous);
});

test('failed packaging preserves the previous collection and propagates the original error', async t => {
  const f = fixture(t);
  attachPackageCopy.copyPackage(f.project);
  const previous = fs.readFileSync(f.destination);
  fs.writeFileSync(f.source, 'stale package must not be copied');
  const failure = new Error('packaging failed');
  const s = buildStub({ execute: async () => { throw failure; } });
  attachPackageCopy(s.build, f.project);
  await assert.rejects(s.build.rig.getTasks().get('package-solution').executable.execute({ production: true }),
    error => error === failure);
  assert.deepEqual(fs.readFileSync(f.destination), previous);
  assert.ok(!s.events.includes('collect-solution-package'));
});

test('missing, empty, directory and invalid configured sources fail without replacing the prior copy', t => {
  const f = fixture(t);
  attachPackageCopy.copyPackage(f.project);
  const previous = fs.readFileSync(f.destination);
  fs.unlinkSync(f.source);
  assert.throws(() => attachPackageCopy.copyPackage(f.project), /ENOENT/);
  fs.writeFileSync(f.source, '');
  assert.throws(() => attachPackageCopy.copyPackage(f.project), /nonempty file/);
  fs.unlinkSync(f.source);
  fs.mkdirSync(f.source);
  assert.throws(() => attachPackageCopy.copyPackage(f.project), /nonempty file/);
  for (const invalid of ['{', '{}', '{"paths":{"zippedPackage":""}}',
    '{"paths":{"zippedPackage":"../../../outside.sppkg"}}',
    '{"paths":{"zippedPackage":"solution/not-a-package.txt"}}']) {
    fs.writeFileSync(f.config, invalid);
    assert.throws(() => attachPackageCopy.copyPackage(f.project));
    assert.deepEqual(fs.readFileSync(f.destination), previous);
  }
});

test('partial copy failures preserve the previous package, remove staging files and fail the task', async t => {
  const f = fixture(t);
  attachPackageCopy.copyPackage(f.project);
  const previous = fs.readFileSync(f.destination);
  const failure = new Error('simulated copy failure');
  t.mock.method(fs, 'copyFileSync', (source, temporary) => {
    fs.writeFileSync(temporary, 'incomplete');
    throw failure;
  });
  const s = buildStub({ execute: async () => {} });
  attachPackageCopy(s.build, f.project);
  await assert.rejects(s.build.rig.getTasks().get('package-solution').executable.execute({ production: true }),
    error => error === failure);
  assert.deepEqual(fs.readFileSync(f.destination), previous);
  assert.deepEqual(fs.readdirSync(f.directory), ['example.sppkg']);
});

test('replacement failures preserve the previous package and remove staging files', t => {
  const f = fixture(t);
  attachPackageCopy.copyPackage(f.project);
  const previous = fs.readFileSync(f.destination);
  const failure = new Error('simulated replacement failure');
  t.mock.method(fs, 'renameSync', () => { throw failure; });
  assert.throws(() => attachPackageCopy.copyPackage(f.project), error => error === failure);
  assert.deepEqual(fs.readFileSync(f.destination), previous);
  assert.deepEqual(fs.readdirSync(f.directory), ['example.sppkg']);
});

test('blocked destination and unavailable packaging task fail explicitly', t => {
  const f = fixture(t);
  fs.writeFileSync(f.directory, 'not a directory');
  assert.throws(() => attachPackageCopy.copyPackage(f.project), /EEXIST|ENOTDIR/);
  const s = buildStub({ execute: async () => {} });
  s.build.rig.getTasks = () => new Map();
  attachPackageCopy(s.build, f.project);
  assert.throws(() => s.build.rig.getTasks(), /package-solution task is unavailable/);
});

test('manual collection discovers independent projects and explicitly labels existing packages', t => {
  const f = fixture(t);
  fs.mkdirSync(path.join(f.repository, 'not-a-project'));
  const messages = [];
  t.mock.method(console, 'log', message => messages.push(message));
  assert.deepEqual(attachPackageCopy.collectExistingPackages(f.repository), [f.destination]);
  assert.deepEqual(fs.readFileSync(f.destination), fs.readFileSync(f.source));
  assert.ok(messages[0].includes('Copied existing Example package'));
});
