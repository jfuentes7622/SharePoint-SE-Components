'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

function parseJson(text, name) {
  let output = '';
  let quoted = false;
  let escaped = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      output += c;
      if (escaped) escaped = false;
      else if (c === '\\') escaped = true;
      else if (c === '"') quoted = false;
    } else if (c === '"') {
      quoted = true;
      output += c;
    } else if (c === '/' && text[i + 1] === '/') {
      while (i < text.length && text[i] !== '\n') i++;
      output += '\n';
    } else if (c === '/' && text[i + 1] === '*') {
      const end = text.indexOf('*/', i + 2);
      if (end < 0) throw new Error(`${name}: unterminated JSON comment`);
      output += ' ';
      i = end + 1;
    } else {
      output += c;
    }
  }
  try {
    return JSON.parse(output.replace(/^\uFEFF/, ''));
  } catch (error) {
    throw new Error(`${name}: ${error.message}`);
  }
}

function filesUnder(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap(entry => {
    if (['node_modules', 'node_modules_old', '.git'].includes(entry.name)) return [];
    const file = path.join(dir, entry.name);
    return entry.isDirectory() ? filesUnder(file) : entry.isFile() ? [file] : [];
  });
}

function projects(root) {
  return fs.readdirSync(root, { withFileTypes: true })
    .filter(entry => entry.isDirectory() &&
      fs.existsSync(path.join(root, entry.name, 'package.json')) &&
      fs.existsSync(path.join(root, entry.name, 'config', 'package-solution.json')))
    .map(entry => entry.name).sort();
}

function git(root, args, optional = false) {
  const result = spawnSync('git', args, { cwd: root, encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    if (optional) return undefined;
    throw new Error(`git ${args[0]} failed: ${result.stderr.trim()}`);
  }
  return result.stdout;
}

function gitPath(file) {
  return file.split(path.sep).join('/');
}

function versionParts(value, size) {
  return typeof value === 'string' && new RegExp(`^\\d+(?:\\.\\d+){${size - 1}}$`).test(value)
    ? value.split('.').map(Number) : undefined;
}

function regressed(current, previous) {
  for (let i = 0; i < current.length; i++) {
    if (current[i] !== previous[i]) return current[i] < previous[i];
  }
  return false;
}

function prohibited(text) {
  const patterns = [
    /['"](?:@microsoft\/(?:mgt[^'"]*|microsoft-graph[^'"]*|sp-adaptive-card-extension-base|teams-js)|@microsoftgraph\/[^'"]*|@fluentui\/[^'"]*|@pnp\/graph(?:\/[^'"]*)?)['"]/gi,
    /(?:https:\/\/graph\.microsoft\.com|\/_api\/v2\.0\/|\bMSGraphClient(?:V3)?\b|\bGraphHttpClient\b|\bmsGraphClientFactory\b)/gi
  ];
  return patterns.flatMap(pattern => text.match(pattern) || []).map(value => value.toLowerCase());
}

function validate(root, base) {
  const errors = [];
  const check = (condition, message) => { if (!condition) errors.push(message); };
  const names = projects(root);
  check(names.length > 0, 'No SPFx projects found.');
  const ids = { solution: new Map(), feature: new Map(), component: new Map() };
  const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  function id(kind, value, file) {
    check(typeof value === 'string' && uuid.test(value), `${file}: invalid ${kind} ID`);
    if (typeof value !== 'string') return;
    const key = value.toLowerCase();
    check(!ids[kind].has(key), `${file}: duplicate ${kind} ID also in ${ids[kind].get(key)}`);
    ids[kind].set(key, file);
  }
  const read = relative => parseJson(fs.readFileSync(path.join(root, relative), 'utf8'), relative);
  const before = relative => {
    if (!base) return undefined;
    const text = git(root, ['show', `${base}:${gitPath(relative)}`], true);
    return text === undefined ? undefined : parseJson(text, `base:${relative}`);
  };
  function version(value, old, size, file, wildcard = false) {
    if (wildcard && value === '*') return;
    const parts = versionParts(value, size);
    check(!!parts, `${file}: invalid version ${value}`);
    if (old !== undefined && old !== '*') {
      const prior = versionParts(old, size);
      check(!!prior, `${file}: invalid base version ${old}`);
      if (parts && prior) check(!regressed(parts, prior), `${file}: version regressed from ${old} to ${value}`);
    }
  }
  if (base) {
    git(root, ['rev-parse', '--verify', `${base}^{commit}`]);
    const dirs = git(root, ['ls-tree', '--name-only', base]).trim().split('\n');
    for (const dir of dirs) {
      if (!git(root, ['show', `${base}:${dir}/config/package-solution.json`], true)) continue;
      check(names.includes(dir), `${dir}: existing project removed or renamed; IDs must be preserved`);
      const tracked = git(root, ['ls-tree', '-r', '--name-only', base, '--', `${dir}/src`]).trim().split('\n');
      for (const file of tracked.filter(file => file.endsWith('.manifest.json'))) {
        check(fs.existsSync(path.join(root, file)), `${file}: existing component manifest removed or renamed`);
      }
    }
  }
  for (const project of names) {
    const packageFile = path.join(project, 'package.json');
    const pkg = read(packageFile);
    const oldPkg = before(packageFile);
    version(pkg.version, oldPkg && oldPkg.version, 3, packageFile);
    for (const section of ['dependencies', 'devDependencies']) {
      for (const [name, value] of Object.entries(pkg[section] || {})) {
        if (name.startsWith('@microsoft/sp-')) {
          check(['1.5.1', '~1.5.1', '^1.5.1'].includes(value), `${packageFile}: ${name} must stay on 1.5.1 (found ${value})`);
        }
      }
    }
    const lockFile = path.join(project, 'package-lock.json');
    const lock = read(lockFile);
    check(lock.lockfileVersion === 1, `${lockFile}: preserve the legacy v1 lockfile`);
    for (const name of Object.keys({ ...pkg.dependencies, ...pkg.devDependencies })) {
      if (name.startsWith('@microsoft/sp-')) {
        check(lock.dependencies && lock.dependencies[name] && lock.dependencies[name].version === '1.5.1',
          `${lockFile}: direct dependency ${name} must resolve to 1.5.1`);
      }
    }
    function checkLocked(dependencies) {
      for (const [name, dependency] of Object.entries(dependencies || {})) {
        if (name.startsWith('@microsoft/sp-')) {
          const parts = versionParts(dependency.version, 3);
          check(parts && !regressed([1, 5, 1], parts),
            `${lockFile}: ${name} resolved to unsupported ${dependency.version}`);
        }
        checkLocked(dependency.dependencies);
      }
    }
    checkLocked(lock.dependencies);
    const solutionFile = path.join(project, 'config', 'package-solution.json');
    const config = read(solutionFile);
    const oldConfig = before(solutionFile);
    const solution = config.solution;
    check(!!solution, `${solutionFile}: missing solution`);
    if (!solution) continue;
    id('solution', solution.id, solutionFile);
    version(solution.version, oldConfig && oldConfig.solution.version, 4, solutionFile);
    if (oldConfig) {
      check(solution.id === oldConfig.solution.id, `${solutionFile}: solution ID changed`);
      check(solution.skipFeatureDeployment === oldConfig.solution.skipFeatureDeployment, `${solutionFile}: skipFeatureDeployment changed`);
      check(solution.isDomainIsolated === oldConfig.solution.isDomainIsolated, `${solutionFile}: deployment isolation changed`);
    }
    check(solution.skipFeatureDeployment === undefined || typeof solution.skipFeatureDeployment === 'boolean',
      `${solutionFile}: skipFeatureDeployment must be boolean`);
    const features = solution.features || [];
    for (const feature of features) {
      id('feature', feature.id, solutionFile);
      const previous = oldConfig && (oldConfig.solution.features || []).find(item => item.id === feature.id);
      version(feature.version, previous && previous.version, 4, `${solutionFile}: feature ${feature.id}`);
    }
    if (oldConfig) {
      for (const feature of oldConfig.solution.features || []) {
        check(features.some(item => item.id === feature.id), `${solutionFile}: existing feature ID removed or changed: ${feature.id}`);
      }
    }
    check(!!config.paths && typeof config.paths.zippedPackage === 'string' && config.paths.zippedPackage.endsWith('.sppkg'),
      `${solutionFile}: missing .sppkg output path`);
    const source = filesUnder(path.join(root, project, 'src'));
    const manifests = source.filter(file => file.endsWith('.manifest.json'));
    check(manifests.length > 0, `${project}: no source component manifests`);
    for (const file of manifests) {
      const relative = path.relative(root, file);
      const manifest = read(relative);
      const old = before(relative);
      id('component', manifest.id, relative);
      version(manifest.version, old && old.version, 3, relative, true);
      check(['WebPart', 'Extension'].includes(manifest.componentType), `${relative}: unsupported componentType`);
      if (old) check(manifest.id === old.id, `${relative}: component ID changed`);
    }
    const build = read(path.join(project, 'config', 'config.json'));
    check(!!build.bundles && Object.keys(build.bundles).length > 0, `${project}: no configured bundles`);
    for (const bundle of Object.values(build.bundles || {})) {
      for (const component of bundle.components || []) {
        check(typeof component.entrypoint === 'string', `${project}: missing bundle entrypoint`);
        check(typeof component.manifest === 'string' &&
          manifests.includes(path.resolve(root, project, component.manifest || '')),
        `${project}: bundle references a missing source manifest: ${component.manifest}`);
      }
    }
    const serve = read(path.join(project, 'config', 'serve.json'));
    check(Number.isInteger(serve.port) && serve.port > 0 && serve.port < 65536, `${project}: invalid serve port`);
    const write = read(path.join(project, 'config', 'write-manifests.json'));
    check(typeof write.cdnBasePath === 'string' && write.cdnBasePath.length > 0, `${project}: missing cdnBasePath`);
    for (const file of [path.join(root, packageFile), ...source.filter(file => /\.(tsx?|jsx?)$/.test(file))]) {
      const relative = path.relative(root, file);
      const matches = prohibited(fs.readFileSync(file, 'utf8'));
      const priorText = base ? git(root, ['show', `${base}:${gitPath(relative)}`], true) : undefined;
      const prior = prohibited(priorText || '');
      for (const match of matches) {
        const index = prior.indexOf(match);
        if (index >= 0) prior.splice(index, 1);
        else errors.push(`${relative}: prohibited new API/dependency pattern: ${match}`);
      }
    }
  }
  return { projects: names, errors };
}

if (require.main === module) {
  try {
    const args = process.argv.slice(2);
    if (args.length && (args.length !== 2 || args[0] !== '--base' || !/^[0-9a-f]{7,40}$/i.test(args[1]))) {
      throw new Error('Usage: npm run validate-se -- [--base <commit-sha>]');
    }
    const result = validate(path.resolve(__dirname, '..'), args[1]);
    result.errors.forEach(error => console.error(`FAIL: ${error}`));
    if (result.errors.length) process.exitCode = 1;
    else console.log(`SE static validation passed for ${result.projects.length} projects (SPFx 1.5.1). Package/build/deployment checks are separate.`);
  } catch (error) {
    console.error(`FAIL: ${error.message}`);
    process.exitCode = 1;
  }
}

module.exports = { parseJson, projects, prohibited, regressed, validate };
