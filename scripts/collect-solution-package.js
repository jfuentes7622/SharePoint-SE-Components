'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function copyPackage(projectDirectory) {
  const project = path.resolve(projectDirectory);
  const configPath = path.join(project, 'config', 'package-solution.json');
  const config = JSON.parse(fs.readFileSync(configPath, 'utf8').replace(/^\uFEFF/, ''));
  if (!config.paths || typeof config.paths.zippedPackage !== 'string' || !config.paths.zippedPackage.trim()) {
    throw new Error(configPath + ': paths.zippedPackage must name the generated solution package.');
  }
  const source = path.resolve(project, config.paths.packageDir || 'sharepoint', config.paths.zippedPackage);
  const relative = path.relative(project, source);
  if (path.isAbsolute(relative) || relative === '..' || relative.indexOf('..' + path.sep) === 0
    || path.extname(source).toLowerCase() !== '.sppkg') {
    throw new Error('Solution package must be an .sppkg inside its project: ' + source);
  }
  const stat = fs.statSync(source);
  if (!stat.isFile() || !stat.size) { throw new Error('Solution package is not a nonempty file: ' + source); }

  const directory = path.resolve(project, '..', '..', 'SharePoint-SE-Packages');
  fs.mkdirSync(directory, { recursive: true });
  const destination = path.join(directory, path.basename(source));
  const temporary = destination + '.' + process.pid + '.' + crypto.randomBytes(8).toString('hex') + '.tmp';
  try {
    fs.copyFileSync(source, temporary, fs.constants.COPYFILE_EXCL);
    // Replace only after the new copy is complete; preserve the previous package on failure.
    fs.renameSync(temporary, destination);
  } finally {
    if (fs.existsSync(temporary)) { fs.unlinkSync(temporary); }
  }
  return destination;
}

function attachPackageCopy(build, projectDirectory) {
  const getTasks = build.rig.getTasks;
  build.rig.getTasks = function() {
    const tasks = getTasks.call(this);
    const definition = tasks.get('package-solution');
    if (!definition || !definition.executable) {
      throw new Error('Cannot attach central package collection: package-solution task is unavailable.');
    }
    const original = definition.executable;
    const copyTask = build.subTask('collect-solution-package', function(gulp, buildOptions, done) {
      if (!buildOptions.production || (original.isEnabled && !original.isEnabled(buildOptions))) {
        this.log('Central package copy skipped: production packaging is required.');
        done();
        return;
      }
      try {
        const destination = copyPackage(projectDirectory);
        this.log('Copied solution package to ' + destination);
        done();
      } catch (error) {
        done(error);
      }
    });
    tasks.set('package-solution', Object.assign({}, definition, {
      executable: build.serial(original, copyTask)
    }));
    return tasks;
  };
}

function collectExistingPackages(repositoryDirectory) {
  const root = path.resolve(repositoryDirectory);
  return fs.readdirSync(root).sort().filter(function(name) {
    const project = path.join(root, name);
    return fs.statSync(project).isDirectory()
      && fs.existsSync(path.join(project, 'package.json'))
      && fs.existsSync(path.join(project, 'config', 'package-solution.json'));
  }).map(function(name) {
    const destination = copyPackage(path.join(root, name));
    console.log('Copied existing ' + name + ' package to ' + destination);
    return destination;
  });
}

module.exports = attachPackageCopy;
module.exports.copyPackage = copyPackage;
module.exports.collectExistingPackages = collectExistingPackages;

if (require.main === module) {
  collectExistingPackages(path.resolve(__dirname, '..'));
}
