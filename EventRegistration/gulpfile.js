'use strict';

const build = require('@microsoft/sp-build-web');
const fs = require('fs');

const syncVersionsSubtask = build.subTask('version-sync', function (gulp, buildOptions, done) {
  try {
    const packageConfig = JSON.parse(fs.readFileSync('./package.json', 'utf8'));
    const parts = String(packageConfig.version || '').split('-')[0].split('.');
    if (parts.length !== 3) {
      throw new Error('package.json version must contain major, minor, and patch numbers.');
    }
    const spfxVersion = parts[0] + '.' + parts[1] + '.' + parts[2] + '.1';
    const path = './config/package-solution.json';
    const config = JSON.parse(fs.readFileSync(path, 'utf8'));
    config.solution.version = spfxVersion;
    (config.solution.features || []).forEach(function (feature) {
      feature.version = spfxVersion;
    });
    fs.writeFileSync(path, JSON.stringify(config, null, 2) + '\n', 'utf8');
    this.log('Synchronized solution and feature versions to ' + spfxVersion + '.');
    done();
  } catch (error) {
    done(error);
  }
});

build.rig.addPreBuildTask(build.task('version-sync', syncVersionsSubtask));
build.tslint.enabled = false;
build.initialize(require('gulp'));
