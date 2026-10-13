'use strict';

const build = require('@microsoft/sp-build-web');
const fs = require('fs');

const syncVersionsSubtask = build.subTask('version-sync', function (gulp, buildOptions, done) {
  try {
    const packageConfig = JSON.parse(fs.readFileSync('./package.json', 'utf8'));
    const packageVersion = String(packageConfig.version || '').split('-')[0];
    const versionParts = packageVersion.split('.');
    if (versionParts.length !== 3) {
      throw new Error('package.json version must contain major, minor, and patch numbers.');
    }
    const spfxVersion = versionParts[0] + '.' + versionParts[1] + '.0.' + versionParts[2];
    const solutionPath = './config/package-solution.json';
    const solutionConfig = JSON.parse(fs.readFileSync(solutionPath, 'utf8'));
    solutionConfig.solution.version = spfxVersion;
    const features = solutionConfig.solution.features || [];
    for (let featureIndex = 0; featureIndex < features.length; featureIndex += 1) {
      features[featureIndex].version = spfxVersion;
    }
    fs.writeFileSync(solutionPath, JSON.stringify(solutionConfig, null, 2) + '\n', 'utf8');
    this.log('Synchronized solution and feature versions to ' + spfxVersion + '.');
    done();
  } catch (error) {
    done(error);
  }
});

build.rig.addPreBuildTask(build.task('version-sync', syncVersionsSubtask));

build.tslint.enabled = false;

build.initialize(require('gulp'));