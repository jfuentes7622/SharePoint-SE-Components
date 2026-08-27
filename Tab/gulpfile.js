'use strict';

const build = require('@microsoft/sp-build-web');
const path = require('path');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// var getTasks = build.rig.getTasks;
// build.rig.getTasks = function () {
//   var result = getTasks.call(build.rig);

//   result.set('serve', result.get('serve-deprecated'));

//   return result;
// };

// // This section is inspired by Stefan Bauer's article at https://n8d.at/how-to-version-new-sharepoint-framework-projects/
// // Stefan rocks!
// let syncVersionsSubtask = build.subTask('version-sync', function (gulp, buildOptions, done) {
//   this.log('Synching versions');
  
//   // import gulp utilits to write error messages
//   const gutil = require('gulp-util');
  
//   // import file system utilities form nodeJS
//   const fs = require('fs');
  
//   // read package.json
//   var pkgConfig = require('./package.json');
  
//   // read configuration of web part solution file
//   var pkgSolution = require('./config/package-solution.json');
  
//   // log old version
//   this.log('Current package-solution.json version:\t' + pkgSolution.solution.version);
  
//   // Generate new MS compliant version number
//   var newVersionNumber = pkgConfig.version.split('-')[0] + '.0';
  
//   if (pkgSolution.solution.version !== newVersionNumber) {
//   // assign newly generated version number to web part version
//   pkgSolution.solution.version = newVersionNumber;
  
//   // log new version
//   this.log('New package-solution.json version:\t' + pkgSolution.solution.version);
  
//   // write changed package-solution file
//    fs.writeFile('./config/package-solution.json', JSON.stringify(pkgSolution, null, 4), function (err, result) {
//           if (err) this.log('error', err);
//       });
//   }
//   else {
//   this.log('package-solution.json version is up-to-date');
//   }
  
//   done();
//   }); 
//   let syncVersionTask = build.task('version-sync', syncVersionsSubtask);
//   build.rig.addPreBuildTask(syncVersionTask);

build.tslint.setConfig({
      lintConfig: require('./tslint.json'),
      displayAsWarning: false
    });

build.writeManifests.setConfig({
  cumulativeManifestOptions: {
    ignoreOutputManifestIds: ['d688e552-a2fb-4904-af1c-c28aa1ee79d3']
  }
});

build.configureWebpack.setConfig({
  additionalConfiguration: (generatedConfiguration) => {
    generatedConfiguration.resolve = generatedConfiguration.resolve || {};
    generatedConfiguration.resolve.alias = generatedConfiguration.resolve.alias || {};
    generatedConfiguration.resolve.alias['@pnp/telemetry-js'] = path.resolve(
      __dirname,
      'lib/webparts/shared/pnpTelemetryNoop.js'
    );
    return generatedConfiguration;
  }
});

build.initialize(require('gulp'));
