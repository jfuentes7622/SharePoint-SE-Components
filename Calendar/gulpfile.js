'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// FullCalendar's npm packages ship modern (ES2017+) JavaScript, and this
// project bundles all of its code (including node_modules) into a single
// per-web-part file. The project's ES5-only UglifyJsPlugin (webpack 2.x)
// cannot parse that modern syntax during --ship (minified) builds, so
// minification is skipped for this bundle (the output is still a valid
// --ship/production build, just not minified).
build.configureWebpack.setConfig({
  additionalConfiguration: (generatedConfiguration) => {
    if (generatedConfiguration.plugins) {
      generatedConfiguration.plugins = generatedConfiguration.plugins.filter((plugin) => {
        return !(plugin && plugin.constructor && plugin.constructor.name === 'UglifyJsPlugin');
      });
    }
    return generatedConfiguration;
  }
});

build.initialize(gulp);
