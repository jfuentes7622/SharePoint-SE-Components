'use strict';

const build = require('@microsoft/sp-build-web');

build.tslint.setConfig({
  lintConfig: require('./tslint.json'),
  displayAsWarning: false
});

build.initialize(require('gulp'));
