'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);


build.tslint.setConfig({
  lintConfig: require('./tslint.json'),
  displayAsWarning: false
});


require('../scripts/collect-solution-package')(build, __dirname);
build.initialize(require('gulp'));
