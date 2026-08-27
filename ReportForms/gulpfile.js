'use strict';

const build = require('@microsoft/sp-build-web');

// Disable TSLint so legacy style warnings do not fail the ship build on stderr.
build.tslint.enabled = false;

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.initialize(require('gulp'));