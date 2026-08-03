'use strict';

const build = require('@microsoft/sp-build-web');

build.tslint.enabled = false;

build.initialize(require('gulp'));