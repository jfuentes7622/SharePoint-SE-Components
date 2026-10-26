'use strict';

const build = require('@microsoft/sp-build-web');
const path = require('path');
const webpack = require('webpack');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

    build.tslint.setConfig({
      lintConfig: require('./tslint.json'),
      displayAsWarning: false
    });


build.configureWebpack.setConfig({
  additionalConfiguration: (generatedConfiguration) => {
    const propertyControlsRoot = path.dirname(require.resolve('@pnp/spfx-property-controls/package.json'));
    generatedConfiguration.plugins = generatedConfiguration.plugins || [];
    // PnP 1.1.0 starts the legacy AI script loader at import time, before runtime opt-out is possible.
    generatedConfiguration.plugins.push(new webpack.NormalModuleReplacementPlugin(
      /[\\/]common[\\/]appInsights$/,
      (resource) => {
        if (resource.context.indexOf(propertyControlsRoot + path.sep) === 0) {
          resource.request = path.resolve(__dirname, 'lib/webparts/shared/pnpTelemetryNoop.js');
        }
      }
    ));
    return generatedConfiguration;
  }
});

build.initialize(require('gulp'));
