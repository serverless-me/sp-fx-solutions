'use strict';

let gulp = require('gulp');
let build = require('@ms/ms-core-build');

build.doBefore('bundle', 'collect-manifests');

build.initializeTasks(
  require('gulp'),
  {
    build: require('./config/build.json'),
    bundle: require('./config/bundle.json'),
    serve: require('./config/serve.json'),
    'package-solution': require('./config/package-solution.json'),
    'upload-cdn': require('./config/upload-cdn.json')
  }
);
