'use strict';

const build = require('@microsoft/sp-build-web');

// Suppress all SASS warnings (CSS class naming and non-module files)
build.addSuppression(/Warning - \[sass\]/gi);

// Disable linting to prevent warnings from causing build failures
// Re-enable once codebase is cleaned up
build.lintCmd.enabled = false;

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
