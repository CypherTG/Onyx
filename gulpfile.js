'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Bypass the NodeJS version check so SPFx builds work on Node 22+
const originalRig = build.rig;
if (originalRig && originalRig.nodeSupportedVersionRange) {
  originalRig.nodeSupportedVersionRange = '>=12.13.0 <13.0.0 || >=14.15.0 <15.0.0 || >=16.13.0 <17.0.0 || >=18.17.1 <23.0.0';
}

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
