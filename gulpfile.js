'use strict';

const build = require('@microsoft/sp-build-web');
const gulp = require('gulp');

build.addSuppression(/Warning/gi);
build.addSuppression(/'icons\/request\.png'/gi);

build.initialize(gulp);

gulp.task('serve', gulp.series('serve-deprecated'));
