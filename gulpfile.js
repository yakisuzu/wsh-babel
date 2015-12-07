var gulp = require('gulp');
var webpack = require('gulp-webpack');
var convertEncoding = require('gulp-convert-encoding');

gulp.task('webpack', function(){
  return gulp.src('src/*.js')
  .pipe(webpack(require('./webpack.config.js')))
  .pipe(convertEncoding({to: 'Shift_JIS'}))
  .pipe(gulp.dest('dest/'));
});

gulp.task('watch', function(){
  gulp.watch(['src/*.js'], ['webpack']);
});

gulp.task('default', ['webpack']);
