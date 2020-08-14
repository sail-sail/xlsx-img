const gulp = require("gulp");
const ts = require('gulp-typescript');
const uglify = require('gulp-uglify-es').default;
const sourcemaps = require('gulp-sourcemaps');
const merge2 = require('merge2');

const project = "xlsx-img";

const dist = `./build/${project}`;
//----------------------------------------------------------------------------------------------------- config
gulp.task("js", function() {
  return gulp.src(["src/*.js"], { allowEmpty: true })
    .pipe(gulp.dest(`${dist}/src/`));
});
gulp.task("config", function() {
  return gulp.src(["package.json"], { allowEmpty: true })
    .pipe(gulp.dest(`${dist}`));
});
//----------------------------------------------------------------------------------------------------- ts
gulp.task("ts", function() {
  const tsProject = ts.createProject('tsconfig.json');
  const tsResult = tsProject.src()
  .pipe(sourcemaps.init())
  .pipe(tsProject());
  
  return merge2([
    tsResult.js.pipe(uglify({toplevel: false}))
    .on('error', function (err) {
      console.error(err);
    })
    .pipe(sourcemaps.write(`./`))
    .pipe(gulp.dest(`${dist}/src`)),
    
    tsResult.dts.pipe(gulp.dest(`${dist}/src`)),
  ]);
});

gulp.task("default", gulp.series("js","config","ts"));