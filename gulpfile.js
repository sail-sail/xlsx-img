const gulp = require("gulp");
const ts = require('gulp-typescript');
const uglify = require('gulp-uglify-es').default;
const sourcemaps = require('gulp-sourcemaps');

const project = "xlsx-img";

const dist = `./build/${project}`;
//----------------------------------------------------------------------------------------------------- config
gulp.task("config", function() {
  return gulp.src(["package.json"], { allowEmpty: true })
    .pipe(gulp.dest(`${dist}`));
});
//----------------------------------------------------------------------------------------------------- ts
gulp.task("ts", function() {
  const tsProject = ts.createProject('tsconfig.json', { sourceMap: false });
  return tsProject.src()
    .pipe(sourcemaps.init())
    .pipe(tsProject()).js
    .pipe(uglify({toplevel: false}))
    .on('error', function (err) {
      console.error(err);
    })
    .pipe(sourcemaps.write(`./`))
    .pipe(gulp.dest(`${dist}/src`));
});

gulp.task("default", gulp.series("config","ts"));