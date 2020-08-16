const { readFileSync, statSync, writeFileSync } = require("fs");

function exec(file) {
  try {
    const stats = statSync(file);
    if(!stats.isFile()) return;
  } catch (err) {
    // console.error(err);
    return;
  }
  let str = readFileSync(file, "utf8");
  str = str.replace("return getdatastr(getzipfile(zip, file))", "return utf8read(getdatastr(getzipfile(zip, file)))");
  writeFileSync(file, str);
}
exec(`${ __dirname }/../node_modules/xlsx/xlsx.js`);
exec(`${ __dirname }/../../xlsx/xlsx.js`);