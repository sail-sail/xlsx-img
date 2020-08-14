const { readFileSync, statSync, writeFileSync } = require("fs");

function exec(file) {
  try {
    const stats = statSync(file);
    if(!stats.isFile()) return;
  } catch (err) {
    return;
  }
  let str = readFileSync(file, "utf8");
  str = str.replace(".asNodeBuffer().toString('binary')", ".asNodeBuffer().toString()");
  writeFileSync(file, str);
}
exec(`${ __dirname }/node_modules/xlsx/xlsx.js`);
exec(`${ __dirname }/../../xlsx/xlsx.js`);