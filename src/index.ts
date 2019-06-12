import { inflateRaw } from "zlib";
import xlsx from "node-xlsx";
import * as Hzip from "hzip";
import { parseString } from "xml2js";
import { basename, dirname } from "path";
import { normalize } from "path";
import * as moment from "moment";

/**
 * 解析Excel
 * @param mixed 
 * @param _options 
 */
export async function parseXlsx(mixed: Buffer, options?: {}): Promise<{
  name: string,
  data: Array<Array<any>>,
}[]> {
  const rvObj = await xlsx.parse(mixed, options);
  const hzip = new Hzip(mixed);
  const workbookEntity = hzip.getEntry(`xl/workbook.xml`);
  const workbookBuf = await new Promise((resolve, reject) => {
    inflateRaw(workbookEntity.cfile,(err, buf) => {
      if(err) {
        reject(err);
        return;
      }
      resolve(buf);
    });
  });
  const workbook = await new Promise((resolve, reject) => {
    parseString(workbookBuf, function (err, result) {
      if(err) {
        reject(err);
        return;
      }
      resolve(result);
    });
  });
  const workbookRelsEntity = hzip.getEntry(`xl/_rels/workbook.xml.rels`);
  const workbookRelsBuf = await new Promise((resolve, reject) => {
    inflateRaw(workbookRelsEntity.cfile,(err, buf) => {
      if(err) {
        reject(err);
        return;
      }
      resolve(buf);
    });
  });
  const workbookRels = await new Promise((resolve, reject) => {
    parseString(workbookRelsBuf, function (err, result) {
      if(err) {
        reject(err);
        return;
      }
      resolve(result);
    });
  });
  const sheets = workbook["workbook"]["sheets"][0].sheet;
  for (let o = 0; o < sheets.length; o++) {
    const wbSheet = sheets[o]["$"];
    const sheetName = wbSheet["name"];
    const wbRId = wbSheet["r:id"];
    const wBrelationships = workbookRels["Relationships"]["Relationship"];
    for (let o2 = 0; o2 < wBrelationships.length; o2++) {
      const wBrelationship = wBrelationships[o2]["$"];
      if(wBrelationship["Id"] !== wbRId) continue;
      const target = normalize(`xl/${ wBrelationship["Target"] }`).replace(/\\/gm, "/");
      const sheetEntity = hzip.getEntry(target);
      const sheetBuf = await new Promise((resolve, reject) => {
        inflateRaw(sheetEntity.cfile,(err, buf) => {
          if(err) {
            reject(err);
            return;
          }
          resolve(buf);
        });
      });
      const sheet = await new Promise((resolve, reject) => {
        parseString(sheetBuf, function (err, result) {
          if(err) {
            reject(err);
            return;
          }
          resolve(result);
        });
      });
      if(!sheet["worksheet"]["drawing"]) continue;
      const rId = sheet["worksheet"]["drawing"][0]["$"]["r:id"];
      const sheetRelsEntity = hzip.getEntry(`${ dirname(sheetEntity.fileName) }/_rels/${ basename(sheetEntity.fileName) }.rels`);
      if(!sheetRelsEntity) continue;
      const sheetRelsBuf = await new Promise((resolve, reject) => {
        inflateRaw(sheetRelsEntity.cfile,(err, buf) => {
          if(err) {
            reject(err);
            return;
          }
          resolve(buf);
        });
      });
      const sheetRels = await new Promise((resolve, reject) => {
        parseString(sheetRelsBuf, function (err, result) {
          if(err) {
            reject(err);
            return;
          }
          resolve(result);
        });
      });
      const relationships2 = sheetRels["Relationships"]["Relationship"];
      for (let m = 0; m < relationships2.length; m++) {
        const relationship = relationships2[m]["$"];
        if(relationship["Id"] !== rId) continue;
        const target = normalize(`xl/worksheets/${ relationship["Target"] }`).replace(/\\/gm, "/");
        const drawing1Entry = hzip.getEntry(target);
        const fileName = drawing1Entry.fileName;
        const drawing1RelsEntry = hzip.getEntry(`${ dirname(fileName) }/_rels/${ basename(fileName) }.rels`);
        const drawing1Buf = await new Promise((resolve, reject) => {
          inflateRaw(drawing1Entry.cfile,(err, buf) => {
            if(err) {
              reject(err);
              return;
            }
            resolve(buf);
          });
        });
        const drawing1 = await new Promise((resolve, reject) => {
          parseString(drawing1Buf, function (err, result) {
            if(err) {
              reject(err);
              return;
            }
            resolve(result);
          });
        });
        const drawing1RelsBuf = await new Promise((resolve, reject) => {
          inflateRaw(drawing1RelsEntry.cfile,(err, buf) => {
            if(err) {
              reject(err);
              return;
            }
            resolve(buf);
          });
        });
        const drawing1Rels = await new Promise((resolve, reject) => {
          parseString(drawing1RelsBuf, function (err, result) {
            if(err) {
              reject(err);
              return;
            }
            resolve(result);
          });
        });
        const relationships = drawing1Rels["Relationships"]["Relationship"];
        const xdrTwoCellAnchors = drawing1["xdr:wsDr"]["xdr:twoCellAnchor"];
        for (let i = 0; i < xdrTwoCellAnchors.length; i++) {
          const xdrTwoCellAnchor = xdrTwoCellAnchors[i];
          const xdrFrom = xdrTwoCellAnchor["xdr:from"];
          const xdrRowNum = Number(xdrFrom[0]["xdr:row"][0]);
          const xdrColNum = Number(xdrFrom[0]["xdr:col"][0]);
          if(!xdrTwoCellAnchor["xdr:pic"]) continue;
          const xdrPic = xdrTwoCellAnchor["xdr:pic"][0];
          const xdrBlipFill = xdrPic["xdr:blipFill"][0];
          const aBlip = xdrBlipFill["a:blip"][0];
          const rEmbed = aBlip["$"]["r:embed"];
          for (let k = 0; k < relationships.length; k++) {
            const relationship = relationships[k]["$"];
            if(relationship["Id"] !== rEmbed) continue;
            const target = normalize(`xl/drawings/${ relationship["Target"] }`).replace(/\\/gm, "/");
            const targetEntity = hzip.getEntry(target);
            const targetBuf = targetEntity.cfile;
            // console.log({ sheetName, xdrRowNum, xdrColNum, target });
            for (let i2 = 0; i2 < rvObj.length; i2++) {
              const dataObj = rvObj[i2];
              if(dataObj.name !== sheetName) continue;
              (<any>dataObj.data[xdrRowNum][xdrColNum]) = targetBuf;
            }
          }
        }
      }
    }
  }
  return rvObj;
};

export function xlsx2Date(val: number | string | Date): Date {
  let date :Date = undefined;
  if(typeof val === "string") {
    date = moment(val).toDate();
  } else if(val instanceof Date) {
    return val;
  } else {
    let valTmp :number = <number>val;
    if(valTmp < 60) valTmp--;
    else if(valTmp > 60) {
      valTmp -= 2;
    }
    if(valTmp < 0 || valTmp === 60) {
      date = undefined;
    } else {
      valTmp = Math.round(valTmp * 86400000) - 2209017600000;
      date = new Date(valTmp);
    }
  }
  if(isNaN(<any>date)) {
    date = undefined;
  }
  console.log(date)
  return date;
}