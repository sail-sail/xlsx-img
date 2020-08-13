import { inflateRaw } from "zlib";
import * as xlsx from "xlsx";
import * as Hzip from "hzip";
import { parseString } from "xml2js";
import { basename, dirname } from "path";
import { normalize } from "path";
import * as moment from "moment";

function safe_decode_range(range: any) {
	var o = {s:{c:0,r:0},e:{c:0,r:0}};
	var idx = 0, i = 0, cc = 0;
	var len = range.length;
	for(idx = 0; i < len; ++i) {
		if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
		idx = 26*idx + cc;
	}
	o.s.c = --idx;

	for(idx = 0; i < len; ++i) {
		if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
		idx = 10*idx + cc;
	}
	o.s.r = --idx;

	if(i === len || range.charCodeAt(++i) === 58) { o.e.c=o.s.c; o.e.r=o.s.r; return o; }

	for(idx = 0; i != len; ++i) {
		if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
		idx = 26*idx + cc;
	}
	o.e.c = --idx;

	for(idx = 0; i != len; ++i) {
		if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
		idx = 10*idx + cc;
	}
	o.e.r = --idx;
	return o;
}

function encode_row(row: number) { return "" + (row + 1); };
function encode_col(col: number) { if(col < 0) throw new Error("invalid column " + col); var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; };
var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
function datenum(v: Date, date1904?: boolean) {
	var epoch = v.getTime();
	if(date1904) epoch -= 1462*24*60*60*1000;
	var dnthresh = basedate.getTime() + (v.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
}
function safe_format_cell(cell: any, v: any) {
  var q = (cell.t == 'd' && v instanceof Date);
	if(cell.z != null) try { return (cell.w = xlsx.SSF.format(cell.z, q ? datenum(v) : v)); } catch(e) { }
	try { return (cell.w = xlsx.SSF.format((cell.XF||{}).numFmtId||(q ? 14 : 0),  q ? datenum(v) : v)); } catch(e) { return ''+v; }
}

function format_cell(cell: any, v: any, o: any) {
	if(cell == null || cell.t == null || cell.t == 'z') return "";
	if(cell.w !== undefined) return cell.w;
	if(cell.t == 'd' && !cell.z && o && o.dateNF) cell.z = o.dateNF;
	if(v == undefined) return safe_format_cell(cell, cell.v);
	return safe_format_cell(cell, v);
}
function make_json_row(sheet: any, r: any, R: any, cols: any, header: number, hdr: any, dense: boolean, o: any) {
	var rr = encode_row(R);
	var defval = o.defval, raw = o.raw || !Object.prototype.hasOwnProperty.call(o, "raw");
	var isempty = true;
  var row: any = (header === 1) ? [] : {};
  var hyperlink = [ ];
	if(header !== 1) {
		if(Object.defineProperty) try { Object.defineProperty(row, '__rowNum__', {value:R, enumerable:false}); } catch(e) { row.__rowNum__ = R; }
		else row.__rowNum__ = R;
	}
	if(!dense || sheet[R]) for (var C = r.s.c; C <= r.e.c; ++C) {
		var val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		if(val === undefined || val.t === undefined) {
			if(defval === undefined) continue;
			if(hdr[C] != null) { row[hdr[C]] = defval; }
			continue;
		}
		var v = val.v;
		switch(val.t){
			case 'z': if(v == null) break; continue;
			case 'e': v = void 0; break;
			case 's': case 'd': case 'b': case 'n': break;
			default: throw new Error('unrecognized type ' + val.t);
		}
		if(hdr[C] != null) {
			if(v == null) {
				if(defval !== undefined) row[hdr[C]] = defval;
				else if(raw && v === null) row[hdr[C]] = null;
				else continue;
			} else {
				row[hdr[C]] = raw || (o.rawNumbers && val.t == "n") ? v : format_cell(val,v,o);
			}
			if(v != null) isempty = false;
    }
    //Sail
    hyperlink[hdr[C]] = val.l;
	}
	return { row, hyperlink, isempty };
}
function sheet_to_json(sheet: any, opts: any) {
	if(sheet == null || sheet["!ref"] == null) return { data: [], hyperlink: [] };
	var val: any = {t:'n',v:0}, header = 0, offset = 1, hdr = [], v=0, vv="";
	var r = {s:{r:0,c:0},e:{r:0,c:0}};
	var o = opts || {};
	var range = o.range != null ? o.range : sheet["!ref"];
	if(o.header === 1) header = 1;
	else if(o.header === "A") header = 2;
	else if(Array.isArray(o.header)) header = 3;
	else if(o.header == null) header = 0;
	switch(typeof range) {
		case 'string': r = safe_decode_range(range); break;
		case 'number': r = safe_decode_range(sheet["!ref"]); r.s.r = range; break;
		default: r = range;
	}
	if(header > 0) offset = 0;
	var rr = encode_row(r.s.r);
	var cols = [];
	var out = [];
	var hyperlink = [];
	var outi = 0, counter = 0;
	var dense = Array.isArray(sheet);
	var R = r.s.r, C = 0, CC = 0;
	if(dense && !sheet[R]) sheet[R] = [];
	for(C = r.s.c; C <= r.e.c; ++C) {
		cols[C] = encode_col(C);
		val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		switch(header) {
			case 1: hdr[C] = C - r.s.c; break;
			case 2: hdr[C] = cols[C]; break;
			case 3: hdr[C] = o.header[C - r.s.c]; break;
			default:
				if(val == null) val = {w: "__EMPTY", t: "s"};
				vv = v = format_cell(val, null, o);
				counter = 0;
				for(CC = 0; CC < hdr.length; ++CC) if(hdr[CC] == vv) vv = v + "_" + (++counter);
				hdr[C] = vv;
		}
	}
	for (R = r.s.r + offset; R <= r.e.r; ++R) {
    var row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
    hyperlink[outi] = row.hyperlink;
		if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) out[outi++] = row.row;
	}
	out.length = outi;
	return { data: out, hyperlink };
}

/**
 * 解析Excel
 * @param buffer 
 * @param _options 
 */
export async function parseXlsx(buffer: Buffer, options?: xlsx.ParsingOptions): Promise<{
  name: string,
  data: Array<Array<any>>,
  image: Array<Array<any>>,
  hyperlink: Array<Array<{
    Target: string,
    Rel: {
      Target: string,
      TargetMode: string,
    },
    Tooltip: string,
  }>>,
}[]> {
  options = options || { };
  const workSheet = xlsx.read(buffer, options);
  const rvObj = Object.keys(workSheet.Sheets).map((name) => {
    const sheet = workSheet.Sheets[name];
    const { data, hyperlink } = sheet_to_json(sheet, { header: 1, raw: options.raw !== false });
    return { name, data, hyperlink };
  });
  for (let i2 = 0; i2 < rvObj.length; i2++) {
    const dataObj = rvObj[i2];
    (<any>dataObj).image = [];
  }
  const hzip = new Hzip(buffer);
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
    parseString(workbookBuf, function (err :any, result :any) {
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
    parseString(workbookRelsBuf, function (err :any, result :any) {
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
        parseString(sheetBuf, function (err :any, result :any) {
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
        parseString(sheetRelsBuf, function (err :any, result :any) {
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
          parseString(drawing1Buf, function (err :any, result :any) {
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
          parseString(drawing1RelsBuf, function (err :any, result :any) {
            if(err) {
              reject(err);
              return;
            }
            resolve(result);
          });
        });
        const relationships = drawing1Rels["Relationships"]["Relationship"];
        const xdrTwoCellAnchors: any[] = drawing1["xdr:wsDr"]["xdr:twoCellAnchor"];
        const xdrOneCellAnchors = drawing1["xdr:wsDr"]["xdr:oneCellAnchor"];
        let xdrCellAnchors = xdrTwoCellAnchors || [];
        if(xdrOneCellAnchors) {
          xdrCellAnchors = xdrCellAnchors.concat(xdrOneCellAnchors);
        }
        for (let i = 0; i < xdrCellAnchors.length; i++) {
          const xdrTwoCellAnchor = xdrCellAnchors[i];
          const xdrFrom = xdrTwoCellAnchor["xdr:from"];
          const xdrRowNum = Number(xdrFrom[0]["xdr:row"][0]);
          const xdrColNum = Number(xdrFrom[0]["xdr:col"][0]);
          if(!xdrTwoCellAnchor["xdr:pic"]) continue;
          const xdrPic = xdrTwoCellAnchor["xdr:pic"][0];
          const xdrBlipFill = xdrPic["xdr:blipFill"][0];
          if(!xdrBlipFill["a:blip"]) continue;
          const aBlip = xdrBlipFill["a:blip"][0];
          const rEmbed = aBlip["$"]["r:embed"];
          for (let k = 0; k < relationships.length; k++) {
            const relationship = relationships[k]["$"];
            if(relationship["Id"] !== rEmbed) continue;
            const target = normalize(`xl/drawings/${ relationship["Target"] }`).replace(/\\/gm, "/");
            const targetEntity = hzip.getEntry(target);
            let targetBuf = targetEntity.cfile;
            //解压版本
            if(targetEntity.unzipVersion[0] === 0x14 && targetEntity.unzipVersion[1] === 0x00) {
              targetBuf = await new Promise((resolve, reject) => {
                inflateRaw(targetBuf, function (err: any, result: any) {
                  if(err) {
                    reject(err);
                    return;
                  }
                  resolve(result);
                });
              });
            }
            // console.log({ sheetName, xdrRowNum, xdrColNum, target });
            for (let i2 = 0; i2 < rvObj.length; i2++) {
              const dataObj = <any>rvObj[i2];
              if(dataObj.name !== sheetName) continue;
              dataObj.image[xdrRowNum] = dataObj.image[xdrRowNum] || [];
              dataObj.image[xdrRowNum][xdrColNum] = targetBuf;
            }
          }
        }
      }
    }
  }
  return <any>rvObj;
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
  return date;
}
