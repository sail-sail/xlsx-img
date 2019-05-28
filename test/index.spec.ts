import { readFile } from "fs-extra";
import { parseXlsx, xlsx2Date } from "../src/index";
import * as moment from "moment";

test("parseXlsx", async () => {
  const buffer = await readFile(`${__dirname}/test.xlsx`);
  const data = await parseXlsx(buffer);
  expect(data[0].name).toBe("Sheet1");
  for (let i = 4; i < 6; i++) {
    const item = data[0].data[i];
    expect(item[4]).toBeInstanceOf(Buffer);
  }
});

test("xlsx2Date", () => {
  const date1 = xlsx2Date(42793);
  expect(moment(date1).format("YYYY-MM-DD HH:mm:ss")).toBe("2017-02-27 00:00:00");
  const date2 = xlsx2Date(42817);
  expect(moment(date2).format("YYYY-MM-DD HH:mm:ss")).toBe("2017-03-23 00:00:00");
});