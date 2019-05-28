### xlsx-img
like [node-xlsx](https://github.com/mgcrea/node-xlsx) and suport parse image
```typescript
npm install --save xlsx-img
```
```typescript
import { readFile } from "fs-extra";
import { parseXlsx } from "xlsx-img";

(async () => {
  const buffer = await readFile(`${__dirname}/test/test.xlsx`);
  const data = await parseXlsx(buffer);
  console.log(data);
})();
```
