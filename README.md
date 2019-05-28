### xlsx-img
node-xlsx suport parse image
```typescript
npm install --save node-xlsx
```
```typescript
import { readFile } from "fs-extra";
import { parseXlsx } from "node-xlsx";

(async () => {
  const buffer = await readFile(`${__dirname}/test/test.xlsx`);
  const data = await parseXlsx(buffer);
  console.log(data);
})();
```
