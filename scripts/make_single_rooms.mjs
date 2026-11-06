// 造一份“不分楼栋”的房间清单：楼层 | 门牌
import path from 'node:path';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('Sheet1');
ws.addRow(['楼层', '门牌']);
[['1F', 101], ['1F', 102], ['2F', 201], ['2F', 202], ['3F', 301], ['3F', 302]]
  .forEach((r) => ws.addRow(r));
await wb.xlsx.writeFile(path.join(here, '..', 'samples', '单楼栋房间.xlsx'));
console.log('ok');
