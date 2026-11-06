// 造一份与真实空模板同构的测试模板：13 列表头 + 楼宇名称下拉 + 楼栋名称工作表
import path from 'node:path';
import ExcelJS from 'exceljs';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
await import(pathToFileURL(path.join(here, '..', 'src', 'core.js')).href);
const FF = globalThis.FormFiller;

const OPTIONS = ['北区3-', '北区4-', '北区5-', '北区9-', '北区10-',
  '南区1-', '南区2-', '南区6-', '南区7-', '南区8-'];

const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('Sheet0');
ws.addRow(FF.COLUMNS);

const header = ws.getRow(1);
header.height = 30;
header.eachCell((cell) => {
  cell.font = { bold: true, size: 10, color: { argb: 'FFC00000' } };
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
  cell.border = {
    top: { style: 'thin' }, left: { style: 'thin' },
    bottom: { style: 'thin' }, right: { style: 'thin' }
  };
});
[58, 8, 10, 10, 10, 12, 16, 8, 8, 8, 8, 8, 8].forEach((w, i) => { ws.getColumn(i + 1).width = w; });

const list = wb.addWorksheet('楼栋名称');
OPTIONS.forEach((o, i) => { list.getCell('A' + (i + 1)).value = o; });

ws.dataValidations.add('C2:C3000', {
  type: 'list',
  allowBlank: true,
  formulae: ['=楼栋名称!$A$1:$A$' + OPTIONS.length]
});
ws.views = [{ state: 'frozen', ySplit: 1 }];

const out = path.join(here, '..', 'samples', '空模板-测试.xlsx');
await wb.xlsx.writeFile(out);
console.log('写出', out, '列数', FF.COLUMNS.length);
