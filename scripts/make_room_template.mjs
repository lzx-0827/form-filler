// 生成 templates/房间编号模板.xlsx，并用核心解析器自检
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import ExcelJS from 'exceljs';

const here = path.dirname(new URL(import.meta.url).pathname).replace(/^\/([A-Za-z]:)/, '$1');
await import(pathToFileURL(path.join(here, '..', 'src', 'room-template.js')).href);
await import(pathToFileURL(path.join(here, '..', 'src', 'core.js')).href);
const FF = globalThis.FormFiller;

const wb = globalThis.RoomTemplate.build(ExcelJS, 300);
const outDir = path.join(here, '..', 'templates');
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, '房间编号模板.xlsx');
await wb.xlsx.writeFile(out);
console.log('written', out);

// 自检：填 3 行后解析
const test = globalThis.RoomTemplate.build(ExcelJS, 10);
const ws = test.getWorksheet('房间清单');
[['1#', '', '1-1001'], ['2#', 'A', '205'], ['6#', '', '906']].forEach((row, i) => {
  ws.getRow(i + 2).values = [row[0], row[1], row[2]];
});
const parsed = FF.parseRoomSheet(ws);
console.log('layout:', parsed.layout, 'rooms:', parsed.rooms.length);
console.log('buildings:', parsed.buildings.map((b) => b.name + ':' + b.count).join(' '));
const mapping = FF.suggestMapping(parsed.buildings.map((b) => b.name), ['1-', '2-', '6-']);
console.log('mapping:', JSON.stringify(mapping));
console.log('rows:', parsed.rooms.map((room) => {
  const v = FF.buildRowValues(room, {
    prefix: '测试小区', buildingName: mapping[room.building], installSite: '厨房表'
  });
  return [v.address, v.floor, v.room];
}));
