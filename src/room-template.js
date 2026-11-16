/*
 * 生成《房间编号》录入模板（ExcelJS 工作簿）。
 * 第一张表是数据，第二张是填写说明。
 */
(function (root) {
  'use strict';

  var THIN = { style: 'thin', color: { argb: 'FFD9D4C8' } };

  function styleHeader(row) {
    row.height = 24;
    row.eachCell(function (cell) {
      cell.font = { bold: true, size: 11, color: { argb: 'FF1D1F21' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F2EE' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = { top: THIN, left: THIN, bottom: THIN, right: THIN };
    });
  }

  function build(ExcelJS, blankRows) {
    var wb = new ExcelJS.Workbook();
    wb.creator = '表计安装点填表器';

    var ws = wb.addWorksheet('房间清单', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });
    ws.columns = [
      { header: '楼栋', key: 'building', width: 18 },
      { header: '单元（选填）', key: 'unit', width: 16 },
      { header: '房间编号', key: 'room', width: 30 }
    ];
    styleHeader(ws.getRow(1));

    var total = blankRows || 300;
    for (var r = 2; r <= total + 1; r++) {
      var row = ws.getRow(r);
      row.height = 20;
      for (var c = 1; c <= 3; c++) {
        var cell = row.getCell(c);
        cell.border = { top: THIN, left: THIN, bottom: THIN, right: THIN };
        cell.numFmt = '@';
      }
    }

    var info = wb.addWorksheet('填写说明');
    info.columns = [{ width: 4 }, { width: 78 }];
    var lines = [
      ['房间编号填写说明', ''],
      ['', ''],
      ['1', '只填「房间清单」这张表；不要改表头，不要合并单元格。'],
      ['2', '楼栋：写楼栋名，如 1#、1栋、北区3；要与导入模板里的「楼宇名称」对得上。'],
      ['3', '单元：没有单元就留空；有单元写 A 或 B。'],
      ['4', '房间编号：写门牌，如 1-1001 或 1001；带不带楼栋前缀都可以。'],
      ['5', '每行一个房间。楼层号、房号由程序拆分：末两位是房号，前面是楼层。'],
      ['6', '同一楼栋可以写多行；示例见下方。'],
      ['', ''],
      ['示例', '楼栋          单元          房间编号'],
      ['', '1#                          1-1001'],
      ['', '2#            A              205'],
      ['', '6#                          906']
    ];
    lines.forEach(function (pair, i) {
      var row = info.getRow(i + 1);
      row.getCell(1).value = pair[0];
      row.getCell(2).value = pair[1];
      row.getCell(2).alignment = { wrapText: true, vertical: 'top' };
    });
    info.getRow(1).getCell(1).font = { bold: true, size: 14 };
    info.mergeCells('A1:B1');

    return wb;
  }

  root.RoomTemplate = { build: build };
})(typeof globalThis !== 'undefined' ? globalThis : this);
