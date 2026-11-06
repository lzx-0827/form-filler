/*
 * 填表器 · 核心逻辑（与界面无关，可在浏览器和 Node 中运行）
 * 只做四件事：解析数据表、检测模板占位符、从文字里抽取字段、把值写回模板。
 */
(function (root) {
  'use strict';

  var PLACEHOLDER = /\{\{\s*([^{}]+?)\s*\}\}/g;

  /* ---------------- 通用 ---------------- */

  function normalizeLabel(value) {
    return String(value == null ? '' : value)
      .trim()
      .toLowerCase()
      .replace(/[\s_\-·、,，.。:：;；()（）\[\]【】]/g, '');
  }

  function pad2(n) {
    return (n < 10 ? '0' : '') + n;
  }

  function formatDate(d) {
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  }

  function cellText(cell) {
    if (!cell) return '';
    var v = cell.value;
    if (v == null) return '';
    if (v instanceof Date) return formatDate(v);
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(function (t) { return t.text || ''; }).join('');
      if (v.text != null) return String(v.text);
      if (v.error) return String(v.error);
      if (v.formula || v.sharedFormula) return String(cell.text == null ? '' : cell.text);
    }
    return String(v);
  }

  /* ---------------- CSV / 数据表 ---------------- */

  function sniffDelimiter(line) {
    var candidates = [',', '\t', ';'];
    var best = ',';
    var bestCount = -1;
    candidates.forEach(function (d) {
      var count = 0;
      var quoted = false;
      for (var i = 0; i < line.length; i++) {
        var ch = line[i];
        if (ch === '"') quoted = !quoted;
        else if (ch === d && !quoted) count++;
      }
      if (count > bestCount) { bestCount = count; best = d; }
    });
    return best;
  }

  function tableFromMatrix(matrix, headerRow) {
    var rows = (matrix || []).filter(function (r) {
      return r.some(function (c) { return String(c == null ? '' : c).trim() !== ''; });
    });
    if (!rows.length) return { headers: [], rows: [] };
    var h = Math.max(1, Math.min(headerRow || 1, rows.length)) - 1;
    var rawHeaders = rows[h] || [];
    var headers = [];
    var used = {};
    for (var i = 0; i < rawHeaders.length; i++) {
      var name = String(rawHeaders[i] == null ? '' : rawHeaders[i]).trim() || '列' + (i + 1);
      if (used[name]) name = name + '(' + (used[name] + 1) + ')';
      used[name] = (used[name] || 0) + 1;
      headers.push(name);
    }
    var data = rows.slice(h + 1).map(function (r) {
      var obj = {};
      headers.forEach(function (name, idx) {
        var v = r[idx];
        obj[name] = v == null ? '' : String(v).trim();
      });
      return obj;
    });
    return { headers: headers, rows: data };
  }

  function parseCSV(text) {
    text = String(text == null ? '' : text).replace(/^\uFEFF/, '');
    var firstLine = text.split(/\r?\n/).find(function (l) { return l.trim() !== ''; }) || '';
    var delimiter = sniffDelimiter(firstLine);
    var rows = [];
    var row = [];
    var field = '';
    var quoted = false;

    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (quoted) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else quoted = false;
        } else field += ch;
        continue;
      }
      if (ch === '"') { quoted = true; continue; }
      if (ch === delimiter) { row.push(field); field = ''; continue; }
      if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(field); field = '';
        rows.push(row); row = [];
        continue;
      }
      field += ch;
    }
    row.push(field);
    rows.push(row);

    rows = rows.filter(function (r) {
      return r.some(function (c) { return String(c).trim() !== ''; });
    });

    if (!rows.length) return { headers: [], rows: [], matrix: [], delimiter: delimiter };

    var table = tableFromMatrix(rows, 1);
    table.matrix = rows;
    table.delimiter = delimiter;
    return table;
  }

  /* ---------------- 模板占位符 ---------------- */
  function placeholderNames(text) {
    var names = [];
    var m;
    PLACEHOLDER.lastIndex = 0;
    while ((m = PLACEHOLDER.exec(text)) !== null) {
      var raw = m[1].trim();
      if (raw) names.push(raw);
    }
    return names;
  }

  function detectPlaceholders(workbook) {
    var fields = [];
    var seen = {};
    var cells = [];
    workbook.eachSheet(function (ws) {
      ws.eachRow({ includeEmpty: false }, function (row) {
        row.eachCell({ includeEmpty: false }, function (cell) {
          if (typeof cell.value !== 'string') return;
          var names = placeholderNames(cell.value);
          if (!names.length) return;
          cells.push({ sheet: ws.name, address: cell.address, names: names });
          names.forEach(function (name) {
            var key = normalizeLabel(name);
            if (!seen[key]) { seen[key] = true; fields.push(name); }
          });
        });
      });
    });
    return { fields: fields, cells: cells };
  }

  /* ---------------- 文字抽取 ---------------- */

  var BUILTIN = [
    {
      label: '日期',
      re: /(\d{4})\s*[年\-\/.]\s*(\d{1,2})\s*[月\-\/.]\s*(\d{1,2})\s*日?/,
      pick: function (m) { return m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]); }
    },
    {
      label: '金额',
      re: /(?:¥|￥|RMB|人民币)?\s*(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)\s*(?:元|块)/,
      pick: function (m) { return m[1].replace(/,/g, ''); }
    },
    { label: '电子邮箱', re: /([\w.+-]+@[\w-]+\.[\w.-]+)/, pick: function (m) { return m[1]; } },
    { label: '手机号', re: /(1[3-9]\d{9})/, pick: function (m) { return m[1]; } },
    { label: '身份证号', re: /(\d{17}[\dXx])/, pick: function (m) { return m[1].toUpperCase(); } }
  ];

  function extractFields(text, knownLabels) {
    var out = [];
    var seen = {};
    var lines = String(text == null ? '' : text).split(/\r?\n/);

    function add(label, value, source) {
      label = String(label).trim().replace(/[：:]\s*$/, '');
      value = String(value == null ? '' : value).trim();
      if (!label || !value) return;
      var key = normalizeLabel(label);
      if (seen[key]) return;
      seen[key] = true;
      out.push({ label: label, value: value, source: source });
    }

    // 1) “标签：值”逐行
    lines.forEach(function (line) {
      var m = line.match(/^\s*([^\s:：=]{1,24})\s*[:：=]\s*(.+?)\s*$/);
      if (m) add(m[1], m[2], '文字');
    });

    // 2) 已知字段名在句中出现（“客户名称 张三”）
    (knownLabels || []).forEach(function (label) {
      if (!label || seen[normalizeLabel(label)]) return;
      var escaped = String(label).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      var re = new RegExp(escaped + '\\s*[:：]?\\s*([^\\n，,；;。]{1,40})');
      var m = String(text).match(re);
      if (m) add(label, m[1], '文字');
    });

    // 3) 内置识别（日期、金额、邮箱、手机、身份证）
    BUILTIN.forEach(function (rule) {
      var m = String(text).match(rule.re);
      if (m) add(rule.label, rule.pick(m), '识别');
    });

    return out;
  }

  /* ---------------- 单元格 / 合并区域 ---------------- */

  function colToNumber(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }

  function numberToCol(n) {
    var s = '';
    while (n > 0) {
      var rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function parseAddress(address) {
    var m = String(address).toUpperCase().match(/^\$?([A-Z]{1,3})\$?(\d+)$/);
    if (!m) return null;
    return { col: colToNumber(m[1]), row: parseInt(m[2], 10) };
  }

  function buildMergeIndex(ws) {
    var byMaster = {};
    var covered = {};
    var list = (ws.model && ws.model.merges) || [];
    list.forEach(function (range) {
      var parts = String(range).split(':');
      var a = parseAddress(parts[0]);
      var b = parseAddress(parts[1] || parts[0]);
      if (!a || !b) return;
      var top = Math.min(a.row, b.row);
      var left = Math.min(a.col, b.col);
      var bottom = Math.max(a.row, b.row);
      var right = Math.max(a.col, b.col);
      var master = numberToCol(left) + top;
      byMaster[master] = { rowSpan: bottom - top + 1, colSpan: right - left + 1 };
      for (var r = top; r <= bottom; r++) {
        for (var c = left; c <= right; c++) {
          var addr = numberToCol(c) + r;
          if (addr !== master) covered[addr] = master;
        }
      }
    });
    return { byMaster: byMaster, covered: covered };
  }

  /* ---------------- 网格（用于可视化绑定） ---------------- */

  function sheetGrid(ws, options) {
    var opts = options || {};
    var maxRows = opts.maxRows || 40;
    var maxCols = opts.maxCols || 14;
    var merge = buildMergeIndex(ws);
    var lastRow = Math.min(ws.rowCount || 1, maxRows);
    var lastCol = Math.min(ws.columnCount || 1, maxCols);

    var rows = [];
    for (var r = 1; r <= lastRow; r++) {
      var cells = [];
      for (var c = 1; c <= lastCol; c++) {
        var addr = numberToCol(c) + r;
        if (merge.covered[addr]) continue;
        var cell = ws.getCell(addr);
        var span = merge.byMaster[addr] || { rowSpan: 1, colSpan: 1 };
        cells.push({
          address: addr,
          text: cellText(cell),
          row: r,
          col: c,
          rowSpan: span.rowSpan,
          colSpan: span.colSpan,
          merged: span.rowSpan > 1 || span.colSpan > 1
        });
      }
      rows.push({ index: r, cells: cells });
    }

    return {
      rows: rows,
      columns: lastCol,
      truncated: (ws.rowCount || 0) > maxRows || (ws.columnCount || 0) > maxCols,
      totalRows: ws.rowCount || 0,
      totalCols: ws.columnCount || 0
    };
  }

  /* ---------------- 填充 ---------------- */

  function resolveMaster(ws, address) {
    var merge = buildMergeIndex(ws);
    return merge.covered[String(address).toUpperCase()] || String(address).toUpperCase();
  }

  function setCellValue(ws, address, value) {
    var target = resolveMaster(ws, address);
    var cell = ws.getCell(target);
    ws.getCell(target).value = value;
    return { address: target, cell: cell };
  }

  function coerce(value, workbook) {
    var s = String(value == null ? '' : value);
    if (s === '') return '';
    if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
    return s;
  }

  /*
   * 把值写进模板。
   * bindings: [{ sheet, address, field }]
   * values:   { 字段名: 值 }
   * 返回 { unresolved: [...], filled: [{sheet, address, field, value}] }
   */
  function fillWorkbook(workbook, bindings, values, options) {
    var opts = options || {};
    var today = formatDate(new Date());
    var unresolved = {};
    var filled = [];
    var bindingsBySheet = {};

    (bindings || []).forEach(function (b) {
      (bindingsBySheet[b.sheet] = bindingsBySheet[b.sheet] || []).push(b);
    });

    workbook.eachSheet(function (ws) {
      // 1) 占位符替换
      ws.eachRow({ includeEmpty: false }, function (row) {
        row.eachCell({ includeEmpty: false }, function (cell) {
          if (typeof cell.value !== 'string') return;
          if (cell.value.indexOf('{{') === -1) return;
          var changed = false;
          var next = cell.value.replace(PLACEHOLDER, function (all, raw) {
            var name = String(raw).trim();
            if (name === '@today') { changed = true; return today; }
            var value = values[name];
            if (value == null || value === '') { unresolved[name] = true; changed = true; return ''; }
            changed = true;
            filled.push({ sheet: ws.name, address: cell.address, field: name, value: String(value) });
            return String(value);
          });
          if (changed) {
            cell.value = opts.coerceNumbers === false ? next : coerce(next, workbook);
          }
        });
      });

      // 2) 手动绑定的单元格
      (bindingsBySheet[ws.name] || []).forEach(function (b) {
        var value = values[b.field];
        if (value == null || value === '') { unresolved[b.field] = true; return; }
        var target = resolveMaster(ws, b.address);
        var cell = ws.getCell(target);
        cell.value = opts.coerceNumbers === false ? String(value) : coerce(value, workbook);
        filled.push({ sheet: ws.name, address: target, field: b.field, value: String(value) });
      });
    });

    return { unresolved: Object.keys(unresolved), filled: filled };
  }

  /* ---------------- 导出 ---------------- */

  /* ---------------- 房产表计单据：模板识别、房间清单、行生成 ---------------- */

  var HEADER_RULES = [
    { key: 'address', exact: '房屋地址', re: /房屋地址/ },
    { key: 'opened', exact: '是否已开支', re: /是否已开/ },
    { key: 'building', exact: '楼宇名称', re: /(楼宇|楼栋)名称/ },
    { key: 'site', exact: '安装点位置', re: /安装点位置/ },
    { key: 'access', exact: '是否可接入', re: /是否可接入/ },
    { key: 'note', exact: '作业注意事项', re: /作业注意事项/ },
    { key: 'generated', exact: '是否通过单元号楼层号房号生成房屋地址', re: /生成房屋地址/ },
    { key: 'unit', exact: '单元号', re: /^单元号$/ },
    { key: 'unitSuffix', exact: '单元后缀', re: /单元后缀/ },
    { key: 'floor', exact: '楼层号', re: /^楼层号$/ },
    { key: 'floorSuffix', exact: '楼层后缀', re: /楼层后缀/ },
    { key: 'room', exact: '房号', re: /^房号$/ },
    { key: 'roomSuffix', exact: '房号后缀', re: /房号后缀/ }
  ];

  function headerText(value) {
    return String(value == null ? '' : value).replace(/[\s*＊]/g, '');
  }

  function mapColumns(values) {
    var texts = values.map(headerText);
    var used = {};
    var map = {};

    // 先精确匹配，避免「楼层号」被“是否通过单元号楼层号房号…”这种长表头抢走
    HEADER_RULES.forEach(function (rule) {
      for (var i = 0; i < texts.length; i++) {
        if (!used[i] && texts[i] === rule.exact) {
          map[rule.key] = i;
          used[i] = true;
          return;
        }
      }
    });

    HEADER_RULES.forEach(function (rule) {
      if (map[rule.key] != null) return;
      for (var i = 0; i < texts.length; i++) {
        if (!used[i] && rule.re.test(texts[i])) {
          map[rule.key] = i;
          used[i] = true;
          return;
        }
      }
    });

    return map;
  }

  function uniq(list) {
    var seen = {};
    return list.filter(function (x) {
      if (!x || seen[x]) return false;
      seen[x] = true;
      return true;
    });
  }

  function resolveListFormula(workbook, formula) {
    if (!formula) return [];
    var f = String(formula).trim().replace(/^=/, '');
    if (f.charAt(0) === '"') {
      return uniq(f.replace(/^"|"$/g, '').split(',').map(function (s) { return s.trim(); }));
    }
    var m = f.match(/^(?:'([^']+)'|([^!]+))!(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)$/);
    if (!m) return [];
    var sheet = workbook.getWorksheet(m[1] || m[2]);
    if (!sheet) return [];
    var a = parseAddress(m[3].replace(/\$/g, ''));
    var b = parseAddress(m[4].replace(/\$/g, ''));
    if (!a || !b) return [];
    var out = [];
    var top = Math.min(a.row, b.row), bottom = Math.max(a.row, b.row);
    var left = Math.min(a.col, b.col), right = Math.max(a.col, b.col);
    for (var r = top; r <= bottom; r++) {
      for (var c = left; c <= right; c++) {
        var v = cellText(sheet.getRow(r).getCell(c)).trim();
        if (v) out.push(v);
      }
    }
    return uniq(out);
  }

  // 读取某一列的数据验证下拉选项
  function readDvOptions(workbook, ws, colIndex) {
    var model = ws.dataValidations && ws.dataValidations.model;
    if (!model) return [];
    var keys = Object.keys(model);
    for (var i = 0; i < keys.length; i++) {
      var range = String(keys[i]).split(' ')[0];
      var cell = parseAddress(String(range).split(':')[0]);
      if (colIndex != null && cell && cell.col !== colIndex + 1) continue;
      var dv = model[keys[i]];
      var formulae = dv.formulae || (dv.formula ? [dv.formula] : []);
      for (var k = 0; k < formulae.length; k++) {
        var values = resolveListFormula(workbook, formulae[k]);
        if (values.length) return values;
      }
    }
    return [];
  }

  // 楼宇名称选项：优先「楼栋名称 / 楼宇名称」工作表，其次单元格数据验证，最后读现有数据
  function readBuildingOptions(workbook, ws, buildingIndex) {
    var fromSheet = [];
    workbook.eachSheet(function (sheet) {
      if (fromSheet.length) return;
      if (/楼(栋|宇)名称/.test(sheet.name)) {
        for (var r = 1; r <= sheet.rowCount; r++) {
          var v = cellText(sheet.getRow(r).getCell(1)).trim();
          if (v && !/名称/.test(v)) fromSheet.push(v);
        }
      }
    });
    if (fromSheet.length) return uniq(fromSheet);

    var dvValues = readDvOptions(workbook, ws, buildingIndex);
    if (dvValues.length) return dvValues;

    if (buildingIndex != null) {
      var colValues = [];
      for (var r = 2; r <= ws.rowCount; r++) {
        var v2 = cellText(ws.getRow(r).getCell(buildingIndex + 1)).trim();
        if (v2) colValues.push(v2);
      }
      return uniq(colValues);
    }
    return [];
  }

  // 识别空模板：表头行、各列位置、楼宇名称下拉选项
  function detectTemplate(workbook) {
    var found = null;
    workbook.eachSheet(function (ws) {
      if (found) return;
      var maxRow = Math.min(10, ws.rowCount || 0);
      for (var r = 1; r <= maxRow; r++) {
        var values = [];
        var hit = false;
        for (var c = 1; c <= (ws.columnCount || 0); c++) {
          var text = cellText(ws.getRow(r).getCell(c)).trim();
          values.push(text);
          if (text.indexOf('房屋地址') !== -1) hit = true;
        }
        if (hit) { found = { ws: ws, headerRow: r, values: values }; break; }
      }
    });
    if (!found) return null;

    var columns = mapColumns(found.values);
    if (columns.address == null || columns.building == null) return null;

    var options = {};
    HEADER_RULES.forEach(function (rule) {
      if (columns[rule.key] == null) return;
      options[rule.key] = rule.key === 'building'
        ? readBuildingOptions(workbook, found.ws, columns[rule.key])
        : readDvOptions(workbook, found.ws, columns[rule.key]);
    });

    return {
      sheetName: found.ws.name,
      headerRow: found.headerRow,
      columns: columns,
      options: options,
      columnCount: found.ws.columnCount || 0,
      headerValues: found.values
    };
  }

  // 从纯值构造模板信息（用于 .xls：按模板自身表头定位，不假设行列）
  function templateFromRows(rows, options, sheetName) {
    rows = rows || [];
    var headerIndex = -1;
    for (var i = 0; i < Math.min(rows.length, 10); i++) {
      var row = rows[i] || [];
      for (var j = 0; j < row.length; j++) {
        if (String(row[j] == null ? '' : row[j]).indexOf('房屋地址') !== -1) {
          headerIndex = i;
          break;
        }
      }
      if (headerIndex !== -1) break;
    }
    if (headerIndex === -1) return null;

    var values = rows[headerIndex];
    var columns = mapColumns(values);
    if (columns.address == null || columns.building == null) return null;
    var opts = options || {};
    if (Array.isArray(opts)) opts = { building: opts };
    return {
      sheetName: sheetName || 'Sheet1',
      headerRow: headerIndex + 1,
      columns: columns,
      options: opts,
      columnCount: values.length,
      headerValues: values,
      source: 'xls'
    };
  }

  // 解析《房间编号》：自动识别「楼栋/单元/楼层/门牌」；没有楼栋就是单楼栋小区
  function parseRoomSheet(ws) {
    var lastRow = ws.rowCount || 0;
    var lastCol = ws.columnCount || 0;
    var headerRow = 0, bCol = 0, uCol = 0, fCol = 0, doorCol = 0;

    for (var r = 1; r <= Math.min(8, lastRow); r++) {
      var texts = [];
      for (var c = 1; c <= lastCol; c++) texts.push(cellText(ws.getRow(r).getCell(c)).trim());
      var fi = -1;
      for (var i = 0; i < texts.length; i++) {
        if (/^楼层/.test(texts[i])) { fi = i; break; }
      }
      if (fi >= 0) {
        headerRow = r;
        fCol = fi + 1;
        for (var j = 0; j < texts.length; j++) {
          if (!bCol && /^(楼栋|楼宇)/.test(texts[j])) bCol = j + 1;
          if (!uCol && /^单元/.test(texts[j])) uCol = j + 1;
          if (!doorCol && /^门牌/.test(texts[j])) doorCol = j + 1;
        }
        break;
      }
    }
    if (!headerRow) { headerRow = 1; bCol = 1; uCol = 2; fCol = 3; }

    var mode = bCol ? 'building' : 'single';
    var rooms = [];
    var buildings = [];
    var index = {};
    var curBuilding = mode === 'single' ? '-' : '';
    var curUnit = '';

    for (var row = headerRow + 1; row <= lastRow; row++) {
      var sheetRow = ws.getRow(row);
      if (mode === 'building') {
        var bText = cellText(sheetRow.getCell(bCol)).trim();
        if (bText) { curBuilding = bText; curUnit = ''; }
        if (!curBuilding) continue;
        if (uCol) {
          var uText = cellText(sheetRow.getCell(uCol)).trim();
          if (uText) curUnit = (uText === '/' || uText === '-') ? '' : uText.replace(/单元\s*$/, '').trim();
        }
        if (!index[curBuilding]) {
          index[curBuilding] = { name: curBuilding, count: 0, units: [] };
          buildings.push(index[curBuilding]);
        }
      }

      var floorText = cellText(sheetRow.getCell(fCol)).replace(/[^\d]/g, '');
      var floor = floorText ? parseInt(floorText, 10) : null;

      var doors = [];
      var doorStart = doorCol || (fCol + 1);
      for (var dc = doorStart; dc <= lastCol; dc++) {
        var v = cellText(sheetRow.getCell(dc)).trim();
        if (v) doors.push(v);
      }

      doors.forEach(function (door) {
        if (door === '/' || door === '-') return;
        rooms.push({ building: curBuilding, unit: curUnit, floor: floor, door: door, row: row });
        if (mode === 'building') {
          index[curBuilding].count++;
          if (index[curBuilding].units.indexOf(curUnit) < 0) index[curBuilding].units.push(curUnit);
        }
      });
    }

    if (mode === 'single') {
      buildings = [{ name: '-', count: rooms.length, units: [] }];
    }

    return { mode: mode, headerRow: headerRow, rooms: rooms, buildings: buildings };
  }

  // 门牌 101 + 楼层 1 → 房号 01；门牌 1301 + 楼层 13 → 房号 01
  function splitDoor(door, floor) {
    var d = String(door == null ? '' : door).trim();
    var f = String(floor == null ? '' : floor).trim();
    if (f && d.indexOf(f) === 0) return d.slice(f.length);
    return d;
  }

  // 地址 = 前缀 + 楼宇名称(带-) + 单元号(带-) + 门牌
  function composeAddress(prefix, buildingName, unit, door) {
    return String(prefix == null ? '' : prefix) +
      String(buildingName == null ? '' : buildingName) +
      (unit ? unit + '-' : '') +
      String(door == null ? '' : door);
  }

  /*
   * 生成一行数据（按模板列名返回对象）。
   * room: {building, unit, floor, door}
   * opts: {prefix, buildingName, installSite, villaSite}
   */
  function buildRowValues(room, opts) {
    var unit = room.unit || '';
    var addr = composeAddress(opts.prefix, opts.buildingName, unit, room.door);
    var floor = room.floor == null ? '' : String(room.floor);
    var roomNo = splitDoor(room.door, floor);
    var site = opts.installSite || '厨房表';
    if (opts.villaSite !== false && addr.indexOf('别墅') !== -1) site = '别墅表';
    return {
      address: addr,
      opened: '否',
      building: opts.buildingName,
      site: site,
      access: '是',
      note: '',
      generated: '否',
      unit: unit ? unit + '-' : '',
      unitSuffix: opts.unitSuffix || '',
      floor: floor,
      floorSuffix: opts.floorSuffix || '',
      room: roomNo,
      roomSuffix: opts.roomSuffix || ''
    };
  }

  // 「北区3-」→ { area: '北区', number: '3' }
  function parseBuildingOption(option) {
    var s = String(option == null ? '' : option).trim();
    var m = s.match(/^(.*?)(\d+)\s*-?\s*$/);
    if (!m) return { raw: s, area: '', number: '' };
    return { raw: s, area: m[1].trim(), number: m[2] };
  }

  // 只靠模板选项做匹配：栋号 3 → 选项里编号为 3 的「北区3-」；有歧义就不猜
  function suggestMapping(buildingNames, options) {
    var parsed = (options || []).map(parseBuildingOption);
    var result = {};
    var used = {};
    buildingNames.forEach(function (name) {
      var m = String(name).match(/(\d+)/);
      if (!m) return;
      var candidates = parsed.filter(function (p) {
        return p.number === m[1] && !used[p.raw];
      });
      if (candidates.length === 1) {
        result[name] = candidates[0].raw;
        used[candidates[0].raw] = true;
      }
    });
    return result;
  }

  // 楼宇名 → 区域（北区3- → 北区），没有区域时返回空
  function areaOf(buildingName) {
    return parseBuildingOption(buildingName).area;
  }

  var api = {
    normalizeLabel: normalizeLabel,
    formatDate: formatDate,
    cellText: cellText,
    parseCSV: parseCSV,
    tableFromMatrix: tableFromMatrix,
    placeholderNames: placeholderNames,
    detectPlaceholders: detectPlaceholders,
    extractFields: extractFields,
    sheetGrid: sheetGrid,
    buildMergeIndex: buildMergeIndex,
    parseAddress: parseAddress,
    numberToCol: numberToCol,
    fillWorkbook: fillWorkbook,
    detectTemplate: detectTemplate,
    buildRowValues: buildRowValues,
    templateFromRows: templateFromRows,
    parseRoomSheet: parseRoomSheet,
    splitDoor: splitDoor,
    composeAddress: composeAddress,
    suggestMapping: suggestMapping,
    parseBuildingOption: parseBuildingOption,
    areaOf: areaOf
  };

  if (typeof module === 'object' && module.exports) module.exports = api;
  root.FormFiller = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);

