/*
 * 表计安装点填表器 · 界面逻辑
 * 数据只在内存：无 localStorage、无网络请求。
 */
(function () {
  'use strict';

  var FF = window.FormFiller;
  var $ = function (sel) { return document.querySelector(sel); };

  var PAGE_SIZE = 100;

  var state = {
    template: null,   // { fileName, buffer, info }
    fileName: '',
    mode: 'building',
    rooms: [],
    buildings: [],
    mapping: {},
    overrides: {},
    suffixes: { unitSuffix: '', floorSuffix: '', roomSuffix: '' },
    suffixOverrides: {},
    filterBuilding: '',
    filterFloor: '',
    query: '',
    page: 0
  };

  /* ---------------- 工具 ---------------- */

  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  var toastTimer = null;
  function toast(msg, isError) {
    var el = $('#toast');
    el.textContent = msg;
    el.hidden = false;
    el.style.background = isError ? '#8a241c' : '#1f2937';
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.hidden = true; }, 3200);
  }

  function roomKey(room) {
    return room.building + '|' + (room.unit || '') + '|' + room.door;
  }

  function unique(list) {
    var seen = {};
    return list.filter(function (x) {
      if (seen[x]) return false;
      seen[x] = true;
      return true;
    });
  }

  function opt(key) {
    var options = state.template && state.template.info.options;
    return (options && options[key]) || [];
  }

  function optionsHtml(selected) {
    var options = opt('building');
    var areas = {};
    options.forEach(function (name) {
      var area = FF.areaOf(name) || '全部';
      (areas[area] = areas[area] || []).push(name);
    });
    var html = '<option value="">— 请选择 —</option>';
    Object.keys(areas).forEach(function (area) {
      if (Object.keys(areas).length > 1) html += '<optgroup label="' + escapeHtml(area) + '">';
      areas[area].forEach(function (name) {
        html += '<option value="' + escapeHtml(name) + '"' +
          (name === selected ? ' selected' : '') + '>' + escapeHtml(name) + '</option>';
      });
      if (Object.keys(areas).length > 1) html += '</optgroup>';
    });
    return html;
  }

  /* ---------------- 读取空模板 ---------------- */

  function optionsFromXls(wb) {
    var listName = wb.SheetNames.filter(function (n) { return /楼(栋|宇)名称/.test(n); })[0];
    if (!listName) return [];
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[listName], { header: 1, blankrows: false });
    return rows.map(function (r) { return String(r[0] == null ? '' : r[0]).trim(); })
      .filter(function (v) { return v && !/名称/.test(v); });
  }

  function applyValidations(info, data) {
    if (!data || !data.validations) return;
    data.validations.forEach(function (item) {
      Object.keys(info.columns).forEach(function (key) {
        if (info.columns[key] + 1 === item.col) info.options[key] = item.values;
      });
    });
  }

  function enrichTemplateOptions(info, buffer) {
    if (hasDesktop()) {
      return window.pywebview.api
        .template_options({ template: toBase64(buffer) })
        .then(function (data) { applyValidations(info, data); })
        .catch(function () { /* 读取失败时保持已有选项 */ });
    }
    return Promise.resolve();
  }

  function loadXlsTemplate(buffer, file) {
    var wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });
    if (!wb.SheetNames.length) throw new Error('工作簿里没有工作表');
    var sheetName = wb.SheetNames[0];
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false });
    var info = FF.templateFromRows(rows, { building: optionsFromXls(wb) }, sheetName);
    if (!info) throw new Error('没找到表头（需要有「房屋地址」「楼宇名称」这两列）');
    state.template = { fileName: file.name, buffer: buffer, info: info };
    state.overrides = {};
    state.page = 0;
    state.suffixes = { unitSuffix: '', floorSuffix: '', roomSuffix: '' };
    state.suffixOverrides = {};
    return enrichTemplateOptions(info, buffer).then(function () {
      if (state.mode === 'building') autoMap(false);
      renderAll();
      markLoaded('template');
    });
  }

  function loadTemplate(file) {
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      toast('模板请选择 .xlsx 或 .xls', true);
      return;
    }
    file.arrayBuffer().then(function (buffer) {
      if (/\.xls$/i.test(file.name)) {
        loadXlsTemplate(buffer, file);
        return;
      }
      var wb = new ExcelJS.Workbook();
      return wb.xlsx.load(buffer).then(function () {
        var info = FF.detectTemplate(wb);
        if (!info) throw new Error('没找到表头（需要有「房屋地址」「楼宇名称」这两列）');
        state.template = { fileName: file.name, buffer: buffer, info: info };
        state.overrides = {};
        state.page = 0;
        state.suffixes = { unitSuffix: '', floorSuffix: '', roomSuffix: '' };
        state.suffixOverrides = {};
        if (state.mode === 'building') autoMap(false);
        renderAll();
        markLoaded('template');
      });
    }).catch(function (err) {
      console.error(err);
      toast('模板读取失败：' + err.message, true);
    });
  }

  /* ---------------- 读取房间清单 ---------------- */

  function loadRooms(file) {
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name)) {
      toast('请选择 .xlsx 文件（.xls 请先另存为 .xlsx）', true);
      return;
    }
    file.arrayBuffer().then(function (buffer) {
      var wb = new ExcelJS.Workbook();
      return wb.xlsx.load(buffer).then(function () {
        var ws = wb.worksheets[0];
        if (!ws) throw new Error('工作簿里没有工作表');
        var parsed = FF.parseRoomSheet(ws);
        if (!parsed.rooms.length) throw new Error('没有读到房间，请检查表头是否包含「楼层」');
        state.fileName = file.name;
        state.rooms = parsed.rooms;
        state.buildings = parsed.buildings;
        state.mode = parsed.mode;
        state.overrides = {};
        state.page = 0;
        state.filterBuilding = '';
        state.filterFloor = '';
        state.query = '';
        $('#searchInput').value = '';
        $('#card-mapping').hidden = state.mode === 'single';
        if (state.mode === 'building' && state.template) autoMap(false);
        renderAll();
        markLoaded('rooms');
        if (!$('#prefixInput').value.trim()) $('#prefixInput').focus();
      });
    }).catch(function (err) {
      console.error(err);
      toast('读取失败：' + err.message, true);
    });
  }

  /* ---------------- 自动匹配 ---------------- */

  function autoMap(notify) {
    var options = opt('building');
    state.mapping = FF.suggestMapping(state.buildings.map(function (b) { return b.name; }), options);
    if (notify) {
      renderMapping();
      renderPreview();
      renderExportState();
      toast('已按栋号重新匹配，请逐条确认');
    }
  }

  /* ---------------- 渲染 ---------------- */

  function renderAll() {
    renderTemplateMeta();
    renderMeta();
    renderMapping();
    renderFilters();
    renderSelectors();
    renderPreview();
    renderExportState();
  }

  function fillSuffix(selector, key) {
    var el = $(selector);
    var options = opt(key);
    var current = state.suffixes[key] || '';
    el.innerHTML = '<option value="">空</option>' + options.map(function (value) {
      return '<option value="' + escapeHtml(value) + '"' +
        (value === current ? ' selected' : '') + '>' + escapeHtml(value) + '</option>';
    }).join('');
    el.value = options.indexOf(current) !== -1 ? current : '';
    el.disabled = !options.length;
  }

  function renderSelectors() {
    var sites = siteList();
    var bulk = $('#bulkSite');
    var current = bulk.value;
    bulk.innerHTML = sites.map(function (site) {
      return '<option' + (site === current ? ' selected' : '') + '>' + escapeHtml(site) + '</option>';
    }).join('');
    if (sites.indexOf(current) === -1) bulk.value = sites[0];

    fillSuffix('#bulkUnitSuffix', 'unitSuffix');
    fillSuffix('#bulkFloorSuffix', 'floorSuffix');
    fillSuffix('#bulkRoomSuffix', 'roomSuffix');
  }

  function updateStepper() {
    var single = state.mode === 'single' && state.rooms.length > 0;
    $('#stepMapping').hidden = single;
    $('#sepMapping').hidden = single;
    $('#stepPreview').querySelector('i').textContent = single ? '3' : '4';

    var mapped = state.mode !== 'building' || !state.buildings.length || state.buildings.every(function (b) {
      return !!state.mapping[b.name];
    });
    var done = {
      1: !!state.template,
      2: state.rooms.length > 0,
      3: single || mapped,
      4: !$('#btnExport').disabled
    };
    var active = 0;
    for (var i = 1; i <= 4; i++) {
      if (!done[i]) { active = i; break; }
    }
    [['#stepTemplate', 1], ['#stepRooms', 2], ['#stepMapping', 3], ['#stepPreview', 4]].forEach(function (pair) {
      var el = $(pair[0]);
      if (!el || el.hidden) return;
      el.classList.toggle('done', !!done[pair[1]]);
      el.classList.toggle('active', active === pair[1]);
    });
  }

  function renderTemplateMeta() {
    var meta = $('#templateMeta');
    var optLine = $('#templateOptions');
    var pill = $('#templatePill');
    if (!state.template) {
      meta.hidden = true;
      optLine.hidden = true;
      pill.hidden = true;
      return;
    }
    var info = state.template.info;
    var cols = Object.keys(info.columns);
    meta.hidden = false;
    meta.innerHTML =
      '<span><b>' + escapeHtml(state.template.fileName) + '</b></span>' +
      '<span>工作表 ' + escapeHtml(info.sheetName) + '</span>' +
      '<span>表头第 ' + info.headerRow + ' 行</span>' +
      '<span>识别到 ' + cols.length + ' 列</span>';

    pill.hidden = false;
    pill.textContent = cols.length + ' 列';

    var buildingOptions = opt('building');
    var siteOptions = opt('site');
    var suffixCount = ['unitSuffix', 'floorSuffix', 'roomSuffix']
      .reduce(function (n, key) { return n + (opt(key).length ? 1 : 0); }, 0);

    optLine.hidden = false;
    optLine.innerHTML = '<span class="opt-label">楼宇名称</span>' +
      (buildingOptions.length
        ? buildingOptions.map(function (o) { return '<span class="opt-chip">' + escapeHtml(o) + '</span>'; }).join('')
        : '<span class="hint">未识别到楼宇选项</span>') +
      (siteOptions.length ? '<span class="opt-label">安装点 ' + siteOptions.length + ' 项</span>' : '') +
      (suffixCount ? '<span class="opt-label">后缀 ' + suffixCount + ' 列</span>' : '');
  }

  function renderMeta() {
    var meta = $('#roomMeta');
    var pill = $('#roomPill');
    if (!state.rooms.length) {
      meta.hidden = true;
      pill.hidden = true;
      return;
    }
    var units = state.buildings.reduce(function (n, b) { return n + b.units.length; }, 0);
    meta.hidden = false;
    meta.innerHTML =
      '<span><b>' + escapeHtml(state.fileName) + '</b></span>' +
      (state.mode === 'single'
        ? '<span>不分楼栋</span>'
        : '<span>' + state.buildings.length + ' 栋</span><span>' + units + ' 个单元</span>') +
      '<span><b>' + state.rooms.length + '</b> 个房间</span>';
    pill.hidden = false;
    pill.textContent = state.rooms.length + ' 间';
  }

  function duplicates() {
    var count = {};
    var dup = {};
    Object.keys(state.mapping).forEach(function (b) {
      var v = state.mapping[b];
      if (!v) return;
      count[v] = (count[v] || 0) + 1;
      if (count[v] > 1) dup[v] = true;
    });
    return dup;
  }

  function renderMapping() {
    var body = $('#mappingBody');
    var empty = $('#mappingEmpty');
    var pill = $('#mappingPill');
    if (state.mode === 'single') {
      $('#card-mapping').hidden = true;
      pill.hidden = true;
      return;
    }
    $('#card-mapping').hidden = false;
    if (!state.buildings.length) {
      body.innerHTML = '';
      empty.hidden = false;
      pill.hidden = true;
      return;
    }
    empty.hidden = true;
    var mappedCount = state.buildings.filter(function (b) { return !!state.mapping[b.name]; }).length;
    pill.hidden = false;
    pill.textContent = mappedCount + ' / ' + state.buildings.length;
    var dup = duplicates();
    body.innerHTML = state.buildings.map(function (b) {
      var value = state.mapping[b.name] || '';
      var status = !value
        ? '<span class="status bad"><i></i>未对应</span>'
        : dup[value]
          ? '<span class="status bad"><i></i>重复</span>'
          : '<span class="status ok"><i></i>已对应</span>';
      return '<tr>' +
        '<td class="map-name">' + escapeHtml(b.name) + '</td>' +
        '<td class="map-count">' + b.count + '</td>' +
        '<td class="map-arrow">→</td>' +
        '<td><select data-building="' + escapeHtml(b.name) + '">' + optionsHtml(value) + '</select></td>' +
        '<td>' + status + '</td>' +
        '</tr>';
    }).join('');

    var missing = state.buildings.filter(function (b) { return !state.mapping[b.name]; });
    var dupNames = Object.keys(dup);
    var warn = $('#mappingWarn');
    if (!missing.length && !dupNames.length) {
      warn.hidden = true;
    } else {
      warn.hidden = false;
      warn.className = 'warn-box error';
      warn.innerHTML =
        (missing.length ? '<div>还有 ' + missing.length + ' 栋没有选择楼宇名称：' + missing.map(function (b) { return b.name; }).join('、') + '</div>' : '') +
        (dupNames.length ? '<div>这些楼宇被重复选择：' + dupNames.join('、') + '</div>' : '');
    }
  }

  function renderFilters() {
    var buildingSelect = $('#filterBuilding');
    buildingSelect.innerHTML = '<option value="">全部楼栋</option>' + state.buildings.map(function (b) {
      return '<option' + (state.filterBuilding === b.name ? ' selected' : '') + '>' + escapeHtml(b.name) + '</option>';
    }).join('');

    var floors = unique(state.rooms.map(function (r) { return r.floor; }).filter(function (f) { return f != null; }))
      .sort(function (a, b) { return a - b; });
    var floorSelect = $('#filterFloor');
    floorSelect.innerHTML = '<option value="">全部楼层</option>' + floors.map(function (f) {
      return '<option value="' + f + '"' + (String(state.filterFloor) === String(f) ? ' selected' : '') + '>' + f + ' 层</option>';
    }).join('');
  }

  function buildingNameFor(room) {
    return state.mode === 'single' ? '-' : (state.mapping[room.building] || '');
  }

  function siteOf(room, address) {
    var key = roomKey(room);
    if (state.overrides[key]) return state.overrides[key];
    if (address.indexOf('别墅') !== -1) return '别墅表';
    return '厨房表';
  }

  var EMPTY_SUFFIX = { unitSuffix: '', floorSuffix: '', roomSuffix: '' };

  function suffixFor(room) {
    return state.suffixOverrides[roomKey(room)] || EMPTY_SUFFIX;
  }

  function rowFor(room) {
    var buildingName = buildingNameFor(room);
    var prefix = $('#prefixInput').value.trim();
    var addr = buildingName ? FF.composeAddress(prefix, buildingName, room.unit, room.door) : '';
    return {
      key: roomKey(room),
      address: addr,
      building: buildingName,
      unit: room.unit || '',
      floor: room.floor,
      door: room.door,
      roomNo: FF.splitDoor(room.door, room.floor),
      site: siteOf(room, addr),
      suffix: suffixFor(room),
      missing: !buildingName,
      mismatch: room.floor != null && String(room.door).indexOf(String(room.floor)) !== 0
    };
  }

  function allRows() { return state.rooms.map(rowFor); }

  function applyFilters(rows) {
    var q = state.query.trim().toLowerCase();
    return rows.filter(function (r) {
      if (state.filterBuilding && r.key.indexOf(state.filterBuilding + '|') !== 0) return false;
      if (state.filterFloor !== '' && String(r.floor) !== String(state.filterFloor)) return false;
      if (q) {
        var hay = (r.address + ' ' + r.door + ' ' + r.building + ' ' + r.unit).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });
  }

  function siteList() {
    var options = opt('site');
    return options.length ? options : ['厨房表', '阳台表', '别墅表'];
  }

  function suffixSelectHtml(key, field, value) {
    var options = opt(field);
    var list = value && options.indexOf(value) === -1 ? [value].concat(options) : options;
    var html = '<option value="">—</option>' + list.map(function (item) {
      return '<option value="' + escapeHtml(item) + '"' +
        (item === value ? ' selected' : '') + '>' + escapeHtml(item) + '</option>';
    }).join('');
    return '<select data-suffix="' + field + '" data-key="' + escapeHtml(key) + '">' + html + '</select>';
  }

  function statsHtml(all, rows) {
    var counts = {};
    var suffixRows = 0;
    all.forEach(function (r) {
      counts[r.site] = (counts[r.site] || 0) + 1;
      if (r.suffix.unitSuffix || r.suffix.floorSuffix || r.suffix.roomSuffix) suffixRows++;
    });
    var mismatch = all.filter(function (r) { return r.mismatch; }).length;

    var top = Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a]; });
    var shown = top.slice(0, 4).map(function (name) {
      return '<span class="stat">' + escapeHtml(name) + ' <b>' + counts[name] + '</b></span>';
    }).join('');
    if (top.length > 4) {
      var rest = top.slice(4).reduce(function (n, name) { return n + counts[name]; }, 0);
      shown += '<span class="stat">其他 <b>' + rest + '</b></span>';
    }
    if (suffixRows) {
      shown += '<span class="stat">已设后缀 <b>' + suffixRows + '</b></span>';
    }

    return '<span class="stat">全部 <b>' + all.length + '</b> 行</span>' +
      '<span class="stat">当前显示 <b>' + rows.length + '</b> 行</span>' +
      shown +
      (mismatch ? '<span class="stat danger">楼层与门牌不一致 <b>' + mismatch + '</b></span>' : '');
  }

  function refreshStats() {
    var all = allRows();
    $('#statRow').innerHTML = statsHtml(all, applyFilters(all));
  }

  function renderPager(total) {
    var pager = $('#pager');
    if (!total) {
      pager.hidden = true;
      pager.innerHTML = '';
      return;
    }
    var pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
    if (state.page >= pages) state.page = pages - 1;
    pager.hidden = false;
    pager.innerHTML =
      '<button class="ghost btn-sm" data-page="prev"' + (state.page === 0 ? ' disabled' : '') + '>上一页</button>' +
      '<span class="pager-info">第 ' + (state.page + 1) + ' / ' + pages + ' 页 · 共 ' + total + ' 行</span>' +
      '<button class="ghost btn-sm" data-page="next"' + (state.page >= pages - 1 ? ' disabled' : '') + '>下一页</button>';
  }

  function renderPreview() {
    var body = $('#previewBody');
    var empty = $('#previewEmpty');
    if (!state.rooms.length) {
      body.innerHTML = '';
      empty.hidden = false;
      $('#statRow').innerHTML = '';
      $('#pager').hidden = true;
      return;
    }
    empty.hidden = true;
    var all = allRows();
    var rows = applyFilters(all);
    var sites = siteList();

    var pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    if (state.page >= pages) state.page = pages - 1;
    if (state.page < 0) state.page = 0;
    var start = state.page * PAGE_SIZE;
    var pageRows = rows.slice(start, start + PAGE_SIZE);

    body.innerHTML = pageRows.map(function (r, i) {
      var list = sites.indexOf(r.site) === -1 ? [r.site].concat(sites) : sites;
      var siteOptions = list.map(function (s) {
        return '<option' + (s === r.site ? ' selected' : '') + '>' + escapeHtml(s) + '</option>';
      }).join('');
      return '<tr class="' + (r.mismatch ? 'row-warn' : '') + '" data-key="' + escapeHtml(r.key) + '">' +
        '<td class="c-idx">' + (start + i + 1) + '</td>' +
        '<td class="c-addr" title="' + escapeHtml(r.address) + '">' + escapeHtml(r.address) + '</td>' +
        '<td>' + escapeHtml(r.building || '—') + '</td>' +
        '<td class="c-unit">' + escapeHtml(r.unit || '') + '</td>' +
        '<td class="c-unit-suffix">' + suffixSelectHtml(r.key, 'unitSuffix', r.suffix.unitSuffix) + '</td>' +
        '<td class="c-floor">' + escapeHtml(r.floor == null ? '' : r.floor) + '</td>' +
        '<td class="c-floor-suffix">' + suffixSelectHtml(r.key, 'floorSuffix', r.suffix.floorSuffix) + '</td>' +
        '<td class="c-room">' + escapeHtml(r.roomNo || '') + '</td>' +
        '<td class="c-room-suffix">' + suffixSelectHtml(r.key, 'roomSuffix', r.suffix.roomSuffix) + '</td>' +
        '<td class="c-site"><select data-key="' + escapeHtml(r.key) + '">' + siteOptions + '</select></td>' +
        '</tr>';
    }).join('');

    renderPager(rows.length);
    $('#statRow').innerHTML = statsHtml(all, rows);
  }

  function renderExportState() {
    var missingMapping = state.mode === 'building' &&
      state.buildings.some(function (b) { return !state.mapping[b.name]; });
    var dup = state.mode === 'building' && Object.keys(duplicates()).length > 0;
    var noOptions = state.template && state.mode === 'building' && !opt('building').length;
    var prefix = $('#prefixInput').value.trim();
    var ready = state.template && state.rooms.length && !missingMapping && !dup && !noOptions && prefix;

    $('#btnExport').disabled = !ready;
    $('#btnExport').textContent = state.template && state.template.info.source === 'xls' ? '导出 .xls' : '导出 .xlsx';
    $('#exportHint').textContent = !state.template
      ? '未就绪'
      : !state.rooms.length
        ? '未就绪'
        : noOptions ? '模板缺少楼宇选项'
          : missingMapping ? '楼栋对应未完成'
            : dup ? '楼宇重复对应'
              : !prefix ? '项目地址前缀未填写'
                : '共 ' + state.rooms.length + ' 行 · 写入「' + state.template.info.sheetName + '」';
    updateStepper();
  }

  /* ---------------- 导出 ---------------- */

  function valuesForRoom(room) {
    var prefix = $('#prefixInput').value.trim();
    var suffix = suffixFor(room);
    var values = FF.buildRowValues(room, {
      prefix: prefix,
      buildingName: buildingNameFor(room),
      installSite: siteList().indexOf('厨房表') !== -1 ? '厨房表' : siteList()[0],
      unitSuffix: suffix.unitSuffix,
      floorSuffix: suffix.floorSuffix,
      roomSuffix: suffix.roomSuffix
    });
    var key = roomKey(room);
    if (state.overrides[key]) values.site = state.overrides[key];
    return values;
  }

  function toBase64(buffer) {
    var bytes = new Uint8Array(buffer);
    var binary = '';
    var chunk = 0x8000;
    for (var i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }

  function hasDesktop() {
    return !!(window.pywebview && window.pywebview.api);
  }

  function stampText() {
    var d = new Date();
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '_' + pad(d.getHours()) + pad(d.getMinutes());
  }

  function xlsPayload(info) {
    var width = info.columnCount || 0;
    var rows = state.rooms.map(function (room) {
      var values = valuesForRoom(room);
      var row = new Array(width).fill('');
      Object.keys(info.columns).forEach(function (key) {
        var v = values[key];
        if (v != null && v !== '') row[info.columns[key]] = v;
      });
      return row;
    });

    return {
      template: toBase64(state.template.buffer),
      sheetName: info.sheetName,
      headerRow: info.headerRow,
      floorCol: info.columns.floor != null ? info.columns.floor + 1 : 0,
      roomCol: info.columns.room != null ? info.columns.room + 1 : 0,
      rows: rows
    };
  }

  function handleDesktopResult(result) {
    if (!result || result.ok === false) {
      if (result && result.canceled) {
        var err = new Error('已取消');
        err.canceled = true;
        throw err;
      }
      throw new Error((result && result.error) || '保存失败');
    }
    return { savedPath: result.path };
  }

  // 桌面客户端：Python 侧写模板并弹原生另存为
  function exportXlsViaDesktop(info) {
    return window.pywebview.api.export_xls(xlsPayload(info)).then(handleDesktopResult);
  }

  function saveViaDesktop(buffer, ext) {
    return window.pywebview.api
      .save_file(toBase64(buffer), '房产表计安装点导入_' + stampText() + '.' + ext)
      .then(handleDesktopResult);
  }

  function exportWorkbook() {
    if (!state.template || !state.rooms.length) throw new Error('还没有模板或房间清单');
    if (!hasDesktop()) throw new Error('请从桌面客户端打开（双击「表计安装点填表器.exe」）');
    var info = state.template.info;

    if (info.source === 'xls') return exportXlsViaDesktop(info);

    var wb = new ExcelJS.Workbook();
    return wb.xlsx.load(state.template.buffer.slice(0)).then(function () {
      var ws = wb.getWorksheet(info.sheetName);
      if (!ws) throw new Error('模板里找不到工作表：' + info.sheetName);

      state.rooms.forEach(function (room, i) {
        var values = valuesForRoom(room);
        var rowNumber = info.headerRow + 1 + i;
        Object.keys(info.columns).forEach(function (key) {
          var col = info.columns[key] + 1;
          var value = values[key];
          var cell = ws.getCell(rowNumber, col);
          cell.value = value === '' ? null : (key === 'floor' ? Number(value) : value);
        });
      });

      // 模板预置的空行数量不固定：多余的删掉，保证输出正好是表头 + N 行
      var lastNeeded = info.headerRow + state.rooms.length;
      if (ws.rowCount > lastNeeded) {
        ws.spliceRows(lastNeeded + 1, ws.rowCount - lastNeeded);
      }

      return wb.xlsx.writeBuffer();
    }).then(function (buffer) {
      return saveViaDesktop(buffer, 'xlsx');
    });
  }

  function download() {
    var btn = $('#btnExport');
    btn.disabled = true;
    btn.textContent = '生成中…';

    exportWorkbook().then(function (result) {
      toast('已保存：' + result.savedPath);
    }).catch(function (err) {
      if (err && err.canceled) { toast('已取消'); return; }
      console.error(err);
      toast('导出失败：' + err.message, true);
    }).then(function () {
      renderExportState();
    });
  }

  function pad(n) { return (n < 10 ? '0' : '') + n; }

  /* ---------------- 事件 ---------------- */

  function bindDrop(dropEl, inputEl, handler) {
    ['dragenter', 'dragover'].forEach(function (evt) {
      dropEl.addEventListener(evt, function (e) { e.preventDefault(); dropEl.classList.add('over'); });
    });
    ['dragleave', 'drop'].forEach(function (evt) {
      dropEl.addEventListener(evt, function (e) { e.preventDefault(); dropEl.classList.remove('over'); });
    });
    dropEl.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handler(f);
    });
    inputEl.addEventListener('change', function () {
      if (inputEl.files && inputEl.files[0]) handler(inputEl.files[0]);
    });
  }

  function markLoaded(which) {
    var pair = which === 'template'
      ? ['#dropTemplate', '#btnReplaceTemplate']
      : ['#dropRooms', '#btnReplaceRooms'];
    $(pair[0]).hidden = true;
    $(pair[1]).hidden = false;
  }

  var searchTimer = null;

  function enableDragScroll(el) {
    if (!el) return;
    var down = false, moved = false, startX = 0, startLeft = 0;
    el.addEventListener('mousedown', function (e) {
      if (e.target.closest('select,input,button,a')) return;
      down = true;
      moved = false;
      startX = e.pageX;
      startLeft = el.scrollLeft;
      el.classList.add('dragging');
    });
    window.addEventListener('mousemove', function (e) {
      if (!down) return;
      var dx = e.pageX - startX;
      if (Math.abs(dx) > 3) { moved = true; e.preventDefault(); }
      el.scrollLeft = startLeft - dx;
    });
    window.addEventListener('mouseup', function () {
      down = false;
      el.classList.remove('dragging');
    });
    el.addEventListener('click', function (e) {
      if (moved) {
        e.preventDefault();
        e.stopPropagation();
        moved = false;
      }
    }, true);
  }

  function init() {
    $('#btnReset').addEventListener('click', function () { location.reload(); });
    document.querySelectorAll('.step-item').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var target = document.querySelector(btn.getAttribute('data-target'));
        if (target && !target.hidden) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      });
    });

    bindDrop($('#dropTemplate'), $('#fileTemplate'), loadTemplate);
    bindDrop($('#dropRooms'), $('#fileRooms'), loadRooms);
    $('#btnReplaceTemplate').addEventListener('click', function () { $('#fileTemplate').click(); });
    $('#btnReplaceRooms').addEventListener('click', function () { $('#fileRooms').click(); });

    $('#prefixInput').addEventListener('input', function () {
      renderPreview();
      renderExportState();
    });

    $('#mappingBody').addEventListener('change', function (e) {
      var select = e.target.closest('select[data-building]');
      if (!select) return;
      var building = select.getAttribute('data-building');
      if (select.value) state.mapping[building] = select.value;
      else delete state.mapping[building];
      renderMapping();
      renderPreview();
      renderExportState();
    });

    $('#btnAutoMap').addEventListener('click', function () {
      if (!state.buildings.length) { toast('先上传房间清单'); return; }
      autoMap(true);
    });

    $('#filterBuilding').addEventListener('change', function (e) {
      state.filterBuilding = e.target.value;
      state.page = 0;
      renderPreview();
    });

    $('#filterFloor').addEventListener('change', function (e) {
      state.filterFloor = e.target.value;
      state.page = 0;
      renderPreview();
    });

    $('#searchInput').addEventListener('input', function (e) {
      clearTimeout(searchTimer);
      var value = e.target.value;
      $('#btnClearSearch').hidden = !value;
      searchTimer = setTimeout(function () {
        state.query = value;
        state.page = 0;
        renderPreview();
      }, 180);
    });

    $('#btnClearSearch').addEventListener('click', function () {
      $('#searchInput').value = '';
      $('#btnClearSearch').hidden = true;
      state.query = '';
      state.page = 0;
      renderPreview();
      $('#searchInput').focus();
    });

    $('#btnResetFilters').addEventListener('click', function () {
      state.filterBuilding = '';
      state.filterFloor = '';
      state.query = '';
      state.page = 0;
      $('#searchInput').value = '';
      $('#btnClearSearch').hidden = true;
      renderFilters();
      renderPreview();
    });

    $('#pager').addEventListener('click', function (e) {
      var btn = e.target.closest('button[data-page]');
      if (!btn || btn.disabled) return;
      state.page += btn.getAttribute('data-page') === 'next' ? 1 : -1;
      if (state.page < 0) state.page = 0;
      renderPreview();
      $('#previewScroll').scrollTop = 0;
    });

    $('#previewBody').addEventListener('change', function (e) {
      var select = e.target.closest('select[data-key]');
      if (!select) return;
      var key = select.getAttribute('data-key');
      var field = select.getAttribute('data-suffix');
      if (field) {
        var suffix = state.suffixOverrides[key] || { unitSuffix: '', floorSuffix: '', roomSuffix: '' };
        suffix[field] = select.value;
        state.suffixOverrides[key] = suffix;
      } else {
        state.overrides[key] = select.value;
      }
      refreshStats();
    });

    $('#btnBulk').addEventListener('click', function () {
      var rows = applyFilters(allRows());
      if (!rows.length) { toast('当前筛选没有行'); return; }
      var site = $('#bulkSite').value;
      rows.forEach(function (r) { state.overrides[r.key] = site; });
      renderPreview();
      toast('已设置 ' + rows.length + ' 行');
    });

    $('#btnBulkSuffix').addEventListener('click', function () {
      var rows = applyFilters(allRows());
      if (!rows.length) { toast('当前筛选没有行'); return; }
      var values = {
        unitSuffix: $('#bulkUnitSuffix').value,
        floorSuffix: $('#bulkFloorSuffix').value,
        roomSuffix: $('#bulkRoomSuffix').value
      };
      rows.forEach(function (r) {
        state.suffixOverrides[r.key] = {
          unitSuffix: values.unitSuffix,
          floorSuffix: values.floorSuffix,
          roomSuffix: values.roomSuffix
        };
      });
      renderPreview();
      toast('已设置 ' + rows.length + ' 行后缀');
    });

    ['#bulkUnitSuffix', '#bulkFloorSuffix', '#bulkRoomSuffix'].forEach(function (selector, index) {
      var key = ['unitSuffix', 'floorSuffix', 'roomSuffix'][index];
      $(selector).addEventListener('change', function (e) {
        state.suffixes[key] = e.target.value;
      });
    });

    $('#btnExport').addEventListener('click', download);

    enableDragScroll($('#previewScroll'));

    renderExportState();

    setTimeout(function () {
      if (!hasDesktop()) {
        var warn = document.createElement('div');
        warn.className = 'browser-warn';
        warn.textContent = '请在桌面客户端中使用（双击「表计安装点填表器.exe」）';
        document.body.insertBefore(warn, document.body.firstChild);
      }
    }, 3000);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();






