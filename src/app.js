/* Coftech Excel — aplicación principal. */
(function () {
  'use strict';

  const F = window.CEFormulas;
  const N = window.CEFormat;
  const IO = window.CEIO;
  const HF = window.HyperFormula;
  const SSF = window.XLSX.SSF;

  const $ = id => document.getElementById(id);
  const key = (r, c) => r + ',' + c;
  const cellName = (r, c) => F.numToCol(c) + (r + 1);
  const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

  const HEAD_W = 44;
  const UNDO_LIMIT = 100;

  // ---------- Motor de fórmulas en español ----------
  const LANG = JSON.parse(JSON.stringify(HF.languages.esES));
  LANG.langCode = 'esAR';
  LANG.errors.NA = '#N/A';
  LANG.errors.CYCLE = '#¡CIRCULAR!';
  HF.registerLanguage('esAR', LANG);
  F.setLanguages(HF.getLanguage('enGB').functions, LANG.functions);

  const HF_CONFIG = {
    licenseKey: 'gpl-v3',
    language: 'esAR',
    functionArgSeparator: ';',
    decimalSeparator: ',',
    thousandSeparator: '',
    dateFormats: ['DD/MM/YYYY', 'DD/MM/YY'],
    timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
    maxRows: 1048576,
    maxColumns: 16384
  };

  // ---------- Estado ----------
  const S = {
    hf: null,
    sheets: [],        // { name, styles: Map, merges, colW, rowH, freezeRow, freezeCol, filter, hidden: Set }
    active: 0,
    fileName: '',
    loaded: false,
    dirty: false,
    sel: { r0: 0, c0: 0, r1: 0, c1: 0, ar: 0, ac: 0 },
    anchor: { r: 0, c: 0 },
    zoom: 1,
    undo: [],
    redo: [],
    clip: null,
    edit: null,
    drag: null,
    L: { colX: [], colW: [], rowY: [], rowH: [], rows: 0, cols: 0, headH: 22, merged: new Map(), covered: new Map() }
  };

  const sheet = () => S.sheets[S.active];
  const sid = (sh) => S.hf.getSheetId((sh || sheet()).name);
  const addr = (r, c, sh) => ({ sheet: sid(sh), row: r, col: c });
  const getStyle = (r, c, sh) => (sh || sheet()).styles.get(key(r, c));
  const colWidth = c => sheet().colW[c] || IO.DEFAULT_COL_W;
  const rowHeight = r => sheet().rowH[r] || IO.DEFAULT_ROW_H;
  // Tamaños en pantalla: el ancho y el alto guardados se multiplican por el zoom.
  const ZOOM_STEPS = [0.5, 0.6, 0.7, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2];
  const headW = () => Math.round(HEAD_W * S.zoom);
  const colPx = c => Math.max(1, Math.round(colWidth(c) * S.zoom));
  const rowPx = r => Math.max(1, Math.round(rowHeight(r) * S.zoom));
  const fontPx = size => Math.round((size ? size * 14 / 11 : 14) * S.zoom * 10) / 10;

  function buildHF(sheetsData) {
    if (S.hf) S.hf.destroy();
    const obj = {};
    sheetsData.forEach(s => { obj[s.name] = s.data; });
    S.hf = HF.buildFromSheets(obj, HF_CONFIG);
    S.hf.addNamedExpression('VERDADERO', '=VERDADERO()');
    S.hf.addNamedExpression('FALSO', '=FALSO()');
  }

  function cellValue(r, c, sh) {
    const v = S.hf.getCellValue(addr(r, c, sh));
    return v === undefined ? null : v;
  }
  function cellRaw(r, c, sh) {
    const v = S.hf.getCellSerialized(addr(r, c, sh));
    return v === undefined ? null : v;
  }
  const isError = v => v && typeof v === 'object' && 'value' in v && 'type' in v;
  const isFormula = raw => typeof raw === 'string' && raw[0] === '=';

  function displayText(r, c, sh) {
    const v = cellValue(r, c, sh);
    if (isError(v)) return v.value;
    const st = getStyle(r, c, sh);
    return N.display(v, st && st.fmt, SSF);
  }

  /** Texto que se muestra en la barra de fórmulas o al editar. */
  function editTextOf(r, c) {
    const raw = cellRaw(r, c);
    if (isFormula(raw)) return raw;
    const v = cellValue(r, c);
    if (typeof v === 'string') {
      const p = N.parseInput(v);
      return p.text && v[0] !== "'" ? v : "'" + v;
    }
    const st = getStyle(r, c);
    return N.editText(v, st && st.fmt, SSF);
  }

  /** Escribe en el motor y anota las filas que cambiaron en la hoja visible. */
  function setContents(a, data) {
    const changes = S.hf.setCellContents(a, data);
    if (S.touched) {
      const id = sid();
      if (a.sheet === id) data.forEach((row, i) => S.touched.add(a.row + i));
      changes.forEach(ch => { if (ch.address && ch.address.sheet === id) S.touched.add(ch.address.row); });
    }
    return changes;
  }

  function setStyleProps(r, c, props, sh) {
    sh = sh || sheet();
    if (S.touched && sh === sheet()) { S.touched.add(r); S.touched.add(r - 1); }
    const k = key(r, c);
    const st = Object.assign({}, sh.styles.get(k) || {});
    Object.keys(props).forEach(p => {
      if (props[p] === undefined || props[p] === null || props[p] === false || props[p] === '') delete st[p];
      else st[p] = props[p];
    });
    if (Object.keys(st).length) sh.styles.set(k, st);
    else sh.styles.delete(k);
  }

  /** Guarda lo que el usuario escribió en una celda. */
  function writeInput(r, c, text) {
    const p = N.parseInput(text);
    let content;
    if (p.value === null) content = null;
    else if (p.formula) content = p.value;
    else if (p.text) content = "'" + p.value;
    else content = p.value;
    setContents(addr(r, c), [[content]]);
    if (p.format) {
      const st = getStyle(r, c);
      const cur = st && st.fmt;
      const weak = /^#,##0/.test(p.format) && !/\$/.test(p.format);
      if (!cur || !weak) setStyleProps(r, c, { fmt: p.format });
    }
  }

  // ---------- Deshacer / rehacer ----------
  function snapshot() {
    return {
      sheets: S.sheets.map(sh => ({
        name: sh.name,
        data: S.hf.getSheetSerialized(sid(sh)),
        styles: Array.from(sh.styles.entries()).map(([k, v]) => [k, Object.assign({}, v)]),
        merges: sh.merges.map(m => Object.assign({}, m)),
        colW: Object.assign({}, sh.colW),
        rowH: Object.assign({}, sh.rowH),
        freezeRow: sh.freezeRow,
        freezeCol: sh.freezeCol,
        filter: sh.filter ? cloneFilter(sh.filter) : null,
        hidden: Array.from(sh.hidden)
      })),
      active: S.active,
      sel: Object.assign({}, S.sel)
    };
  }

  function cloneFilter(f) {
    const crit = {};
    Object.keys(f.crit).forEach(c => { crit[c] = new Set(f.crit[c]); });
    return { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1, crit };
  }

  function restore(snap) {
    buildHF(snap.sheets);
    S.sheets = snap.sheets.map(s => ({
      name: s.name,
      styles: new Map(s.styles),
      merges: s.merges,
      colW: s.colW,
      rowH: s.rowH,
      freezeRow: s.freezeRow,
      freezeCol: s.freezeCol,
      filter: s.filter,
      hidden: new Set(s.hidden)
    }));
    S.active = clamp(snap.active, 0, S.sheets.length - 1);
    S.sel = snap.sel;
    S.anchor = { r: S.sel.ar, c: S.sel.ac };
    renderTabs();
    renderGrid();
  }

  /** Ejecuta un cambio que se puede deshacer. */
  function commit(fn, opts) {
    const snap = snapshot();
    let result;
    S.touched = new Set();
    try {
      result = fn();
    } catch (e) {
      S.touched = null;
      console.error(e);
      toast('No se pudo completar la acción: ' + (e.message || e), 'error');
      restore(snap);
      return;
    }
    const touched = S.touched;
    S.touched = null;
    if (result === false) return;
    S.undo.push(snap);
    if (S.undo.length > UNDO_LIMIT) S.undo.shift();
    S.redo = [];
    setDirty(true);
    if (opts && opts.partial && sheetShape() === S.L.shape && Math.max(-1, ...touched) < S.L.rows - 5) renderRows(touched);
    else if (!opts || !opts.noRender) renderGrid();
    if (opts && opts.tabs) renderTabs();
  }

  function undo() {
    if (!S.undo.length) return toast('No hay nada para deshacer.');
    S.redo.push(snapshot());
    restore(S.undo.pop());
    setDirty(true);
  }
  function redo() {
    if (!S.redo.length) return toast('No hay nada para rehacer.');
    S.undo.push(snapshot());
    restore(S.redo.pop());
    setDirty(true);
  }

  function setDirty(v) {
    S.dirty = v;
    $('dirtyFlag').hidden = !v;
  }

  // ---------- Combinadas ----------
  function mergeAt(r, c) {
    return sheet().merges.find(m => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1) || null;
  }

  function expandRange(rg) {
    let { r0, c0, r1, c1 } = rg;
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of sheet().merges) {
        if (m.r1 < r0 || m.r0 > r1 || m.c1 < c0 || m.c0 > c1) continue;
        if (m.r0 < r0) { r0 = m.r0; changed = true; }
        if (m.c0 < c0) { c0 = m.c0; changed = true; }
        if (m.r1 > r1) { r1 = m.r1; changed = true; }
        if (m.c1 > c1) { c1 = m.c1; changed = true; }
      }
    }
    return { r0, c0, r1, c1 };
  }

  // ---------- Dibujo de la cuadrícula ----------
  function usedExtent(sh) {
    sh = sh || sheet();
    const dim = S.hf.getSheetDimensions(sid(sh));
    let rows = dim.height, cols = dim.width;
    sh.styles.forEach((v, k) => {
      const [r, c] = k.split(',').map(Number);
      if (r + 1 > rows) rows = r + 1;
      if (c + 1 > cols) cols = c + 1;
    });
    sh.merges.forEach(m => { rows = Math.max(rows, m.r1 + 1); cols = Math.max(cols, m.c1 + 1); });
    return { rows, cols };
  }

  function cellStyleCss(st, r, c) {
    let css = '';
    if (!st) return css;
    if (st.b) css += 'font-weight:700;';
    if (st.i) css += 'font-style:italic;';
    if (st.u) css += 'text-decoration:underline;';
    if (st.color) css += 'color:' + st.color + ';';
    if (st.bg) css += 'background:' + st.bg + ';';
    if (st.size) css += 'font-size:' + fontPx(st.size) + 'px;';
    if (st.align) css += 'text-align:' + st.align + ';';
    return css;
  }

  /** HTML de una fila de la cuadrícula. */
  function rowHTML(r, sh, valAt) {
    const L = S.L;
    const filt = sh.filter;
    const parts = [];
    const hidden = sh.hidden.has(r);
    const frow = sh.freezeRow && r === 0;
    parts.push('<tr data-r="' + r + '"' + (frow ? ' class="frow"' : '') + ' style="height:' + rowPx(r) + 'px' + (hidden ? ';display:none' : '') + '">');
    parts.push('<th data-hr="' + r + '"' + (r >= S.sel.r0 && r <= S.sel.r1 ? ' class="hsel"' : '') + '>' + (r + 1) + '<div class="rzr" data-rzr="' + r + '"></div></th>');
    for (let c = 0; c < L.cols; c++) {
      const k = key(r, c);
      if (L.covered.has(k)) continue;
      const v = valAt(r, c);
      const st = sh.styles.get(k);
      let cls = '';
      let text = '';
      if (isError(v)) { text = v.value; cls = 'err'; }
      else if (v !== null && v !== '') {
        text = N.display(v, st && st.fmt, SSF);
        if (typeof v === 'number') cls = 'num';
        else if (typeof v === 'boolean') cls = 'ctr';
      }
      let css = cellStyleCss(st, r, c);
      // Bordes: el inferior y el derecho de cada celda; el superior/izquierdo se toman del vecino.
      const below = sh.styles.get(key(r + 1, c));
      const right = sh.styles.get(key(r, c + 1));
      const m = L.merged.get(k);
      if ((st && st.bb) || (below && below.bt && !m)) css += 'border-bottom-color:#000;';
      if ((st && st.br) || (right && right.bl && !m)) css += 'border-right-color:#000;';
      if (r === 0 && st && st.bt) css += 'box-shadow:inset 0 1px 0 #000;';
      if (c === 0 && st && st.bl) css += 'box-shadow:inset 1px 0 0 #000;';
      if (text && !cls && !m && (!st || !st.align || st.align === 'left') && valAt(r, c + 1) === null && !L.covered.has(key(r, c + 1)) && !(sh.freezeCol && c === 0)) cls += ' ov';
      if (frow) cls += ' fr';
      if (sh.freezeCol && c === 0) cls += ' fc';
      let span = '';
      if (m) {
        span = (m.r1 > m.r0 ? ' rowspan="' + (m.r1 - m.r0 + 1) + '"' : '') + (m.c1 > m.c0 ? ' colspan="' + (m.c1 - m.c0 + 1) + '"' : '');
        if (!st || !st.align) css += 'text-align:center;';
      }
      let inner = esc(text);
      if (filt && r === filt.r0 && c >= filt.c0 && c <= filt.c1) {
        const on = filt.crit[c] ? ' active' : '';
        inner = '<span class="fbtn' + on + '" data-fc="' + c + '" title="Filtrar esta columna">▼</span>' + inner;
      }
      parts.push('<td data-r="' + r + '" data-c="' + c + '"' + span + (cls ? ' class="' + cls.trim() + '"' : '') + (css ? ' style="' + css + '"' : '') + '>' + inner + '</td>');
    }
    parts.push('</tr>');
    return parts.join('');
  }

  function renderGrid() {
    if (!S.loaded) return;
    const sh = sheet();
    const ext = usedExtent(sh);
    const rows = Math.max(100, ext.rows + 30, (S.sel.wholeCols ? 0 : S.sel.r1) + 20);
    const cols = Math.max(26, ext.cols + 5, (S.sel.wholeRows ? 0 : S.sel.c1) + 5);
    if (S.sel.wholeCols) S.sel.r1 = rows - 1;
    if (S.sel.wholeRows) S.sel.c1 = cols - 1;
    const L = S.L;
    L.rows = rows; L.cols = cols;
    L.merged = new Map(); L.covered = new Map();
    sh.merges.forEach(m => {
      L.merged.set(key(m.r0, m.c0), m);
      for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) L.covered.set(key(r, c), m);
    });

    const values = S.hf.getSheetValues(sid(sh));
    const valAt = (r, c) => (values[r] && values[r][c] !== undefined ? values[r][c] : null);

    const parts = [];
    parts.push('<colgroup><col style="width:' + headW() + 'px">');
    L.colW = [];
    for (let c = 0; c < cols; c++) {
      L.colW[c] = colPx(c);
      parts.push('<col style="width:' + L.colW[c] + 'px">');
    }
    parts.push('</colgroup><thead><tr><th class="corner" data-corner="1"></th>');
    for (let c = 0; c < cols; c++) parts.push('<th data-hc="' + c + '">' + F.numToCol(c) + '<div class="rz" data-rzc="' + c + '"></div></th>');
    parts.push('</tr></thead><tbody></tbody>');
    const grid = $('grid');
    grid.innerHTML = parts.join('');
    let totalW = headW();
    L.colX = [];
    for (let c = 0; c < cols; c++) { L.colX[c] = totalW; totalW += L.colW[c]; }
    grid.style.width = totalW + 'px';
    L.headH = grid.tHead.rows[0].offsetHeight;
    grid.style.setProperty('--hh', L.headH + 'px');
    grid.style.setProperty('--hw', headW() + 'px');
    $('gridWrap').style.setProperty('--z', S.zoom);
    // Alturas previstas; las filas dibujadas se miden y corrigen en renderWindow.
    L.rowH = [];
    for (let r = 0; r < rows; r++) L.rowH[r] = sh.hidden.has(r) ? 0 : rowPx(r);
    L.rowY = [];
    computeRowY();
    L.thCols = Array.from(grid.tHead.rows[0].cells).slice(1);
    L.shape = sheetShape();
    L.values = values;
    renderWindow(true);
    L.values = null;
    if (S.clip && S.clip.sheet === sh.name && !S.clip.done) showCopyBox(); else $('copyBox').hidden = true;
  }

  /** Fila que está en la coordenada vertical y (búsqueda binaria). */
  function rowAtY(y) {
    const L = S.L;
    let lo = 0, hi = L.rows - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (L.rowY[mid] <= y) lo = mid; else hi = mid - 1;
    }
    return lo;
  }

  const WIN_BUFFER = 40;

  /** Dibuja solo las filas cercanas a la parte visible (las demás se reemplazan por espacio vacío). */
  function renderWindow(force) {
    if (!S.loaded) return;
    const L = S.L;
    const sh = sheet();
    const sc = $('scroller');
    const top = sc.scrollTop, bottom = top + Math.max(sc.clientHeight, 600);
    let start = Math.max(0, rowAtY(top) - WIN_BUFFER);
    let end = Math.min(L.rows - 1, rowAtY(bottom) + WIN_BUFFER);
    if (!force && L.win && rowAtY(top) >= L.win[0] + 5 && rowAtY(bottom) <= L.win[1] - 5) return;
    // Una combinada no puede quedar cortada por el borde de la ventana.
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of sh.merges) {
        if (m.r0 < start && m.r1 >= start) { start = m.r0; changed = true; }
        if (m.r0 <= end && m.r1 > end) { end = m.r1; changed = true; }
      }
    }
    L.win = [start, end];
    const values = L.values;
    const valAt = values
      ? (r, c) => (values[r] && values[r][c] !== undefined ? values[r][c] : null)
      : (r, c) => { const v = S.hf.getCellValue({ sheet: sid(sh), row: r, col: c }); return v === undefined ? null : v; };
    const parts = [];
    const spacer = h => '<tr class="spacer" style="height:' + h + 'px"><td colspan="' + (L.cols + 1) + '"></td></tr>';
    const frozenApart = sh.freezeRow && start > 0;
    if (frozenApart) parts.push(rowHTML(0, sh, valAt));
    const firstY = frozenApart ? L.rowY[1] : L.rowY[0];
    if (L.rowY[start] - firstY > 0) parts.push(spacer(L.rowY[start] - firstY));
    for (let r = Math.max(start, frozenApart ? 1 : 0); r <= end; r++) parts.push(rowHTML(r, sh, valAt));
    const lastBottom = L.rowY[L.rows - 1] + L.rowH[L.rows - 1];
    const endBottom = L.rowY[end] + L.rowH[end];
    if (lastBottom - endBottom > 0) parts.push(spacer(lastBottom - endBottom));
    const tbody = $('grid').tBodies[0];
    tbody.innerHTML = parts.join('');
    L.trMap = new Map();
    L.thRows = new Map();
    Array.from(tbody.rows).forEach(tr => {
      if (tr.dataset.r === undefined) return;
      L.trMap.set(+tr.dataset.r, tr);
      L.thRows.set(+tr.dataset.r, tr.cells[0]);
    });
    measureRows(Array.from(L.trMap.keys()));
    updateSelectionUI();
  }

  /** Corrige las alturas previstas con las reales (por ejemplo, letra grande). */
  function measureRows(list) {
    const L = S.L;
    let changed = false;
    list.forEach(r => {
      const tr = L.trMap.get(r);
      if (!tr) return;
      const h = tr.offsetHeight;
      if (h !== L.rowH[r]) { L.rowH[r] = h; changed = true; }
    });
    if (changed) {
      computeRowY();
      // Ajusta los espaciadores al nuevo total
      const sp = $('grid').tBodies[0].querySelectorAll('tr.spacer');
      const sh = sheet();
      const [start, end] = L.win;
      const frozenApart = sh.freezeRow && start > 0;
      const firstY = frozenApart ? L.rowY[1] : L.rowY[0];
      const topH = L.rowY[start] - firstY;
      const bottomH = L.rowY[L.rows - 1] + L.rowH[L.rows - 1] - (L.rowY[end] + L.rowH[end]);
      let i = 0;
      if (topH > 0 && sp[i]) sp[i++].style.height = topH + 'px';
      if (bottomH > 0 && sp[i]) sp[i].style.height = bottomH + 'px';
    }
  }

  /** Lo que obliga a redibujar toda la hoja si cambia (columnas, combinadas, inmovilizar, filtro, anchos). */
  function sheetShape() {
    const sh = sheet();
    const ext = usedExtent(sh);
    return [S.active, Math.max(26, ext.cols + 5, S.sel.c1 + 5), JSON.stringify(sh.merges), sh.freezeRow, sh.freezeCol,
      sh.filter ? JSON.stringify([sh.filter.r0, sh.filter.c0, sh.filter.r1, sh.filter.c1, Object.keys(sh.filter.crit)]) : '',
      JSON.stringify(sh.colW), Array.from(sh.hidden).join(',')].join('|');
  }

  function computeRowY() {
    const L = S.L;
    let y = L.headH;
    for (let r = 0; r < L.rows; r++) { L.rowY[r] = y; y += L.rowH[r]; }
  }

  /** Redibuja solo las filas indicadas (edición de celdas y formato). */
  function renderRows(rowSet) {
    const sh = sheet();
    const L = S.L;
    const tmp = document.createElement('tbody');
    const valAt = (r, c) => {
      const v = S.hf.getCellValue({ sheet: sid(sh), row: r, col: c });
      return v === undefined ? null : v;
    };
    const list = Array.from(rowSet).filter(r => L.trMap.has(r));
    list.forEach(r => {
      tmp.innerHTML = rowHTML(r, sh, valAt);
      const tr = tmp.firstChild;
      L.trMap.get(r).replaceWith(tr);
      L.trMap.set(r, tr);
      L.thRows.set(r, tr.cells[0]);
    });
    measureRows(list);
    updateSelectionUI();
    if (S.clip && S.clip.sheet === sh.name && !S.clip.done) showCopyBox();
  }

  function rectOf(r0, c0, r1, c1) {
    const L = S.L;
    r1 = Math.min(r1, L.rows - 1); c1 = Math.min(c1, L.cols - 1);
    return {
      left: L.colX[c0],
      top: L.rowY[r0],
      width: L.colX[c1] + L.colW[c1] - L.colX[c0],
      height: L.rowY[r1] + L.rowH[r1] - L.rowY[r0]
    };
  }

  function place(el, rc, inset) {
    inset = inset || 0;
    el.style.left = (rc.left - 1 + inset) + 'px';
    el.style.top = (rc.top - 1 + inset) + 'px';
    el.style.width = Math.max(0, rc.width + 1 - inset * 2) + 'px';
    el.style.height = Math.max(0, rc.height + 1 - inset * 2) + 'px';
  }

  function updateSelectionUI() {
    if (!S.loaded) return;
    const s = S.sel;
    const L = S.L;
    const selBox = $('selBox'), actBox = $('actBox'), fh = $('fillHandle');
    const multi = s.r0 !== s.r1 || s.c0 !== s.c1;
    const am = mergeAt(s.ar, s.ac);
    const isMergeOnly = am && am.r0 === s.r0 && am.c0 === s.c0 && am.r1 === s.r1 && am.c1 === s.c1;
    const rc = rectOf(s.r0, s.c0, s.r1, s.c1);
    if (multi && !isMergeOnly) { place(selBox, rc); selBox.style.display = 'block'; } else selBox.style.display = 'none';
    const ac = am ? rectOf(am.r0, am.c0, am.r1, am.c1) : rectOf(s.ar, s.ac, s.ar, s.ac);
    place(actBox, ac);
    actBox.style.display = 'block';
    fh.style.left = (rc.left + rc.width - 4) + 'px';
    fh.style.top = (rc.top + rc.height - 4) + 'px';
    fh.style.display = 'block';

    // Encabezados resaltados
    if (L.thCols) {
      L.thCols.forEach((th, c) => th.classList.toggle('hsel', c >= s.c0 && c <= s.c1));
      L.thRows.forEach((th, r) => th.classList.toggle('hsel', r >= s.r0 && r <= s.r1));
    }
    const gridEl = $('grid');
    gridEl.classList.toggle('sel-cols', isWholeCols() && !isWholeRows());
    gridEl.classList.toggle('sel-rows', isWholeRows() && !isWholeCols());
    $('nameBox').value = multi && !isMergeOnly ? cellName(s.r0, s.c0) + ':' + cellName(s.r1, s.c1) : cellName(s.ar, s.ac);
    if (!S.edit) $('fbar').value = editTextOf(s.ar, s.ac);
    updateToolbarState();
  }

  function updateToolbarState() {
    const st = getStyle(S.sel.ar, S.sel.ac) || {};
    $('btnBold').classList.toggle('on', !!st.b);
    $('btnItalic').classList.toggle('on', !!st.i);
    $('btnUnderline').classList.toggle('on', !!st.u);
    $('btnAlignL').classList.toggle('on', st.align === 'left');
    $('btnAlignC').classList.toggle('on', st.align === 'center');
    $('btnAlignR').classList.toggle('on', st.align === 'right');
    $('fontSize').value = String(st.size || 11);
    const fmt = N.normalizeFormat(st.fmt);
    let kind = 'otro';
    if (!fmt) kind = 'general';
    else if (fmt.includes('$')) kind = 'moneda';
    else if (fmt.includes('%')) kind = 'porcentaje';
    else if (N.isDateFormat(fmt)) kind = 'fecha';
    else if (/^#,##0(\.0+)?$|^0(\.0+)?$/.test(fmt)) kind = 'numero';
    $('numFormat').value = kind;
    $('btnMerge').classList.toggle('on', !!mergeAt(S.sel.ar, S.sel.ac));
    $('btnFilter').classList.toggle('on', !!sheet().filter);
    $('btnUndo').disabled = !S.undo.length;
    $('btnRedo').disabled = !S.redo.length;
  }

  function showCopyBox() {
    const c = S.clip;
    const box = $('copyBox');
    place(box, rectOf(c.r0, c.c0, c.r1, c.c1));
    box.hidden = false;
  }

  function clearClip() {
    if (S.clip) S.clip.done = true;
    $('copyBox').hidden = true;
  }

  function ensureVisible(r, c) {
    const sc = $('scroller');
    const L = S.L;
    if (r >= L.rows || c >= L.cols) return;
    const rc = rectOf(r, c, r, c);
    const sh = sheet();
    const topFixed = L.headH + (sh.freezeRow ? L.rowH[0] : 0);
    const leftFixed = headW() + (sh.freezeCol ? L.colW[0] : 0);
    if (!(sh.freezeRow && r === 0)) {
      if (rc.top - topFixed < sc.scrollTop) sc.scrollTop = rc.top - topFixed;
      else if (rc.top + rc.height > sc.scrollTop + sc.clientHeight) sc.scrollTop = rc.top + rc.height - sc.clientHeight;
    }
    if (!(sh.freezeCol && c === 0)) {
      if (rc.left - leftFixed < sc.scrollLeft) sc.scrollLeft = rc.left - leftFixed;
      else if (rc.left + rc.width > sc.scrollLeft + sc.clientWidth) sc.scrollLeft = rc.left + rc.width - sc.clientWidth;
    }
  }

  // ---------- Selección ----------
  function select(r0, c0, r1, c1, ar, ac, noScroll) {
    const rg = expandRange({ r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) });
    // Columna(s) o fila(s) enteras: se recuerda para que la cuadrícula no las trate como un rango común.
    const wholeCols = rg.r0 === 0 && rg.r1 >= S.L.rows - 1;
    const wholeRows = rg.c0 === 0 && rg.c1 >= S.L.cols - 1;
    S.sel = Object.assign(rg, { ar: ar === undefined ? r0 : ar, ac: ac === undefined ? c0 : ac, wholeCols, wholeRows });
    const grow = (!wholeCols && S.sel.r1 + 5 >= S.L.rows) || (!wholeRows && S.sel.c1 + 2 >= S.L.cols);
    if (grow) renderGrid(); else updateSelectionUI();
    if (!noScroll) ensureVisible(S.sel.ar, S.sel.ac);
  }

  function selectCell(r, c) {
    r = clamp(r, 0, 1048575); c = clamp(c, 0, 16383);
    S.anchor = { r, c };
    const m = mergeAt(r, c);
    if (m) select(m.r0, m.c0, m.r1, m.c1, m.r0, m.c0);
    else select(r, c, r, c, r, c);
  }

  function extendTo(r, c) {
    r = Math.max(0, r); c = Math.max(0, c);
    select(S.anchor.r, S.anchor.c, r, c, S.anchor.r, S.anchor.c);
    S.sel.ar = S.anchor.r; S.sel.ac = S.anchor.c;
    updateSelectionUI();
    ensureVisible(r, c);
  }

  function isEmptyCell(r, c) {
    const v = cellValue(r, c);
    return v === null || v === '';
  }

  function isRowHidden(r) { return sheet().hidden.has(r); }

  function stepRow(r, d) {
    let n = r + d;
    while (n >= 0 && isRowHidden(n)) n += d;
    return n < 0 ? r : n;
  }

  /** Ctrl+flecha: salta al borde del bloque de datos. */
  function jump(r, c, dr, dc) {
    const ext = usedExtent();
    const maxR = Math.max(ext.rows, 1) + 50, maxC = Math.max(ext.cols, 1) + 10;
    const inBounds = (rr, cc) => rr >= 0 && cc >= 0 && rr < maxR && cc < maxC;
    let nr = r + dr, nc = c + dc;
    if (!inBounds(nr, nc)) return [r, c];
    if (!isEmptyCell(r, c) && !isEmptyCell(nr, nc)) {
      while (inBounds(nr + dr, nc + dc) && !isEmptyCell(nr + dr, nc + dc)) { nr += dr; nc += dc; }
      return [nr, nc];
    }
    while (inBounds(nr, nc) && isEmptyCell(nr, nc)) { nr += dr; nc += dc; }
    if (!inBounds(nr, nc)) {
      if (dr < 0 || dc < 0) return [Math.max(0, nr - dr), Math.max(0, nc - dc)];
      return [dr ? Math.max(ext.rows - 1, r) : r, dc ? Math.max(ext.cols - 1, c) : c];
    }
    return [nr, nc];
  }

  function moveActive(dr, dc, extend, ctrl) {
    const s = S.sel;
    let r, c;
    if (extend) {
      // Se mueve el extremo opuesto al ancla
      r = s.r0 === S.anchor.r ? s.r1 : s.r0;
      c = s.c0 === S.anchor.c ? s.c1 : s.c0;
      if (ctrl) [r, c] = jump(r, c, dr, dc);
      else { r = dr ? stepRow(r, dr) : r; c = Math.max(0, c + dc); }
      extendTo(r, c);
      return;
    }
    r = s.ar; c = s.ac;
    const m = mergeAt(r, c);
    if (ctrl) [r, c] = jump(r, c, dr, dc);
    else {
      if (m) {
        if (dr > 0) r = m.r1; if (dc > 0) c = m.c1;
        if (dr < 0) r = m.r0; if (dc < 0) c = m.c0;
      }
      r = dr ? stepRow(r, dr) : r;
      c = Math.max(0, c + dc);
    }
    selectCell(r, c);
  }

  // ---------- Edición de celdas ----------
  function startEdit(mode, initial) {
    if (!S.loaded) return;
    const s = S.sel;
    const r = s.ar, c = s.ac;
    const m = mergeAt(r, c);
    const rc = m ? rectOf(m.r0, m.c0, m.r1, m.c1) : rectOf(r, c, r, c);
    ensureVisible(r, c);
    const ed = $('editor');
    const text = initial !== undefined ? initial : editTextOf(r, c);
    S.edit = { r, c, mode, original: editTextOf(r, c), refStart: -1 };
    ed.value = text;
    ed.style.left = (rc.left - 1) + 'px';
    ed.style.top = (rc.top - 1) + 'px';
    ed.style.minWidth = (rc.width + 2) + 'px';
    ed.style.height = (rc.height + 2) + 'px';
    const st = getStyle(r, c) || {};
    ed.style.fontWeight = st.b ? '700' : '';
    ed.style.fontStyle = st.i ? 'italic' : '';
    ed.style.fontSize = fontPx(st.size) + 'px';
    ed.hidden = false;
    autosizeEditor();
    $('fbar').value = text;
    ed.focus();
    const end = ed.value.length;
    ed.setSelectionRange(end, end);
  }

  function autosizeEditor() {
    const ed = $('editor');
    ed.style.width = '0px';
    ed.style.width = Math.max(parseFloat(ed.style.minWidth) || 0, ed.scrollWidth + 8) + 'px';
  }

  function cancelEdit() {
    if (!S.edit) return;
    S.edit = null;
    $('editor').hidden = true;
    $('fbar').value = editTextOf(S.sel.ar, S.sel.ac);
    focusGrid();
  }

  /** Confirma lo escrito. Devuelve false si la fórmula tiene un error. */
  function commitEdit() {
    if (!S.edit) return true;
    const { r, c, original } = S.edit;
    const text = $('editor').value;
    if (text === original) { cancelEdit(); return true; }
    const p = N.parseInput(text);
    if (p.formula && !S.hf.validateFormula(p.value)) {
      toast('La fórmula tiene un error. Revisá paréntesis y separadores (usá ";" entre argumentos y "," para decimales).', 'error');
      $('editor').focus();
      return false;
    }
    S.edit = null;
    $('editor').hidden = true;
    commit(() => writeInput(r, c, text), { partial: true });
    focusGrid();
    return true;
  }

  function focusGrid() {
    const k = $('kbd');
    k.value = ' ';
    k.focus({ preventScroll: true });
    k.select();
  }

  // Modo "señalar": al escribir una fórmula, hacer clic en una celda agrega su referencia.
  function canPoint() {
    if (!S.edit) return false;
    const ed = S.edit.source === 'bar' ? $('fbar') : $('editor');
    const v = ed.value;
    if (v[0] !== '=') return false;
    if (S.edit.refStart >= 0) return true;
    const pos = ed.selectionStart;
    const before = v.slice(0, pos).replace(/\s+$/, '');
    return /[=(;:+\-*/^&<>]$/.test(before);
  }

  function insertRef(r0, c0, r1, c1) {
    const ed = S.edit.source === 'bar' ? $('fbar') : $('editor');
    const ref = (r0 === r1 && c0 === c1) ? cellName(r0, c0) : cellName(Math.min(r0, r1), Math.min(c0, c1)) + ':' + cellName(Math.max(r0, r1), Math.max(c0, c1));
    let v = ed.value;
    if (S.edit.refStart < 0) {
      S.edit.refStart = ed.selectionStart;
      S.edit.refEnd = ed.selectionStart;
    }
    v = v.slice(0, S.edit.refStart) + ref + v.slice(S.edit.refEnd);
    S.edit.refEnd = S.edit.refStart + ref.length;
    ed.value = v;
    syncEditors(ed);
    ed.focus();
    ed.setSelectionRange(S.edit.refEnd, S.edit.refEnd);
  }

  function syncEditors(from) {
    if (from === $('editor')) $('fbar').value = from.value;
    else $('editor').value = from.value;
    autosizeEditor();
  }

  // ---------- Comandos sobre la selección ----------
  function forEachSel(fn) {
    const s = S.sel;
    for (let r = s.r0; r <= s.r1; r++) for (let c = s.c0; c <= s.c1; c++) fn(r, c);
  }

  function clearContents() {
    commit(() => {
      const s = S.sel;
      const rows = [];
      for (let r = s.r0; r <= s.r1; r++) rows.push(new Array(s.c1 - s.c0 + 1).fill(null));
      setContents(addr(s.r0, s.c0), rows);
    }, { partial: true });
  }

  function toggleStyle(prop) {
    const cur = getStyle(S.sel.ar, S.sel.ac) || {};
    const val = !cur[prop];
    commit(() => forEachSel((r, c) => setStyleProps(r, c, { [prop]: val })), { partial: true });
  }

  function setStyleAll(props) {
    commit(() => forEachSel((r, c) => setStyleProps(r, c, props)), { partial: true });
  }

  function setAlign(a) {
    const cur = getStyle(S.sel.ar, S.sel.ac) || {};
    setStyleAll({ align: cur.align === a ? null : a });
  }

  function setFormat(kind) {
    if (kind === 'otro') return;
    const fmt = kind === 'general' ? null : N.FORMATS[kind];
    setStyleAll({ fmt });
  }

  function changeDecimals(delta) {
    const st = getStyle(S.sel.ar, S.sel.ac) || {};
    const v = cellValue(S.sel.ar, S.sel.ac);
    const fmt = N.changeDecimals(st.fmt, delta, typeof v === 'number' ? v : 0);
    setStyleAll({ fmt });
  }

  function applyBorders(kind) {
    const s = S.sel;
    commit(() => forEachSel((r, c) => {
      if (kind === 'all') setStyleProps(r, c, { bt: true, br: true, bb: true, bl: true });
      else if (kind === 'none') {
        setStyleProps(r, c, { bt: null, br: null, bb: null, bl: null });
        if (r === s.r0 && r > 0) setStyleProps(r - 1, c, { bb: null });
        if (c === s.c0 && c > 0) setStyleProps(r, c - 1, { br: null });
      } else if (kind === 'outer') {
        const p = {};
        if (r === s.r0) p.bt = true;
        if (r === s.r1) p.bb = true;
        if (c === s.c0) p.bl = true;
        if (c === s.c1) p.br = true;
        setStyleProps(r, c, p);
      } else if (kind === 'bottom' && r === s.r1) setStyleProps(r, c, { bb: true });
      else if (kind === 'top' && r === s.r0) setStyleProps(r, c, { bt: true });
    }), { partial: true });
  }

  async function toggleMerge() {
    const s = S.sel;
    const sh = sheet();
    const inter = sh.merges.filter(m => !(m.r1 < s.r0 || m.r0 > s.r1 || m.c1 < s.c0 || m.c0 > s.c1));
    if (inter.length) {
      commit(() => { sh.merges = sh.merges.filter(m => !inter.includes(m)); });
      return;
    }
    if (s.r0 === s.r1 && s.c0 === s.c1) return toast('Seleccioná dos o más celdas para combinarlas.');
    let others = 0;
    forEachSel((r, c) => { if ((r !== s.r0 || c !== s.c0) && !isEmptyCell(r, c)) others++; });
    if (others) {
      const ok = await confirmModal('Combinar celdas', 'Al combinar se conserva solo el valor de la celda superior izquierda. Los demás valores se borran.', 'Combinar');
      if (!ok) return;
    }
    commit(() => {
      forEachSel((r, c) => { if (r !== s.r0 || c !== s.c0) setContents(addr(r, c), [[null]]); });
      sh.merges.push({ r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 });
      setStyleProps(s.r0, s.c0, { align: 'center' });
    });
    S.sel.ar = s.r0; S.sel.ac = s.c0;
    updateSelectionUI();
  }

  function setFreeze(kind) {
    const sh = sheet();
    commit(() => {
      sh.freezeRow = kind === 'row' || kind === 'both';
      sh.freezeCol = kind === 'col' || kind === 'both';
    });
  }

  // ---------- Insertar y eliminar filas / columnas ----------
  function shiftSheetMeta(sh, axis, at, count) {
    // count > 0: inserta; count < 0: elimina |count| desde "at".
    const isRow = axis === 'row';
    const mapIdx = i => {
      if (i < at) return i;
      if (count < 0 && i < at - count) return null;
      return i + count;
    };
    const styles = new Map();
    sh.styles.forEach((v, k) => {
      const [r, c] = k.split(',').map(Number);
      const n = mapIdx(isRow ? r : c);
      if (n === null) return;
      styles.set(isRow ? key(n, c) : key(r, n), v);
    });
    sh.styles = styles;
    const sizes = isRow ? sh.rowH : sh.colW;
    const newSizes = {};
    Object.keys(sizes).forEach(i => { const n = mapIdx(+i); if (n !== null) newSizes[n] = sizes[i]; });
    if (isRow) sh.rowH = newSizes; else sh.colW = newSizes;
    sh.merges = sh.merges.map(m => {
      let a = isRow ? m.r0 : m.c0, b = isRow ? m.r1 : m.c1;
      if (count > 0) {
        if (a >= at) { a += count; b += count; } else if (b >= at) b += count;
      } else {
        const del0 = at, del1 = at - count - 1;
        const na = a < del0 ? a : (a > del1 ? a + count : del0);
        const nb = b < del0 ? b : (b > del1 ? b + count : del0 - 1);
        if (nb < na) return null;
        a = na; b = nb;
      }
      const out = Object.assign({}, m);
      if (isRow) { out.r0 = a; out.r1 = b; } else { out.c0 = a; out.c1 = b; }
      return (out.r0 === out.r1 && out.c0 === out.c1) ? null : out;
    }).filter(Boolean);
    if (sh.filter) {
      const f = sh.filter;
      if (isRow) {
        const a = mapIdx(f.r0), b = mapIdx(f.r1);
        if (a === null) { sh.filter = null; sh.hidden = new Set(); }
        else { f.r0 = a; f.r1 = b === null ? Math.max(a, at - 1) : b; }
      } else {
        sh.filter = null;
        sh.hidden = new Set();
      }
    }
    if (isRow) {
      const hidden = new Set();
      sh.hidden.forEach(r => { const n = mapIdx(r); if (n !== null) hidden.add(n); });
      sh.hidden = hidden;
    }
  }

  function insertRows() {
    const s = S.sel, n = s.r1 - s.r0 + 1;
    commit(() => {
      S.hf.addRows(sid(), [s.r0, n]);
      shiftSheetMeta(sheet(), 'row', s.r0, n);
    });
  }
  function insertCols() {
    const s = S.sel, n = s.c1 - s.c0 + 1;
    commit(() => {
      S.hf.addColumns(sid(), [s.c0, n]);
      shiftSheetMeta(sheet(), 'col', s.c0, n);
    });
  }
  function deleteRows() {
    const s = S.sel, n = s.r1 - s.r0 + 1;
    commit(() => {
      S.hf.removeRows(sid(), [s.r0, n]);
      shiftSheetMeta(sheet(), 'row', s.r0, -n);
    });
    selectCell(s.r0, s.ac);
  }
  function deleteCols() {
    const s = S.sel, n = s.c1 - s.c0 + 1;
    commit(() => {
      S.hf.removeColumns(sid(), [s.c0, n]);
      shiftSheetMeta(sheet(), 'col', s.c0, -n);
    });
    selectCell(s.ar, s.c0);
  }

  // ---------- Mover columnas y filas arrastrando ----------
  /** Posición nueva de cada índice al mover `count` elementos desde `start` hasta antes de `boundary`. */
  function movedIndex(i, start, count, boundary) {
    const end = start + count;
    if (boundary > end) {
      if (i >= start && i < end) return i + (boundary - end);
      if (i >= end && i < boundary) return i - count;
    } else if (boundary < start) {
      if (i >= start && i < end) return i - (start - boundary);
      if (i >= boundary && i < start) return i + count;
    }
    return i;
  }

  /** Mueve columnas (axis 'col') o filas (axis 'row'); las fórmulas se ajustan solas. */
  function moveLines(axis, start, count, boundary) {
    if (boundary >= start && boundary <= start + count) return;
    const isRow = axis === 'row';
    const sh = sheet();
    const map = i => movedIndex(i, start, count, boundary);
    for (const m of sh.merges) {
      const a = isRow ? m.r0 : m.c0, b = isRow ? m.r1 : m.c1;
      for (let i = a; i < b; i++) {
        if (map(i + 1) !== map(i) + 1) {
          toast('No se puede mover: hay celdas combinadas que quedarían separadas.', 'error');
          return;
        }
      }
    }
    const first = map(start);
    commit(() => {
      if (isRow) S.hf.moveRows(sid(), start, count, boundary);
      else S.hf.moveColumns(sid(), start, count, boundary);
      const styles = new Map();
      sh.styles.forEach((v, k) => {
        const [r, c] = k.split(',').map(Number);
        styles.set(isRow ? key(map(r), c) : key(r, map(c)), v);
      });
      sh.styles = styles;
      const sizes = isRow ? sh.rowH : sh.colW;
      const moved = {};
      Object.keys(sizes).forEach(i => { moved[map(+i)] = sizes[i]; });
      if (isRow) sh.rowH = moved; else sh.colW = moved;
      sh.merges = sh.merges.map(m => (isRow
        ? { r0: map(m.r0), c0: m.c0, r1: map(m.r1), c1: m.c1 }
        : { r0: m.r0, c0: map(m.c0), r1: m.r1, c1: map(m.c1) }));
      const hidden = new Set();
      sh.hidden.forEach(r => hidden.add(isRow ? map(r) : r));
      sh.hidden = hidden;
      const f = sh.filter;
      if (f) {
        const lo = isRow ? f.r0 : f.c0, hi = isRow ? f.r1 : f.c1;
        const ids = [];
        for (let i = lo; i <= hi; i++) ids.push(map(i));
        const min = Math.min(...ids), max = Math.max(...ids);
        const contiguous = max - min === hi - lo;
        if (!contiguous || (isRow && map(lo) !== min)) {
          sh.filter = null;
          sh.hidden = new Set();
          toast('Se quitó el filtro porque el movimiento cambió la tabla filtrada.');
        } else if (isRow) {
          f.r0 = min; f.r1 = max;
        } else {
          f.c0 = min; f.c1 = max;
          const crit = {};
          Object.keys(f.crit).forEach(c => { crit[map(+c)] = f.crit[c]; });
          f.crit = crit;
        }
      }
      S.clip = null;
      $('copyBox').hidden = true;
    });
    if (isRow) { S.anchor = { r: first, c: 0 }; select(first, 0, first + count - 1, S.L.cols - 1, first, 0); }
    else { S.anchor = { r: 0, c: first }; select(0, first, S.L.rows - 1, first + count - 1, 0, first); }
  }

  const isWholeCols = () => !!S.sel.wholeCols;
  const isWholeRows = () => !!S.sel.wholeRows;

  /** Borde de columna/fila más cercano al puntero: índice antes del cual se insertaría. */
  function boundaryAt(axis, x, y) {
    const L = S.L;
    const { px, py } = gridPoint(x, y);
    if (axis === 'col') {
      let c = 0;
      while (c < L.cols - 1 && L.colX[c] + L.colW[c] <= px) c++;
      return px < L.colX[c] + L.colW[c] / 2 ? c : c + 1;
    }
    let r = 0;
    while (r < L.rows - 1 && L.rowY[r] + L.rowH[r] <= py) r++;
    return py < L.rowY[r] + L.rowH[r] / 2 ? r : r + 1;
  }

  function showMoveLine(axis, boundary) {
    const L = S.L;
    const line = $('moveLine');
    const totalH = L.rowY[L.rows - 1] + L.rowH[L.rows - 1];
    const totalW = L.colX[L.cols - 1] + L.colW[L.cols - 1];
    if (axis === 'col') {
      const x = boundary >= L.cols ? totalW : L.colX[boundary];
      line.style.cssText = 'left:' + (x - 2) + 'px;top:0;width:4px;height:' + totalH + 'px';
    } else {
      const y = boundary >= L.rows ? totalH : L.rowY[boundary];
      line.style.cssText = 'top:' + (y - 2) + 'px;left:0;height:4px;width:' + totalW + 'px';
    }
    line.hidden = false;
  }

  function endMoveDrag() {
    $('moveLine').hidden = true;
    document.body.classList.remove('moving');
  }

  // ---------- Copiar, cortar y pegar ----------
  function selectionTSV() {
    const s = S.sel;
    const lines = [];
    for (let r = s.r0; r <= s.r1; r++) {
      if (isRowHidden(r)) continue;
      const row = [];
      for (let c = s.c0; c <= s.c1; c++) {
        const t = displayText(r, c);
        row.push(/[\t\n"]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t);
      }
      lines.push(row.join('\t'));
    }
    return lines.join('\r\n');
  }

  function doCopy(e, cut) {
    const s = S.sel;
    const cells = [];
    for (let r = s.r0; r <= s.r1; r++) {
      const row = [];
      for (let c = s.c0; c <= s.c1; c++) row.push({ raw: cellRaw(r, c), style: getStyle(r, c) ? Object.assign({}, getStyle(r, c)) : null });
      cells.push(row);
    }
    const merges = sheet().merges.filter(m => m.r0 >= s.r0 && m.r1 <= s.r1 && m.c0 >= s.c0 && m.c1 <= s.c1).map(m => ({ r0: m.r0 - s.r0, c0: m.c0 - s.c0, r1: m.r1 - s.r0, c1: m.c1 - s.c0 }));
    const text = selectionTSV();
    S.clip = { sheet: sheet().name, r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1, cells, merges, cut: !!cut, text };
    if (e && e.clipboardData) {
      e.clipboardData.setData('text/plain', text);
      e.preventDefault();
    }
    showCopyBox();
  }

  const normText = t => (t || '').replace(/\r\n/g, '\n').replace(/\n$/, '');

  function doPaste(text) {
    const clip = S.clip;
    if (clip && !clip.done && normText(text) === normText(clip.text)) return pasteInternal(clip);
    pasteExternal(text);
  }

  function pasteInternal(clip) {
    const s = S.sel;
    const h = clip.cells.length, w = clip.cells[0].length;
    // Si se seleccionó un rango mayor que es múltiplo del copiado, se repite (como Excel).
    const selH = s.r1 - s.r0 + 1, selW = s.c1 - s.c0 + 1;
    const reps = { r: selH % h === 0 && selH > h ? selH / h : 1, c: selW % w === 0 && selW > w ? selW / w : 1 };
    const srcSheet = S.sheets.find(x => x.name === clip.sheet);
    commit(() => {
      if (clip.cut && srcSheet) {
        // Cortar: se vacía el origen y las fórmulas se mueven sin ajustar referencias.
        for (let r = clip.r0; r <= clip.r1; r++) for (let c = clip.c0; c <= clip.c1; c++) {
          setContents(addr(r, c, srcSheet), [[null]]);
          srcSheet.styles.delete(key(r, c));
          if (srcSheet === sheet()) { S.touched.add(r); S.touched.add(r - 1); }
        }
        srcSheet.merges = srcSheet.merges.filter(m => !(m.r0 >= clip.r0 && m.r1 <= clip.r1 && m.c0 >= clip.c0 && m.c1 <= clip.c1));
      }
      const sh = sheet();
      for (let rr = 0; rr < reps.r; rr++) for (let cc = 0; cc < reps.c; cc++) {
        const br = s.r0 + rr * h, bc = s.c0 + cc * w;
        const data = clip.cells.map((row, i) => row.map((cell, j) => {
          const raw = cell.raw;
          if (isFormula(raw) && !clip.cut) return '=' + F.shift(raw.slice(1), br - clip.r0, bc - clip.c0);
          return raw === undefined ? null : raw;
        }));
        sh.merges = sh.merges.filter(m => m.r1 < br || m.r0 > br + h - 1 || m.c1 < bc || m.c0 > bc + w - 1);
        setContents(addr(br, bc), data);
        clip.cells.forEach((row, i) => row.forEach((cell, j) => {
          if (cell.style) sh.styles.set(key(br + i, bc + j), Object.assign({}, cell.style));
          else sh.styles.delete(key(br + i, bc + j));
          S.touched.add(br + i); S.touched.add(br + i - 1);
        }));
        clip.merges.forEach(m => sh.merges.push({ r0: br + m.r0, c0: bc + m.c0, r1: br + m.r1, c1: bc + m.c1 }));
      }
    }, { partial: !clip.cut || clip.sheet === sheet().name });
    select(s.r0, s.c0, s.r0 + h * reps.r - 1, s.c0 + w * reps.c - 1, s.r0, s.c0);
    if (clip.cut) { S.clip = null; $('copyBox').hidden = true; }
  }

  function pasteExternal(text) {
    if (!text) return;
    const rows = IO.parseDelimited(text.replace(/\r?\n$/, ''), '\t');
    if (!rows.length) return;
    const s = S.sel;
    commit(() => {
      rows.forEach((row, i) => row.forEach((t, j) => writeInput(s.r0 + i, s.c0 + j, t)));
    }, { partial: true });
    const w = Math.max(...rows.map(r => r.length));
    select(s.r0, s.c0, s.r0 + rows.length - 1, s.c0 + w - 1, s.r0, s.c0);
  }

  async function pasteFromButton() {
    try {
      const text = await navigator.clipboard.readText();
      doPaste(text);
    } catch (e) {
      toast('El navegador no permite pegar desde el botón. Usá Ctrl+V.');
    }
  }

  function copyFromButton(cut) {
    focusGrid();
    S.pendingCut = cut;
    const ok = document.execCommand && document.execCommand(cut ? 'cut' : 'copy');
    if (!ok) doCopy(null, cut);
    S.pendingCut = false;
  }

  // ---------- Autocompletar arrastrando ----------
  const DAYS = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo'];
  const MONTHS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

  function matchCase(sample, word) {
    if (sample === sample.toUpperCase()) return word.toUpperCase();
    if (sample[0] === sample[0].toUpperCase()) return word[0].toUpperCase() + word.slice(1);
    return word;
  }

  /** Genera los valores de una línea (fila o columna) a partir de los de origen. */
  function seriesFor(src, count, sh) {
    // src: [{ raw, value, style }], devuelve [{ raw, style, dr }] para las posiciones 1..count
    const n = src.length;
    const nums = src.every(x => typeof x.raw === 'number');
    const out = [];
    const listOf = v => {
      if (typeof v !== 'string') return null;
      const low = v.replace(/^'/, '').toLowerCase();
      if (DAYS.includes(low)) return DAYS;
      if (MONTHS.includes(low)) return MONTHS;
      return null;
    };
    const textNum = x => (typeof x.raw === 'string' && !isFormula(x.raw) ? /^(.*?)(\d+)$/.exec(x.raw.replace(/^'/, '')) : null);
    const lists = src.map(x => listOf(x.raw));
    const tnums = src.map(textNum);
    for (let k = 1; k <= count; k++) {
      const i = (k - 1) % n;
      const base = src[i];
      let raw = base.raw;
      if (nums && n >= 2) {
        const step = (src[n - 1].raw - src[0].raw) / (n - 1);
        raw = Number((src[n - 1].raw + step * k).toPrecision(15));
      } else if (nums && n === 1 && base.style && N.isDateFormat(base.style.fmt)) {
        raw = base.raw + k;
      } else if (lists.every(l => l && l === lists[0])) {
        const list = lists[0];
        const last = list.indexOf(src[n - 1].raw.replace(/^'/, '').toLowerCase());
        const step = n >= 2 ? (list.indexOf(src[1].raw.replace(/^'/, '').toLowerCase()) - list.indexOf(src[0].raw.replace(/^'/, '').toLowerCase()) + list.length) % list.length || 1 : 1;
        raw = "'" + matchCase(src[n - 1].raw.replace(/^'/, ''), list[(last + step * k) % list.length]);
      } else if (tnums.every(t => t) && tnums.every(t => t[1] === tnums[0][1])) {
        const step = n >= 2 ? (+tnums[n - 1][2] - +tnums[0][2]) / (n - 1) : 1;
        const v = +tnums[n - 1][2] + step * k;
        raw = "'" + tnums[0][1] + (v >= 0 ? String(v).padStart(tnums[n - 1][2].length, '0') : v);
      } else if (isFormula(raw)) {
        raw = { formula: raw, offset: n * Math.ceil(k / n) };
      }
      out.push({ raw, style: base.style, srcIndex: i });
    }
    return out;
  }

  function doFill(dir, count) {
    if (count <= 0) return;
    const s = S.sel;
    const sh = sheet();
    commit(() => {
      const vertical = dir === 'down' || dir === 'up';
      const lines = vertical ? [s.c0, s.c1] : [s.r0, s.r1];
      for (let L = lines[0]; L <= lines[1]; L++) {
        const src = [];
        const len = vertical ? s.r1 - s.r0 + 1 : s.c1 - s.c0 + 1;
        for (let i = 0; i < len; i++) {
          const r = vertical ? s.r0 + i : L, c = vertical ? L : s.c0 + i;
          src.push({ raw: cellRaw(r, c), style: getStyle(r, c) || null });
        }
        const ordered = (dir === 'up' || dir === 'left') ? src.slice().reverse() : src;
        const gen = seriesFor(ordered, count, sh);
        gen.forEach((g, idx) => {
          const k = idx + 1;
          let r, c;
          if (dir === 'down') { r = s.r1 + k; c = L; }
          else if (dir === 'up') { r = s.r0 - k; c = L; }
          else if (dir === 'right') { r = L; c = s.c1 + k; }
          else { r = L; c = s.c0 - k; }
          if (r < 0 || c < 0) return;
          let raw = g.raw;
          if (raw && typeof raw === 'object') {
            const sign = (dir === 'up' || dir === 'left') ? -1 : 1;
            const d = sign * raw.offset;
            raw = '=' + F.shift(raw.formula.slice(1), vertical ? d : 0, vertical ? 0 : d);
          }
          setContents(addr(r, c), [[raw === undefined ? null : raw]]);
          if (g.style) sh.styles.set(key(r, c), Object.assign({}, g.style)); else sh.styles.delete(key(r, c));
          S.touched.add(r);
        });
      }
    }, { partial: true });
    if (dir === 'down') select(s.r0, s.c0, s.r1 + count, s.c1, s.ar, s.ac);
    else if (dir === 'up') select(s.r0 - count, s.c0, s.r1, s.c1, s.ar, s.ac);
    else if (dir === 'right') select(s.r0, s.c0, s.r1, s.c1 + count, s.ar, s.ac);
    else select(s.r0, s.c0 - count, s.r1, s.c1, s.ar, s.ac);
  }

  // ---------- Ordenar y filtrar ----------
  /** Bloque de datos contiguo alrededor de una celda (como Ctrl+A en Excel). */
  function currentRegion(r, c) {
    let r0 = r, r1 = r, c0 = c, c1 = c;
    const filled = (rr, cc) => rr >= 0 && cc >= 0 && !isEmptyCell(rr, cc);
    let grew = true;
    while (grew) {
      grew = false;
      const rowHas = rr => { for (let cc = c0 - 1; cc <= c1 + 1; cc++) if (filled(rr, cc)) return true; return false; };
      const colHas = cc => { for (let rr = r0 - 1; rr <= r1 + 1; rr++) if (filled(rr, cc)) return true; return false; };
      if (r0 > 0 && rowHas(r0 - 1)) { r0--; grew = true; }
      if (rowHas(r1 + 1)) { r1++; grew = true; }
      if (c0 > 0 && colHas(c0 - 1)) { c0--; grew = true; }
      if (colHas(c1 + 1)) { c1++; grew = true; }
    }
    return { r0, c0, r1, c1 };
  }

  function looksLikeHeader(rg) {
    if (rg.r1 <= rg.r0) return false;
    if (sheet().freezeRow && rg.r0 === 0) return true;
    let headerText = true, bodyHasNumber = false;
    for (let c = rg.c0; c <= rg.c1; c++) {
      const v = cellValue(rg.r0, c);
      if (v !== null && typeof v !== 'string') headerText = false;
      const b = cellValue(rg.r0 + 1, c);
      if (typeof b === 'number') bodyHasNumber = true;
    }
    const st0 = getStyle(rg.r0, rg.c0) || {};
    return headerText && (bodyHasNumber || !!st0.b);
  }

  function compareValues(a, b) {
    const rank = v => (v === null || v === '' ? 4 : isError(v) ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
    const ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra - rb;
    if (ra === 0) return a - b;
    if (ra === 1) return a.localeCompare(b, 'es', { sensitivity: 'base', numeric: true });
    if (ra === 2) return (a ? 1 : 0) - (b ? 1 : 0);
    return 0;
  }

  function sortRange(rg, col, desc, hasHeader) {
    const sh = sheet();
    const start = hasHeader ? rg.r0 + 1 : rg.r0;
    if (start >= rg.r1 + 1) return;
    if (sh.merges.some(m => !(m.r1 < start || m.r0 > rg.r1 || m.c1 < rg.c0 || m.c0 > rg.c1))) {
      toast('No se puede ordenar un rango que tiene celdas combinadas.', 'error');
      return;
    }
    const rows = [];
    for (let r = start; r <= rg.r1; r++) {
      const cells = [];
      for (let c = rg.c0; c <= rg.c1; c++) cells.push({ raw: cellRaw(r, c), style: getStyle(r, c) || null });
      rows.push({ r, key: cellValue(r, col), cells });
    }
    rows.sort((a, b) => {
      const ea = a.key === null || a.key === '', eb = b.key === null || b.key === '';
      if (ea || eb) return ea === eb ? a.r - b.r : (ea ? 1 : -1);
      const cmp = compareValues(a.key, b.key);
      return (desc ? -cmp : cmp) || a.r - b.r;
    });
    commit(() => {
      const data = rows.map((row, i) => row.cells.map(cell => {
        const raw = cell.raw;
        if (isFormula(raw)) return '=' + F.shift(raw.slice(1), start + i - row.r, 0);
        return raw === undefined ? null : raw;
      }));
      setContents(addr(start, rg.c0), data);
      rows.forEach((row, i) => row.cells.forEach((cell, j) => {
        const k = key(start + i, rg.c0 + j);
        if (cell.style) sh.styles.set(k, cell.style); else sh.styles.delete(k);
        S.touched.add(start + i); S.touched.add(start + i - 1);
      }));
    }, { partial: true });
  }

  function sortCommand(desc) {
    const s = S.sel;
    let rg, header;
    if (s.r0 === s.r1 && s.c0 === s.c1) {
      if (isEmptyCell(s.ar, s.ac)) return toast('Ubicate en una celda con datos para ordenar.');
      rg = currentRegion(s.ar, s.ac);
      header = looksLikeHeader(rg);
    } else {
      rg = { r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 };
      header = looksLikeHeader(rg);
    }
    const f = sheet().filter;
    if (f && s.ar >= f.r0 && s.ar <= f.r1 && s.ac >= f.c0 && s.ac <= f.c1) { rg = { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1 }; header = true; }
    sortRange(rg, s.ac, desc, header);
    toast('Ordenado por la columna ' + F.numToCol(s.ac) + (header ? ' (la primera fila se tomó como encabezado).' : '.'));
  }

  function toggleFilter() {
    const sh = sheet();
    if (sh.filter) {
      commit(() => { sh.filter = null; sh.hidden = new Set(); });
      return;
    }
    const s = S.sel;
    let rg = (s.r0 === s.r1 && s.c0 === s.c1) ? currentRegion(s.ar, s.ac) : { r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 };
    if (rg.r0 === rg.r1 && isEmptyCell(rg.r0, rg.c0)) return toast('Ubicate en la tabla de datos para activar el filtro.');
    commit(() => { sh.filter = { r0: rg.r0, c0: rg.c0, r1: rg.r1, c1: rg.c1, crit: {} }; });
    toast('Filtro activado. Usá las flechas de la fila de encabezados.');
  }

  function applyFilter(sh) {
    sh = sh || sheet();
    const f = sh.filter;
    sh.hidden = new Set();
    if (!f) return;
    const cols = Object.keys(f.crit).map(Number);
    for (let r = f.r0 + 1; r <= f.r1; r++) {
      for (const c of cols) {
        if (!f.crit[c].has(displayText(r, c, sh))) { sh.hidden.add(r); break; }
      }
    }
  }

  function filterExtent(f) {
    // Incluye filas nuevas agregadas al final de la tabla.
    let r1 = f.r1;
    const ext = usedExtent();
    while (r1 + 1 < ext.rows) {
      let any = false;
      for (let c = f.c0; c <= f.c1; c++) if (!isEmptyCell(r1 + 1, c)) { any = true; break; }
      if (!any) break;
      r1++;
    }
    return r1;
  }

  function openFilterMenu(col, anchorEl) {
    const sh = sheet();
    const f = sh.filter;
    f.r1 = filterExtent(f);
    const values = new Map();
    for (let r = f.r0 + 1; r <= f.r1; r++) {
      const t = displayText(r, col);
      if (!values.has(t)) values.set(t, cellValue(r, col));
    }
    const sorted = Array.from(values.keys()).sort((a, b) => {
      if (a === '') return 1; if (b === '') return -1;
      return compareValues(values.get(a), values.get(b));
    });
    const current = f.crit[col];
    const menu = $('filterMenu');
    menu.innerHTML = '';
    const mk = (tag, props) => Object.assign(document.createElement(tag), props || {});
    const bAsc = mk('button', { textContent: 'Ordenar de A a Z' });
    const bDesc = mk('button', { textContent: 'Ordenar de Z a A' });
    bAsc.onclick = () => { closeFilterMenu(); sortRange({ r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1 }, col, false, true); };
    bDesc.onclick = () => { closeFilterMenu(); sortRange({ r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1 }, col, true, true); };
    const search = mk('input', { type: 'search', placeholder: 'Buscar', id: 'filterSearch' });
    const list = mk('div', { className: 'filter-list' });
    const allLbl = mk('label');
    const allCb = mk('input', { type: 'checkbox', checked: !current || current.size === sorted.length });
    allLbl.append(allCb, ' (Seleccionar todo)');
    list.append(allLbl);
    const boxes = sorted.map(t => {
      const lbl = mk('label', { title: t || '(Vacías)' });
      const cb = mk('input', { type: 'checkbox', checked: !current || current.has(t) });
      cb.dataset.val = t;
      lbl.append(cb, ' ' + (t === '' ? '(Vacías)' : t));
      list.append(lbl);
      return { lbl, cb, t };
    });
    allCb.onchange = () => boxes.forEach(b => { if (!b.lbl.hidden) b.cb.checked = allCb.checked; });
    search.oninput = () => {
      const q = search.value.toLowerCase();
      boxes.forEach(b => { b.lbl.hidden = q && !b.t.toLowerCase().includes(q); b.cb.checked = !b.lbl.hidden; });
    };
    const actions = mk('div', { className: 'filter-actions' });
    const bOk = mk('button', { className: 'primary', textContent: 'Aceptar' });
    const bCancel = mk('button', { textContent: 'Cancelar' });
    const bClear = mk('button', { textContent: 'Quitar filtro de la columna' });
    bOk.onclick = () => {
      const chosen = new Set(boxes.filter(b => b.cb.checked && !b.lbl.hidden).map(b => b.t));
      closeFilterMenu();
      commit(() => {
        if (chosen.size === sorted.length) delete f.crit[col]; else f.crit[col] = chosen;
        applyFilter();
      });
    };
    bCancel.onclick = closeFilterMenu;
    bClear.onclick = () => { closeFilterMenu(); commit(() => { delete f.crit[col]; applyFilter(); }); };
    actions.append(bCancel, bOk);
    menu.append(bAsc, bDesc, mk('hr'), search, list, actions);
    if (current) menu.append(bClear);
    const rc = anchorEl.getBoundingClientRect();
    menu.hidden = false;
    menu.style.left = Math.min(rc.left, window.innerWidth - 262) + 'px';
    menu.style.top = Math.min(rc.bottom + 2, window.innerHeight - menu.offsetHeight - 8) + 'px';
    search.focus();
  }

  function closeFilterMenu() {
    $('filterMenu').hidden = true;
    if (!S.edit) focusGrid();
  }

  // ---------- Buscar y reemplazar ----------
  function searchTextOf(r, c) {
    const raw = cellRaw(r, c);
    if (isFormula(raw)) return raw;
    return displayText(r, c);
  }

  function makeMatcher() {
    const q = $('findText').value;
    if (!q) return null;
    const cs = $('findCase').checked, whole = $('findWhole').checked;
    const norm = t => (cs ? t : t.toLowerCase());
    const nq = norm(q);
    return t => (whole ? norm(t) === nq : norm(t).includes(nq));
  }

  function findNext(silent) {
    const match = makeMatcher();
    if (!match) { $('findMsg').textContent = 'Escribí qué buscar.'; return false; }
    const ext = usedExtent();
    const total = ext.rows * ext.cols;
    let r = S.sel.ar, c = S.sel.ac;
    for (let i = 0; i < total; i++) {
      c++;
      if (c >= ext.cols) { c = 0; r++; }
      if (r >= ext.rows) r = 0;
      if (S.L.covered.has(key(r, c)) || isRowHidden(r)) continue;
      const t = searchTextOf(r, c);
      if (t && match(t)) {
        selectCell(r, c);
        $('findMsg').textContent = 'Encontrado en ' + cellName(r, c) + '.';
        return true;
      }
    }
    if (!silent) $('findMsg').textContent = 'No se encontró "' + $('findText').value + '" en esta hoja.';
    return false;
  }

  function replaceIn(text) {
    const q = $('findText').value, rep = $('replaceText').value;
    if ($('findWhole').checked) return rep;
    const flags = $('findCase').checked ? 'g' : 'gi';
    return text.replace(new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags), () => rep);
  }

  function writeReplaced(r, c) {
    const raw = cellRaw(r, c);
    const text = isFormula(raw) ? raw : editTextOf(r, c);
    const out = replaceIn(text);
    if (isFormula(out) && !S.hf.validateFormula(out)) return false;
    writeInput(r, c, out);
    return true;
  }

  function replaceOne() {
    const match = makeMatcher();
    if (!match) return;
    const { ar, ac } = S.sel;
    const t = searchTextOf(ar, ac);
    if (t && match(t)) {
      commit(() => { if (!writeReplaced(ar, ac)) return false; }, { partial: true });
    }
    findNext();
  }

  function replaceAll() {
    const match = makeMatcher();
    if (!match) { $('findMsg').textContent = 'Escribí qué buscar.'; return; }
    const ext = usedExtent();
    let n = 0;
    commit(() => {
      for (let r = 0; r < ext.rows; r++) for (let c = 0; c < ext.cols; c++) {
        const t = searchTextOf(r, c);
        if (t && match(t) && writeReplaced(r, c)) n++;
      }
      if (!n) return false;
    }, { partial: true });
    $('findMsg').textContent = n ? 'Se hicieron ' + n + ' reemplazos.' : 'No se encontró "' + $('findText').value + '" en esta hoja.';
  }

  function openFind() {
    $('findPanel').hidden = false;
    $('findText').focus();
    $('findText').select();
    $('findMsg').textContent = '';
  }

  // ---------- Zoom ----------
  function setZoom(z) {
    z = clamp(z, ZOOM_STEPS[0], ZOOM_STEPS[ZOOM_STEPS.length - 1]);
    if (z === S.zoom) return;
    if (S.edit && !commitEdit()) return;
    const sc = $('scroller');
    const ratio = z / S.zoom;
    const top = sc.scrollTop * ratio, left = sc.scrollLeft * ratio;
    S.zoom = z;
    updateZoomUI();
    if (!S.loaded) return;
    renderGrid();
    sc.scrollTop = top; sc.scrollLeft = left;
    renderWindow(true);
    updateSelectionUI();
  }

  function stepZoom(dir) {
    const i = ZOOM_STEPS.findIndex(v => Math.abs(v - S.zoom) < 0.001);
    setZoom(ZOOM_STEPS[clamp(i + dir, 0, ZOOM_STEPS.length - 1)]);
  }

  function updateZoomUI() {
    $('zoomLbl').textContent = Math.round(S.zoom * 100) + '%';
    $('zoomOut').disabled = !S.loaded || S.zoom <= ZOOM_STEPS[0];
    $('zoomIn').disabled = !S.loaded || S.zoom >= ZOOM_STEPS[ZOOM_STEPS.length - 1];
    $('zoomLbl').disabled = !S.loaded;
  }

  // ---------- Hojas ----------
  function renderTabs() {
    const tabs = $('tabs');
    tabs.innerHTML = '';
    S.sheets.forEach((sh, i) => {
      const b = document.createElement('button');
      b.className = 'tab' + (i === S.active ? ' active' : '');
      b.textContent = sh.name;
      b.dataset.idx = i;
      b.title = 'Doble clic para cambiar el nombre';
      tabs.append(b);
    });
  }

  function switchSheet(i) {
    if (S.edit && !commitEdit()) return;
    if (i === S.active) return;
    S.active = i;
    S.sel = { r0: 0, c0: 0, r1: 0, c1: 0, ar: 0, ac: 0 };
    S.anchor = { r: 0, c: 0 };
    renderTabs();
    renderGrid();
    $('scroller').scrollTo(0, 0);
  }

  function validSheetName(name, exceptIdx) {
    if (!name || !name.trim()) return 'El nombre no puede estar vacío.';
    if (name.length > 31) return 'El nombre puede tener como máximo 31 caracteres.';
    if (/[\[\]:*?\/\\]/.test(name)) return 'El nombre no puede tener estos caracteres: [ ] : * ? / \\';
    if (S.sheets.some((s, i) => i !== exceptIdx && s.name.toLowerCase() === name.toLowerCase())) return 'Ya existe una hoja con ese nombre.';
    return null;
  }

  function addSheet() {
    let n = S.sheets.length + 1;
    while (S.sheets.some(s => s.name.toLowerCase() === ('hoja' + n))) n++;
    const name = 'Hoja' + n;
    commit(() => {
      S.hf.addSheet(name);
      S.sheets.push({ name, styles: new Map(), merges: [], colW: {}, rowH: {}, freezeRow: false, freezeCol: false, filter: null, hidden: new Set() });
    }, { noRender: true });
    switchSheet(S.sheets.length - 1);
  }

  function renameSheet(i) {
    const tab = $('tabs').children[i];
    if (!tab) return;
    const sh = S.sheets[i];
    const input = document.createElement('input');
    input.value = sh.name;
    input.id = 'renameInput';
    tab.textContent = '';
    tab.append(input);
    input.focus();
    input.select();
    let done = false;
    const finish = ok => {
      if (done) return;
      done = true;
      const name = input.value.trim();
      if (ok && name !== sh.name) {
        const err = validSheetName(name, i);
        if (err) { toast(err, 'error'); renderTabs(); return; }
        commit(() => { S.hf.renameSheet(S.hf.getSheetId(sh.name), name); sh.name = name; }, { tabs: true });
      }
      renderTabs();
      focusGrid();
    };
    input.onkeydown = e => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    };
    input.onblur = () => finish(true);
    input.onmousedown = e => e.stopPropagation();
  }

  async function deleteSheet(i) {
    if (S.sheets.length === 1) return toast('El libro tiene que tener al menos una hoja.', 'error');
    const sh = S.sheets[i];
    const ok = await confirmModal('Eliminar hoja', 'Se va a eliminar la hoja "' + sh.name + '" con todo su contenido. Las fórmulas de otras hojas que la usen mostrarán #¡REF!.', 'Eliminar');
    if (!ok) return;
    commit(() => {
      S.hf.removeSheet(S.hf.getSheetId(sh.name));
      S.sheets.splice(i, 1);
      if (S.active >= S.sheets.length) S.active = S.sheets.length - 1;
      else if (S.active > i) S.active--;
    }, { tabs: true });
    S.sel = { r0: 0, c0: 0, r1: 0, c1: 0, ar: 0, ac: 0 };
    renderTabs();
    renderGrid();
  }

  // ---------- Abrir ----------
  async function openFile(file) {
    if (!file) return;
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    if (!['xlsx', 'xls', 'xlsm', 'csv', 'tsv', 'ods'].includes(ext)) {
      toast('Ese tipo de archivo no se puede abrir. Usá .xlsx, .xls, .xlsm, .csv, .tsv u .ods.', 'error');
      return;
    }
    if (S.dirty) {
      const ok = await confirmModal('Cambios sin guardar', 'El archivo actual tiene cambios sin guardar. Si abrís otro, esos cambios se pierden.', 'Abrir igual');
      if (!ok) return;
    }
    const t = toast('Abriendo ' + file.name + '…', null, 0);
    try {
      const book = await IO.readFile(file);
      if (!book.sheets.length) throw new Error('El archivo no tiene hojas.');
      loadBook(book, file.name);
      t.remove();
      if (ext === 'xlsm') toast('El archivo tiene macros: se abrió el contenido, pero las macros no se ejecutan ni se guardan.');
    } catch (e) {
      console.error(e);
      t.remove();
      const msg = /password|encrypt|cfb/i.test(String(e && e.message)) ? 'El archivo está protegido con contraseña o dañado, y no se puede abrir.' : 'No se pudo abrir el archivo. Puede estar dañado o tener un formato no compatible.';
      toast(msg, 'error');
    }
  }

  function loadBook(book, fileName) {
    const data = book.sheets.map(sh => {
      const d = [];
      sh.cells.forEach(([r, c, v]) => {
        while (d.length <= r) d.push([]);
        d[r][c] = v;
      });
      for (let r = 0; r < d.length; r++) {
        const row = d[r];
        for (let c = 0; c < row.length; c++) if (row[c] === undefined) row[c] = null;
      }
      return { name: sh.name, data: d };
    });
    buildHF(data);
    S.sheets = book.sheets.map(sh => ({
      name: sh.name, styles: sh.styles, merges: sh.merges, colW: sh.colW, rowH: sh.rowH,
      freezeRow: sh.freezeRow, freezeCol: sh.freezeCol, filter: null, hidden: new Set()
    }));
    S.active = 0;
    S.fileName = fileName;
    S.undo = []; S.redo = []; S.clip = null; S.edit = null;
    S.sel = { r0: 0, c0: 0, r1: 0, c1: 0, ar: 0, ac: 0 };
    S.anchor = { r: 0, c: 0 };
    S.loaded = true;
    setDirty(false);
    $('fileName').textContent = fileName;
    document.title = fileName + ' · Coftech Excel';
    $('welcome').hidden = true;
    $('scroller').hidden = false;
    document.querySelectorAll('.needs-file').forEach(b => { b.disabled = false; });
    updateZoomUI();
    renderTabs();
    renderGrid();
    $('scroller').scrollTo(0, 0);
    focusGrid();
  }

  // ---------- Guardar ----------
  function exportSource(onlyActive) {
    const list = onlyActive ? [sheet()] : S.sheets;
    return {
      sheets: list.map(sh => {
        const ext = usedExtent(sh);
        return {
          name: sh.name,
          rows: ext.rows,
          cols: ext.cols,
          cell: (r, c) => {
            const raw = cellRaw(r, c, sh);
            if (raw === null || raw === '') return null;
            let value = cellValue(r, c, sh);
            if (isError(value)) value = { error: value.value };
            return { formula: isFormula(raw) ? raw : null, value };
          },
          style: (r, c) => sh.styles.get(key(r, c)) || null,
          merges: sh.merges,
          colW: sh.colW,
          rowH: sh.rowH,
          freezeRow: sh.freezeRow,
          freezeCol: sh.freezeCol
        };
      })
    };
  }

  async function save(fmt) {
    if (S.edit && !commitEdit()) return;
    const base = S.fileName.replace(/\.[^.]+$/, '') || 'planilla';
    try {
      let blob;
      if (fmt === 'xlsx') blob = await IO.writeXlsx(exportSource());
      else if (fmt === 'xls') blob = IO.writeSheetJS(exportSource(), 'biff8');
      else if (fmt === 'ods') blob = IO.writeSheetJS(exportSource(), 'ods');
      else {
        const sh = sheet();
        const ext = usedExtent(sh);
        const rows = [];
        for (let r = 0; r < ext.rows; r++) {
          const row = [];
          for (let c = 0; c < ext.cols; c++) row.push(displayText(r, c, sh));
          rows.push(row);
        }
        while (rows.length && rows[rows.length - 1].every(t => t === '')) rows.pop();
        blob = IO.writeDelimited(rows, fmt === 'tsv' ? '\t' : ';');
      }
      IO.download(blob, base + '.' + fmt);
      if (fmt === 'xlsx' || fmt === 'xls' || fmt === 'ods') setDirty(false);
      const extra = fmt === 'csv' || fmt === 'tsv' ? ' Solo se guardó la hoja "' + sheet().name + '".' : '';
      toast('Se descargó ' + base + '.' + fmt + ' en tu carpeta de Descargas.' + extra);
    } catch (e) {
      console.error(e);
      toast('No se pudo guardar el archivo: ' + (e.message || e), 'error');
    }
  }

  // ---------- Imprimir ----------
  function buildPrint() {
    const sh = sheet();
    const ext = usedExtent(sh);
    const area = $('print');
    if (!ext.rows || !ext.cols) { area.innerHTML = ''; return; }
    const covered = new Set();
    const merged = new Map();
    sh.merges.forEach(m => {
      merged.set(key(m.r0, m.c0), m);
      for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(key(r, c));
    });
    const parts = ['<table><colgroup>'];
    let w = 0;
    for (let c = 0; c < ext.cols; c++) { const cw = colWidth(c); w += cw; parts.push('<col style="width:' + cw + 'px">'); }
    parts.push('</colgroup>');
    for (let r = 0; r < ext.rows; r++) {
      if (sh.hidden.has(r)) continue;
      parts.push('<tr style="height:' + rowHeight(r) + 'px">');
      for (let c = 0; c < ext.cols; c++) {
        const k = key(r, c);
        if (covered.has(k)) continue;
        const v = cellValue(r, c, sh);
        const st = sh.styles.get(k);
        let css = cellStyleCss(st, r, c);
        if (!st || !st.align) {
          if (typeof v === 'number') css += 'text-align:right;';
          else if (isError(v) || typeof v === 'boolean' || merged.has(k)) css += 'text-align:center;';
        }
        if (st && st.bt) css += 'border-top:1px solid #000;';
        if (st && st.bb) css += 'border-bottom:1px solid #000;';
        if (st && st.bl) css += 'border-left:1px solid #000;';
        if (st && st.br) css += 'border-right:1px solid #000;';
        const m = merged.get(k);
        const span = m ? (m.r1 > m.r0 ? ' rowspan="' + (m.r1 - m.r0 + 1) + '"' : '') + (m.c1 > m.c0 ? ' colspan="' + (m.c1 - m.c0 + 1) + '"' : '') : '';
        parts.push('<td' + span + (css ? ' style="' + css + '"' : '') + '>' + esc(displayText(r, c, sh)) + '</td>');
      }
      parts.push('</tr>');
    }
    parts.push('</table>');
    area.innerHTML = parts.join('');
    area.firstChild.style.width = w + 'px';
  }

  function doPrint() {
    if (S.edit && !commitEdit()) return;
    buildPrint();
    window.print();
  }

  // ---------- Avisos y diálogos ----------
  function toast(msg, type, ms) {
    const el = document.createElement('div');
    el.className = 'toast' + (type ? ' ' + type : '');
    el.textContent = msg;
    $('toasts').append(el);
    if (ms !== 0) setTimeout(() => el.remove(), ms || (type === 'error' ? 6000 : 3500));
    return el;
  }

  function modal(build) {
    return new Promise(resolve => {
      const back = $('modal');
      const box = back.firstElementChild;
      box.innerHTML = '';
      const close = v => { back.hidden = true; document.removeEventListener('keydown', onKey, true); resolve(v); if (!S.edit) focusGrid(); };
      const onKey = e => {
        if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(null); }
        else if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); const p = box.querySelector('button.primary'); if (p) p.click(); }
        else e.stopPropagation();
      };
      build(box, close);
      back.hidden = false;
      document.addEventListener('keydown', onKey, true);
      const first = box.querySelector('input') || box.querySelector('button.primary');
      if (first) { first.focus(); if (first.select) first.select(); }
    });
  }

  function confirmModal(title, text, okLabel) {
    return modal((box, close) => {
      box.innerHTML = '<h2></h2><p></p><div class="modal-actions"><button data-v="0">Cancelar</button><button class="primary" data-v="1"></button></div>';
      box.querySelector('h2').textContent = title;
      box.querySelector('p').textContent = text;
      box.querySelector('.primary').textContent = okLabel || 'Aceptar';
      box.querySelectorAll('button').forEach(b => { b.onclick = () => close(b.dataset.v === '1'); });
    });
  }

  function numberModal(title, label, value) {
    return modal((box, close) => {
      box.innerHTML = '<h2></h2><p></p><input type="number" min="4" max="1000" id="modalNumber"><div class="modal-actions"><button>Cancelar</button><button class="primary">Aceptar</button></div>';
      box.querySelector('h2').textContent = title;
      box.querySelector('p').textContent = label;
      const inp = box.querySelector('input');
      inp.value = value;
      const [cancel, ok] = box.querySelectorAll('button');
      cancel.onclick = () => close(null);
      ok.onclick = () => { const n = parseInt(inp.value, 10); close(n >= 4 && n <= 1000 ? n : null); };
    });
  }

  async function askColWidth() {
    const s = S.sel;
    const n = await numberModal('Ancho de columna', 'Ancho en píxeles (el estándar es ' + IO.DEFAULT_COL_W + ').', colWidth(s.c0));
    if (n) commit(() => { for (let c = s.c0; c <= s.c1; c++) sheet().colW[c] = n; });
  }
  async function askRowHeight() {
    const s = S.sel;
    const n = await numberModal('Alto de fila', 'Alto en píxeles (el estándar es ' + IO.DEFAULT_ROW_H + ').', rowHeight(s.r0));
    if (n) commit(() => { for (let r = s.r0; r <= s.r1; r++) sheet().rowH[r] = n; });
  }

  // ---------- Menús ----------
  function closeMenus() {
    document.querySelectorAll('.menu.open').forEach(m => m.classList.remove('open'));
    $('ctxMenu').hidden = true;
  }

  function showContextMenu(x, y, items) {
    const menu = $('ctxMenu');
    menu.innerHTML = '';
    items.forEach(it => {
      if (it === '-') { menu.append(document.createElement('hr')); return; }
      const b = document.createElement('button');
      b.textContent = it[0];
      b.onclick = () => { closeMenus(); it[1](); };
      menu.append(b);
    });
    menu.hidden = false;
    menu.style.left = Math.min(x, window.innerWidth - menu.offsetWidth - 8) + 'px';
    menu.style.top = Math.min(y, window.innerHeight - menu.offsetHeight - 8) + 'px';
  }

  const PALETTE = [
    '#000000', '#434343', '#666666', '#999999', '#B7B7B7', '#D9D9D9', '#EFEFEF', '#FFFFFF',
    '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0',
    '#F4CCCC', '#FCE5CD', '#FFF2CC', '#D9EAD3', '#D0E0E3', '#CFE2F3', '#D9D2E9', '#EAD1DC',
    '#E06666', '#F6B26B', '#FFD966', '#93C47D', '#76A5AF', '#6FA8DC', '#8E7CC3', '#C27BA0',
    '#990000', '#B45F06', '#BF9000', '#38761D', '#134F5C', '#0B5394', '#351C75', '#741B47'
  ];

  function buildPalette(menuId, prop, noneLabel, swatchId) {
    const menu = $(menuId);
    const grid = document.createElement('div');
    grid.className = 'colors';
    PALETTE.forEach(col => {
      const b = document.createElement('button');
      b.style.background = col;
      b.title = col;
      b.onclick = () => { closeMenus(); $(swatchId).style.background = col; setStyleAll({ [prop]: col }); };
      grid.append(b);
    });
    const row = document.createElement('div');
    row.className = 'pal-row';
    const none = document.createElement('button');
    none.textContent = noneLabel;
    none.onclick = () => { closeMenus(); setStyleAll({ [prop]: null }); };
    const more = document.createElement('label');
    more.textContent = 'Más colores…';
    const inp = document.createElement('input');
    inp.type = 'color';
    inp.id = menuId + 'Custom';
    inp.onchange = () => { closeMenus(); $(swatchId).style.background = inp.value; setStyleAll({ [prop]: inp.value.toUpperCase() }); };
    more.append(inp);
    row.append(none, more);
    menu.append(grid, row);
  }

  // ---------- Íconos ----------
  const ICONS = {
    open: '<path d="M3 7h6l2 2h10v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z"/><path d="M3 7V5a1 1 0 0 1 1-1h5l2 2"/>',
    save: '<path d="M5 3h11l4 4v13a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"/><path d="M8 3v5h8V3M7 21v-7h10v7"/>',
    print: '<path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="1"/><path d="M7 14h10v7H7z"/>',
    undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>',
    redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>',
    paste: '<rect x="5" y="4" width="14" height="17" rx="1"/><path d="M9 4V2h6v2M9 10h6M9 14h6M9 18h4"/>',
    cut: '<circle cx="6" cy="18" r="3"/><circle cx="18" cy="18" r="3"/><path d="M8 16 19 3M16 16 5 3"/>',
    copy: '<rect x="8" y="8" width="13" height="13" rx="1"/><path d="M5 16H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h11a1 1 0 0 1 1 1v1"/>',
    borders: '<rect x="3" y="3" width="18" height="18"/><path d="M3 12h18M12 3v18" stroke-dasharray="2 2"/>',
    fill: '<path d="m5 11 7-7 8 8-7 7z"/><path d="M5 11h15"/><path d="M20 16s2 2.5 2 3.5a2 2 0 0 1-4 0c0-1 2-3.5 2-3.5z"/>',
    fontcolor: '<path d="m6 16 6-13 6 13M8.5 11h7"/>',
    alignl: '<path d="M3 5h18M3 10h12M3 15h18M3 20h12"/>',
    alignc: '<path d="M3 5h18M6 10h12M3 15h18M6 20h12"/>',
    alignr: '<path d="M3 5h18M9 10h12M3 15h18M9 20h12"/>',
    merge: '<rect x="3" y="5" width="18" height="14"/><path d="m7 12 3-3M7 12l3 3M7 12h10M17 12l-3-3M17 12l-3 3"/>',
    insert: '<rect x="3" y="3" width="18" height="18"/><path d="M12 8v8M8 12h8"/>',
    delete: '<rect x="3" y="3" width="18" height="18"/><path d="m9 9 6 6M15 9l-6 6"/>',
    sortasc: '<path d="M7 4v16M3 16l4 4 4-4"/><path d="M14 4h6M14 10h4M14 16h2"/>',
    sortdesc: '<path d="M7 4v16M3 16l4 4 4-4"/><path d="M14 16h6M14 10h4M14 4h2"/>',
    filter: '<path d="M3 4h18l-7 8v7l-4 2v-9z"/>',
    search: '<circle cx="11" cy="11" r="7"/><path d="m21 21-5-5"/>',
    freeze: '<rect x="3" y="3" width="18" height="18"/><path d="M3 9h18M9 3v18" stroke-width="2.5"/>',
    sheet: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>'
  };

  function injectIcons() {
    document.querySelectorAll('[data-icon]').forEach(el => {
      const svg = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + ICONS[el.dataset.icon] + '</svg>';
      el.insertAdjacentHTML('afterbegin', svg);
    });
  }

  // ---------- Eventos ----------
  function cellFromEvent(e) {
    const td = e.target.closest && e.target.closest('td[data-r]');
    if (!td) return null;
    return { r: +td.dataset.r, c: +td.dataset.c, td };
  }

  /** Coordenadas del puntero dentro de la cuadrícula (sin contar el scroll visible). */
  function gridPoint(x, y) {
    const sc = $('scroller');
    const rect = sc.getBoundingClientRect();
    return {
      px: clamp(x, rect.left + 1, rect.right - 2) - rect.left + sc.scrollLeft,
      py: clamp(y, rect.top + 1, rect.bottom - 2) - rect.top + sc.scrollTop
    };
  }

  function cellAtPoint(x, y) {
    const sc = $('scroller');
    const rect = sc.getBoundingClientRect();
    const L = S.L;
    const { px, py } = gridPoint(x, y);
    let c = 0, r = 0;
    while (c < L.cols - 1 && L.colX[c] + L.colW[c] <= px) c++;
    while (r < L.rows - 1 && L.rowY[r] + L.rowH[r] <= py) r++;
    // Ajuste por filas/columnas inmovilizadas
    const sh = sheet();
    const el = document.elementFromPoint(clamp(x, rect.left + 1, rect.right - 2), clamp(y, rect.top + 1, rect.bottom - 2));
    const td = el && el.closest && el.closest('td[data-r]');
    if (td && (sh.freezeRow || sh.freezeCol)) return { r: +td.dataset.r, c: +td.dataset.c };
    return { r, c };
  }

  function autoScroll(e) {
    const sc = $('scroller');
    const rect = sc.getBoundingClientRect();
    if (e.clientY > rect.bottom - 16) sc.scrollTop += 20;
    else if (e.clientY < rect.top + S.L.headH + 8) sc.scrollTop -= 20;
    if (e.clientX > rect.right - 16) sc.scrollLeft += 20;
    else if (e.clientX < rect.left + headW() + 8) sc.scrollLeft -= 20;
  }

  function bindGrid() {
    const grid = $('grid');
    const scroller = $('scroller');

    grid.addEventListener('mousedown', e => {
      if (e.button !== 0 && e.button !== 2) return;
      closeMenus();
      // Botón de filtro
      const fb = e.target.closest('.fbtn');
      if (fb) { e.preventDefault(); openFilterMenu(+fb.dataset.fc, fb); return; }
      // Cambiar ancho / alto
      const rz = e.target.closest('[data-rzc],[data-rzr]');
      if (rz && e.button === 0) {
        e.preventDefault();
        const isCol = rz.dataset.rzc !== undefined;
        const idx = +(isCol ? rz.dataset.rzc : rz.dataset.rzr);
        S.drag = { type: isCol ? 'rzc' : 'rzr', idx, start: isCol ? e.clientX : e.clientY, orig: isCol ? colPx(idx) : rowPx(idx) };
        return;
      }
      const th = e.target.closest('th');
      if (th) {
        e.preventDefault();
        if (S.edit && !commitEdit()) return;
        if (th.dataset.corner) { S.anchor = { r: 0, c: 0 }; select(0, 0, S.L.rows - 1, S.L.cols - 1, 0, 0, true); return; }
        if (th.dataset.hc !== undefined) {
          const c = +th.dataset.hc;
          if (e.button === 2 && c >= S.sel.c0 && c <= S.sel.c1 && S.sel.r0 === 0) return;
          if (e.button === 0 && !e.shiftKey && isWholeCols() && !isWholeRows() && c >= S.sel.c0 && c <= S.sel.c1) {
            S.drag = { type: 'movecols', c, startX: e.clientX, active: false, from: S.sel.c0, count: S.sel.c1 - S.sel.c0 + 1 };
            return;
          }
          if (e.shiftKey) { select(0, S.anchor.c, S.L.rows - 1, c, 0, S.anchor.c, true); }
          else { S.anchor = { r: 0, c }; select(0, c, S.L.rows - 1, c, 0, c, true); }
          S.drag = { type: 'cols' };
        } else if (th.dataset.hr !== undefined) {
          const r = +th.dataset.hr;
          if (e.button === 2 && r >= S.sel.r0 && r <= S.sel.r1 && S.sel.c0 === 0) return;
          if (e.button === 0 && !e.shiftKey && isWholeRows() && !isWholeCols() && r >= S.sel.r0 && r <= S.sel.r1) {
            S.drag = { type: 'moverows', r, startY: e.clientY, active: false, from: S.sel.r0, count: S.sel.r1 - S.sel.r0 + 1 };
            return;
          }
          if (e.shiftKey) { select(S.anchor.r, 0, r, S.L.cols - 1, S.anchor.r, 0, true); }
          else { S.anchor = { r, c: 0 }; select(r, 0, r, S.L.cols - 1, r, 0, true); }
          S.drag = { type: 'rows' };
        }
        if (e.button === 2) S.drag = null;
        return;
      }
      const cell = cellFromEvent(e);
      if (!cell) return;
      if (S.edit && canPoint() && e.button === 0) {
        e.preventDefault();
        S.drag = { type: 'point', r: cell.r, c: cell.c };
        insertRef(cell.r, cell.c, cell.r, cell.c);
        return;
      }
      if (S.edit) { if (!commitEdit()) { e.preventDefault(); return; } }
      e.preventDefault();
      if (e.button === 2) {
        const s = S.sel;
        if (cell.r >= s.r0 && cell.r <= s.r1 && cell.c >= s.c0 && cell.c <= s.c1) { focusGrid(); return; }
      }
      if (e.shiftKey) extendTo(cell.r, cell.c);
      else { selectCell(cell.r, cell.c); }
      if (e.button === 0) S.drag = { type: 'cells' };
      focusGrid();
    });

    grid.addEventListener('dblclick', e => {
      const rz = e.target.closest('[data-rzc],[data-rzr]');
      if (rz) return;
      const cell = cellFromEvent(e);
      if (!cell || S.edit) return;
      startEdit('edit');
    });

    grid.addEventListener('contextmenu', e => {
      e.preventDefault();
      const th = e.target.closest('th');
      if (th && th.dataset.hc !== undefined) {
        showContextMenu(e.clientX, e.clientY, [
          ['Cortar', () => copyFromButton(true)], ['Copiar', () => copyFromButton(false)], ['Pegar', pasteFromButton], '-',
          ['Insertar columnas', insertCols], ['Eliminar columnas', deleteCols], ['Borrar contenido', clearContents], '-',
          ['Ancho de columna…', askColWidth]
        ]);
      } else if (th && th.dataset.hr !== undefined) {
        showContextMenu(e.clientX, e.clientY, [
          ['Cortar', () => copyFromButton(true)], ['Copiar', () => copyFromButton(false)], ['Pegar', pasteFromButton], '-',
          ['Insertar filas', insertRows], ['Eliminar filas', deleteRows], ['Borrar contenido', clearContents], '-',
          ['Alto de fila…', askRowHeight]
        ]);
      } else if (cellFromEvent(e)) {
        showContextMenu(e.clientX, e.clientY, [
          ['Cortar', () => copyFromButton(true)], ['Copiar', () => copyFromButton(false)], ['Pegar', pasteFromButton], '-',
          ['Insertar filas arriba', insertRows], ['Insertar columnas a la izquierda', insertCols],
          ['Eliminar filas', deleteRows], ['Eliminar columnas', deleteCols], '-',
          ['Borrar contenido', clearContents], ['Combinar o separar celdas', toggleMerge], '-',
          ['Ordenar de A a Z', () => sortCommand(false)], ['Ordenar de Z a A', () => sortCommand(true)]
        ]);
      }
    });

    $('fillHandle').addEventListener('mousedown', e => {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      if (S.edit && !commitEdit()) return;
      S.drag = { type: 'fill', dir: null, count: 0 };
    });

    document.addEventListener('mousemove', e => {
      const d = S.drag;
      if (!d) return;
      if (d.type === 'rzc' || d.type === 'rzr') {
        const delta = (d.type === 'rzc' ? e.clientX : e.clientY) - d.start;
        const size = Math.max(d.type === 'rzc' ? 16 : 8, Math.round(d.orig + delta));
        d.size = size;
        if (d.type === 'rzc') {
          const col = $('grid').querySelector('colgroup').children[d.idx + 1];
          if (col) col.style.width = size + 'px';
        } else {
          const tr = S.L.trMap.get(d.idx);
          if (tr) tr.style.height = size + 'px';
        }
        return;
      }
      if (d.type === 'movecols' || d.type === 'moverows') {
        const isCol = d.type === 'movecols';
        if (!d.active && Math.abs((isCol ? e.clientX - d.startX : e.clientY - d.startY)) < 5) return;
        d.active = true;
        document.body.classList.add('moving');
        autoScroll(e);
        d.boundary = boundaryAt(isCol ? 'col' : 'row', e.clientX, e.clientY);
        if (d.boundary >= d.from && d.boundary <= d.from + d.count) $('moveLine').hidden = true;
        else showMoveLine(isCol ? 'col' : 'row', d.boundary);
        return;
      }
      autoScroll(e);
      const p = cellAtPoint(e.clientX, e.clientY);
      if (d.type === 'cells') {
        if (p.r !== S.sel.r1 || p.c !== S.sel.c1 || true) {
          select(S.anchor.r, S.anchor.c, p.r, p.c, S.anchor.r, S.anchor.c, true);
          S.sel.ar = S.anchor.r; S.sel.ac = S.anchor.c;
          updateSelectionUI();
        }
      } else if (d.type === 'cols') {
        select(0, S.anchor.c, S.L.rows - 1, p.c, 0, S.anchor.c, true);
      } else if (d.type === 'rows') {
        select(S.anchor.r, 0, p.r, S.L.cols - 1, S.anchor.r, 0, true);
      } else if (d.type === 'point') {
        insertRef(d.r, d.c, p.r, p.c);
      } else if (d.type === 'fill') {
        const s = S.sel;
        const down = p.r - s.r1, up = s.r0 - p.r, right = p.c - s.c1, left = s.c0 - p.c;
        const best = Math.max(down, up, right, left);
        const prev = $('fillPreview');
        if (best <= 0) { d.dir = null; prev.hidden = true; return; }
        if (best === down) { d.dir = 'down'; d.count = down; place(prev, rectOf(s.r0, s.c0, p.r, s.c1)); }
        else if (best === up) { d.dir = 'up'; d.count = up; place(prev, rectOf(p.r, s.c0, s.r1, s.c1)); }
        else if (best === right) { d.dir = 'right'; d.count = right; place(prev, rectOf(s.r0, s.c0, s.r1, p.c)); }
        else { d.dir = 'left'; d.count = left; place(prev, rectOf(s.r0, p.c, s.r1, s.c1)); }
        prev.hidden = false;
      }
    });

    document.addEventListener('mouseup', () => {
      const d = S.drag;
      S.drag = null;
      if (!d) return;
      if ((d.type === 'rzc' || d.type === 'rzr') && d.size) {
        const s = S.sel;
        commit(() => {
          if (d.type === 'rzc') {
            const all = s.r0 === 0 && s.r1 >= S.L.rows - 1 && d.idx >= s.c0 && d.idx <= s.c1;
            const base = Math.max(1, Math.round(d.size / S.zoom));
            for (let c = all ? s.c0 : d.idx; c <= (all ? s.c1 : d.idx); c++) sheet().colW[c] = base;
          } else {
            const all = s.c0 === 0 && s.c1 >= S.L.cols - 1 && d.idx >= s.r0 && d.idx <= s.r1;
            const base = Math.max(1, Math.round(d.size / S.zoom));
            for (let r = all ? s.r0 : d.idx; r <= (all ? s.r1 : d.idx); r++) sheet().rowH[r] = base;
          }
        });
      } else if (d.type === 'movecols' || d.type === 'moverows') {
        endMoveDrag();
        const isCol = d.type === 'movecols';
        if (!d.active) {
          // Fue un clic: queda seleccionada solo esa columna o fila.
          if (isCol) { S.anchor = { r: 0, c: d.c }; select(0, d.c, S.L.rows - 1, d.c, 0, d.c, true); }
          else { S.anchor = { r: d.r, c: 0 }; select(d.r, 0, d.r, S.L.cols - 1, d.r, 0, true); }
        } else if (d.boundary !== undefined) {
          moveLines(isCol ? 'col' : 'row', d.from, d.count, d.boundary);
        }
      } else if (d.type === 'fill') {
        $('fillPreview').hidden = true;
        if (d.dir) doFill(d.dir, d.count);
      } else if (d.type === 'point') {
        S.edit && (S.edit.refStart = -1);
      }
    });

    let raf = 0;
    scroller.addEventListener('scroll', () => {
      if (!raf) raf = requestAnimationFrame(() => { raf = 0; renderWindow(false); });
      closeMenus();
      if (!$('filterMenu').hidden) $('filterMenu').hidden = true;
    });
  }

  function onKeyDown(e) {
    if (!S.loaded) return;
    const t = e.target;
    const inEditor = t === $('editor') || t === $('fbar');
    if (t === $('nameBox') || t.closest && (t.closest('#findPanel') || t.closest('#filterMenu') || t.closest('.tab'))) {
      if (t.closest && t.closest('#findPanel')) {
        if (e.key === 'Escape') { e.preventDefault(); $('findPanel').hidden = true; focusGrid(); }
        else if (e.key === 'Enter') { e.preventDefault(); findNext(); }
      }
      if (t.closest && t.closest('#filterMenu') && e.key === 'Escape') closeFilterMenu();
      return;
    }
    if (!$('modal').hidden) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    const lower = k.length === 1 ? k.toLowerCase() : k;

    // Atajos de formato (también mientras no se edita)
    if (ctrl && e.altKey && lower === 'n') { e.preventDefault(); if (!S.edit) toggleStyle('b'); return; }

    if (inEditor && S.edit) {
      if (k === 'Enter') { e.preventDefault(); if (commitEdit()) moveActive(e.shiftKey ? -1 : 1, 0); return; }
      if (k === 'Tab') { e.preventDefault(); if (commitEdit()) moveActive(0, e.shiftKey ? -1 : 1); return; }
      if (k === 'Escape') { e.preventDefault(); cancelEdit(); return; }
      if (k === 'F2') { e.preventDefault(); S.edit.mode = S.edit.mode === 'enter' ? 'edit' : 'enter'; return; }
      if (S.edit.mode === 'enter' && k.startsWith('Arrow') && $('editor').value[0] !== '=') {
        e.preventDefault();
        if (commitEdit()) moveActive(k === 'ArrowDown' ? 1 : k === 'ArrowUp' ? -1 : 0, k === 'ArrowRight' ? 1 : k === 'ArrowLeft' ? -1 : 0);
        return;
      }
      S.edit.refStart = -1;
      return;
    }
    if (inEditor) return;

    if (ctrl && !e.altKey) {
      if (lower === 'z') { e.preventDefault(); undo(); return; }
      if (lower === 'y') { e.preventDefault(); redo(); return; }
      if (lower === 'b') { e.preventDefault(); openFind(); return; }
      if (lower === 'k') { e.preventDefault(); toggleStyle('i'); return; }
      if (lower === 's') { e.preventDefault(); toggleStyle('u'); return; }
      if (lower === 'a') { e.preventDefault(); const s = S.sel; const rg = isEmptyCell(s.ar, s.ac) ? null : currentRegion(s.ar, s.ac); if (rg && !(rg.r0 === s.r0 && rg.r1 === s.r1 && rg.c0 === s.c0 && rg.c1 === s.c1)) select(rg.r0, rg.c0, rg.r1, rg.c1, s.ar, s.ac, true); else select(0, 0, S.L.rows - 1, S.L.cols - 1, s.ar, s.ac, true); return; }
      if (k === 'Home') { e.preventDefault(); selectCell(0, 0); return; }
      if (k === 'End') { e.preventDefault(); const ex = usedExtent(); selectCell(Math.max(0, ex.rows - 1), Math.max(0, ex.cols - 1)); return; }
      if (lower === 'c' || lower === 'x' || lower === 'v') return; // eventos nativos copy/cut/paste
    }

    switch (k) {
      case 'ArrowUp': e.preventDefault(); moveActive(-1, 0, e.shiftKey, ctrl); return;
      case 'ArrowDown': e.preventDefault(); moveActive(1, 0, e.shiftKey, ctrl); return;
      case 'ArrowLeft': e.preventDefault(); moveActive(0, -1, e.shiftKey, ctrl); return;
      case 'ArrowRight': e.preventDefault(); moveActive(0, 1, e.shiftKey, ctrl); return;
      case 'Enter': e.preventDefault(); moveActive(e.shiftKey ? -1 : 1, 0); return;
      case 'Tab': e.preventDefault(); moveActive(0, e.shiftKey ? -1 : 1); return;
      case 'PageDown': e.preventDefault(); moveActive(20, 0, e.shiftKey); return;
      case 'PageUp': e.preventDefault(); moveActive(-20, 0, e.shiftKey); return;
      case 'Home': e.preventDefault(); selectCell(S.sel.ar, 0); return;
      case 'F2': e.preventDefault(); startEdit('edit'); return;
      case 'Delete': e.preventDefault(); clearContents(); return;
      case 'Backspace': e.preventDefault(); startEdit('enter', ''); return;
      case 'Escape': if (S.drag && S.drag.active) { S.drag = null; endMoveDrag(); } clearClip(); closeMenus(); return;
    }
  }

  function bindUI() {
    injectIcons();
    buildPalette('menuFill', 'bg', 'Sin relleno', 'swFill');
    buildPalette('menuColor', 'color', 'Automático', 'swColor');
    document.querySelectorAll('.needs-file').forEach(b => { b.disabled = true; });

    const fileInput = $('fileInput');
    $('btnOpen').onclick = $('btnOpen2').onclick = () => { fileInput.value = ''; fileInput.click(); };
    fileInput.onchange = () => openFile(fileInput.files[0]);

    // Menús desplegables de la cinta
    document.querySelectorAll('.ribbon .dd > button').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        const menu = btn.nextElementSibling;
        const was = menu.classList.contains('open');
        closeMenus();
        if (!was) {
          menu.classList.add('open');
          const r = btn.getBoundingClientRect();
          menu.style.position = 'fixed';
          menu.style.left = Math.min(r.left, window.innerWidth - menu.offsetWidth - 8) + 'px';
          menu.style.top = r.bottom + 2 + 'px';
        }
      });
    });
    document.addEventListener('mousedown', e => {
      if (!e.target.closest('.menu') && !e.target.closest('.dd > button')) closeMenus();
      if (!$('filterMenu').hidden && !e.target.closest('#filterMenu') && !e.target.closest('.fbtn')) $('filterMenu').hidden = true;
    });

    document.querySelectorAll('[data-save]').forEach(b => { b.onclick = () => { closeMenus(); save(b.dataset.save); }; });
    document.querySelectorAll('[data-border]').forEach(b => { b.onclick = () => { closeMenus(); applyBorders(b.dataset.border); }; });
    document.querySelectorAll('[data-freeze]').forEach(b => { b.onclick = () => { closeMenus(); setFreeze(b.dataset.freeze); }; });
    const cmds = { insertRows, insertCols, deleteRows, deleteCols };
    document.querySelectorAll('[data-cmd]').forEach(b => { b.onclick = () => { closeMenus(); cmds[b.dataset.cmd](); }; });

    const withGrid = fn => () => { if (S.edit && !commitEdit()) return; fn(); focusGrid(); };
    $('btnPrint').onclick = doPrint;
    $('btnUndo').onclick = withGrid(undo);
    $('btnRedo').onclick = withGrid(redo);
    $('btnPaste').onclick = withGrid(pasteFromButton);
    $('btnCut').onclick = withGrid(() => copyFromButton(true));
    $('btnCopy').onclick = withGrid(() => copyFromButton(false));
    $('btnBold').onclick = withGrid(() => toggleStyle('b'));
    $('btnItalic').onclick = withGrid(() => toggleStyle('i'));
    $('btnUnderline').onclick = withGrid(() => toggleStyle('u'));
    $('btnAlignL').onclick = withGrid(() => setAlign('left'));
    $('btnAlignC').onclick = withGrid(() => setAlign('center'));
    $('btnAlignR').onclick = withGrid(() => setAlign('right'));
    $('btnMerge').onclick = withGrid(toggleMerge);
    $('btnCurrency').onclick = withGrid(() => setFormat('moneda'));
    $('btnPercent').onclick = withGrid(() => setFormat('porcentaje'));
    $('btnDecInc').onclick = withGrid(() => changeDecimals(1));
    $('btnDecDec').onclick = withGrid(() => changeDecimals(-1));
    $('btnSortAsc').onclick = withGrid(() => sortCommand(false));
    $('btnSortDesc').onclick = withGrid(() => sortCommand(true));
    $('btnFilter').onclick = withGrid(toggleFilter);
    $('btnFind').onclick = () => { if (S.edit && !commitEdit()) return; openFind(); };
    $('btnAddSheet').onclick = () => { if (S.edit && !commitEdit()) return; addSheet(); };
    $('zoomOut').onclick = () => stepZoom(-1);
    $('zoomIn').onclick = () => stepZoom(1);
    $('zoomLbl').onclick = () => setZoom(1);
    updateZoomUI();
    $('fontSize').onchange = withGrid(() => { const n = +$('fontSize').value; setStyleAll({ size: n === 11 ? null : n }); });
    $('numFormat').onchange = withGrid(() => setFormat($('numFormat').value));

    // Buscar y reemplazar
    $('findClose').onclick = () => { $('findPanel').hidden = true; focusGrid(); };
    $('btnFindNext').onclick = () => findNext();
    $('btnReplace').onclick = replaceOne;
    $('btnReplaceAll').onclick = replaceAll;

    // Pestañas
    const tabs = $('tabs');
    tabs.addEventListener('click', e => { const t = e.target.closest('.tab'); if (t && t.dataset.idx !== undefined && !t.querySelector('input')) switchSheet(+t.dataset.idx); });
    tabs.addEventListener('dblclick', e => { const t = e.target.closest('.tab'); if (t) renameSheet(+t.dataset.idx); });
    tabs.addEventListener('contextmenu', e => {
      const t = e.target.closest('.tab');
      if (!t) return;
      e.preventDefault();
      const i = +t.dataset.idx;
      switchSheet(i);
      showContextMenu(e.clientX, e.clientY - 70, [['Cambiar nombre', () => renameSheet(i)], ['Eliminar hoja', () => deleteSheet(i)], '-', ['Agregar hoja', addSheet]]);
    });

    // Cuadro de nombre: ir a una celda o rango
    $('nameBox').addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const v = $('nameBox').value.trim().toUpperCase().replace(/\$/g, '');
        const m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(v);
        if (!m) { toast('Escribí una celda como B3 o un rango como A1:C10.', 'error'); return; }
        const r0 = +m[2] - 1, c0 = F.colToNum(m[1]);
        if (m[3]) { S.anchor = { r: r0, c: c0 }; select(r0, c0, +m[4] - 1, F.colToNum(m[3]), r0, c0); }
        else selectCell(r0, c0);
        focusGrid();
      } else if (e.key === 'Escape') { e.preventDefault(); updateSelectionUI(); focusGrid(); }
    });

    // Barra de fórmulas
    const fbar = $('fbar');
    fbar.addEventListener('focus', () => {
      if (!S.loaded) return;
      if (!S.edit) {
        const pos = fbar.selectionStart;
        startEdit('edit');
        S.edit.source = 'bar';
        fbar.focus();
        try { fbar.setSelectionRange(pos, pos); } catch (err) { /* */ }
      } else S.edit.source = 'bar';
    });
    fbar.addEventListener('input', () => { if (S.edit) { S.edit.refStart = -1; syncEditors(fbar); } });
    $('editor').addEventListener('input', () => { if (S.edit) { S.edit.refStart = -1; syncEditors($('editor')); } });
    $('editor').addEventListener('focus', () => { if (S.edit) S.edit.source = 'cell'; });

    // Teclado
    document.addEventListener('keydown', onKeyDown);
    const kbd = $('kbd');
    kbd.addEventListener('input', () => {
      const v = kbd.value;
      // Solo se empieza a editar con caracteres visibles (evita que un Enter o espacio sueltos borren la celda).
      if (!S.loaded || S.edit || !v.trim()) { kbd.value = ' '; kbd.select(); return; }
      startEdit('enter', v.replace(/^ /, ''));
    });
    kbd.addEventListener('blur', () => { /* se recupera al hacer clic en la cuadrícula */ });

    // Portapapeles
    document.addEventListener('copy', e => {
      if (!S.loaded || S.edit || isTextField(e.target)) return;
      doCopy(e, S.pendingCut);
    });
    document.addEventListener('cut', e => {
      if (!S.loaded || S.edit || isTextField(e.target)) return;
      doCopy(e, true);
    });
    document.addEventListener('paste', e => {
      if (!S.loaded || S.edit || isTextField(e.target)) return;
      e.preventDefault();
      doPaste(e.clipboardData.getData('text/plain'));
    });

    // Arrastrar y soltar archivos
    let dragDepth = 0;
    window.addEventListener('dragenter', e => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      dragDepth++;
      $('dropOverlay').hidden = false;
    });
    window.addEventListener('dragover', e => { if (hasFiles(e)) e.preventDefault(); });
    window.addEventListener('dragleave', e => {
      if (!hasFiles(e)) return;
      dragDepth = Math.max(0, dragDepth - 1);
      if (!dragDepth) $('dropOverlay').hidden = true;
    });
    window.addEventListener('drop', e => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      dragDepth = 0;
      $('dropOverlay').hidden = true;
      const f = e.dataTransfer.files[0];
      if (f) openFile(f);
    });

    window.addEventListener('beforeprint', () => { if (S.loaded) buildPrint(); });
    window.addEventListener('beforeunload', e => { if (S.dirty) { e.preventDefault(); e.returnValue = ''; } });
    window.addEventListener('resize', () => { closeMenus(); renderWindow(false); });

    bindGrid();
  }

  function isTextField(el) {
    return el && el !== $('kbd') && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA');
  }
  function hasFiles(e) {
    return e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
  }

  bindUI();
  window.CoftechExcel = { openFile, state: S };
})();
