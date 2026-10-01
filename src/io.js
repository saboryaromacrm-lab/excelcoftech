/* Coftech Excel — abrir y guardar archivos. */
(function (root) {
  'use strict';

  const F = root.CEFormulas;
  const N = root.CEFormat;

  const DEFAULT_COL_W = 80;
  const DEFAULT_ROW_H = 21;
  const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

  function newSheet(name) {
    return { name, cells: [], styles: new Map(), merges: [], colW: {}, rowH: {}, freezeRow: false, freezeCol: false };
  }

  const key = (r, c) => r + ',' + c;

  function applyTint(hex, tint) {
    if (!tint) return hex;
    const rgb = [0, 2, 4].map(i => parseInt(hex.substr(i, 2), 16));
    const out = rgb.map(v => Math.round(tint < 0 ? v * (1 + tint) : v + (255 - v) * tint));
    return out.map(v => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  function excelColor(col) {
    if (!col) return null;
    if (col.argb && /^[0-9A-Fa-f]{8}$/.test(col.argb)) return '#' + col.argb.slice(2).toUpperCase();
    if (typeof col.theme === 'number' && THEME[col.theme]) return '#' + applyTint(THEME[col.theme], col.tint);
    return null;
  }

  function serialFromDate(d) {
    return (d.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
  }

  function cellValueFromExcel(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number' || typeof v === 'boolean') return v;
    if (typeof v === 'string') return "'" + v;
    if (v instanceof Date) return serialFromDate(v);
    if (v.richText) return "'" + v.richText.map(t => t.text).join('');
    if (v.text !== undefined) return "'" + (typeof v.text === 'string' ? v.text : (v.text.richText || []).map(t => t.text).join(''));
    if (v.error) return '=' + F.enToEs(v.error);
    return null;
  }

  // ---------- Lectura con ExcelJS (.xlsx / .xlsm): datos, fórmulas y formato ----------
  async function readWithExcelJS(buffer) {
    const wb = new root.ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const book = { sheets: [] };
    wb.eachSheet(ws => {
      const sh = newSheet(ws.name);
      ws.eachRow({ includeEmpty: true }, (row, rn) => {
        const r = rn - 1;
        if (row.height && Math.abs(row.height - 15) > 0.5) sh.rowH[r] = Math.round(row.height / 0.75);
        row.eachCell({ includeEmpty: true }, (cell, cn) => {
          const c = cn - 1;
          if (cell.type !== 1) {
            let val = null;
            const formula = cell.formula;
            if (formula) val = '=' + F.enToEs(formula);
            else val = cellValueFromExcel(cell.value);
            if (cell.value instanceof Date && !(cell.numFmt && N.isDateFormat(cell.numFmt))) {
              sh.styles.set(key(r, c), Object.assign(sh.styles.get(key(r, c)) || {}, { fmt: N.FORMATS.fecha }));
            }
            if (val !== null && val !== "'") sh.cells.push([r, c, val]);
          }
          const st = styleFromExcelJS(cell.style || {});
          if (st) sh.styles.set(key(r, c), Object.assign(sh.styles.get(key(r, c)) || {}, st));
        });
      });
      const cols = ws.columnCount;
      for (let c = 1; c <= cols; c++) {
        const col = ws.getColumn(c);
        if (col && col.width) sh.colW[c - 1] = Math.max(16, Math.round(col.width * 7 + 5));
      }
      const merges = (ws.model && ws.model.merges) || [];
      merges.forEach(m => {
        const [a, b] = m.split(':');
        const p1 = parseRef(a), p2 = parseRef(b || a);
        if (p1 && p2) sh.merges.push({ r0: p1.r, c0: p1.c, r1: p2.r, c1: p2.c });
      });
      const view = (ws.views || [])[0];
      if (view && view.state === 'frozen') {
        sh.freezeRow = (view.ySplit || 0) >= 1;
        sh.freezeCol = (view.xSplit || 0) >= 1;
      }
      book.sheets.push(sh);
    });
    return book;
  }

  function styleFromExcelJS(s) {
    const st = {};
    const f = s.font;
    if (f) {
      if (f.bold) st.b = true;
      if (f.italic) st.i = true;
      if (f.underline) st.u = true;
      if (f.size && f.size !== 11) st.size = f.size;
      const col = excelColor(f.color);
      if (col && col !== '#000000') st.color = col;
    }
    if (s.fill && s.fill.type === 'pattern' && s.fill.pattern === 'solid') {
      const col = excelColor(s.fill.fgColor);
      if (col) st.bg = col;
    }
    if (s.border) {
      if (s.border.top && s.border.top.style) st.bt = true;
      if (s.border.right && s.border.right.style) st.br = true;
      if (s.border.bottom && s.border.bottom.style) st.bb = true;
      if (s.border.left && s.border.left.style) st.bl = true;
    }
    if (s.alignment && ['left', 'center', 'right'].includes(s.alignment.horizontal)) st.align = s.alignment.horizontal;
    const fmt = N.normalizeFormat(s.numFmt);
    if (fmt) st.fmt = fmt;
    return Object.keys(st).length ? st : null;
  }

  function parseRef(ref) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref || '');
    return m ? { r: parseInt(m[2], 10) - 1, c: F.colToNum(m[1]) } : null;
  }

  // ---------- Lectura con SheetJS (.xls, .ods y respaldo) ----------
  function readWithSheetJS(buffer) {
    const X = root.XLSX;
    const wb = X.read(buffer, { type: 'array', cellFormula: true, cellNF: true, cellDates: false, bookVBA: false });
    const book = { sheets: [] };
    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name];
      const sh = newSheet(name);
      if (ws['!ref']) {
        const range = X.utils.decode_range(ws['!ref']);
        for (let r = range.s.r; r <= range.e.r; r++) {
          for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = ws[X.utils.encode_cell({ r, c })];
            if (!cell) continue;
            let val = null;
            if (cell.f) val = '=' + F.enToEs(cell.f);
            else if (cell.t === 'n') val = cell.v;
            else if (cell.t === 'b') val = !!cell.v;
            else if (cell.t === 's' || cell.t === 'str') val = cell.v === '' ? null : "'" + cell.v;
            else if (cell.t === 'd') val = serialFromDate(new Date(cell.v));
            else if (cell.t === 'e') val = '=' + F.enToEs(cell.w || '#N/A');
            if (val !== null) sh.cells.push([r, c, val]);
            const fmt = N.normalizeFormat(cell.z);
            if (fmt && (cell.t === 'n' || cell.f)) sh.styles.set(key(r, c), { fmt });
          }
        }
      }
      (ws['!merges'] || []).forEach(m => sh.merges.push({ r0: m.s.r, c0: m.s.c, r1: m.e.r, c1: m.e.c }));
      (ws['!cols'] || []).forEach((col, c) => {
        if (!col) return;
        const px = col.wpx || (col.wch ? col.wch * 7 + 5 : 0);
        if (px) sh.colW[c] = Math.round(px);
      });
      (ws['!rows'] || []).forEach((row, r) => {
        if (row && row.hpx && Math.abs(row.hpx - 20) > 1) sh.rowH[r] = Math.round(row.hpx);
      });
      book.sheets.push(sh);
    });
    return book;
  }

  // ---------- CSV / TSV ----------
  function decodeText(buffer) {
    const bytes = new Uint8Array(buffer);
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
    } catch (e) {
      return new TextDecoder('windows-1252').decode(bytes);
    }
  }

  function detectDelimiter(text) {
    const lines = text.split(/\r?\n/).slice(0, 10).join('\n');
    const count = ch => lines.split(ch).length - 1;
    const cands = ['\t', ';', ','].map(ch => [ch, count(ch)]).sort((a, b) => b[1] - a[1]);
    return cands[0][1] > 0 ? cands[0][0] : ',';
  }

  function parseDelimited(text, delim) {
    const rows = [];
    let row = [], field = '', i = 0, quoted = false;
    while (i < text.length) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          quoted = false; i++; continue;
        }
        field += ch; i++; continue;
      }
      if (ch === '"' && field === '') { quoted = true; i++; continue; }
      if (ch === delim) { row.push(field); field = ''; i++; continue; }
      if (ch === '\r' || ch === '\n') {
        row.push(field); rows.push(row); row = []; field = '';
        if (ch === '\r' && text[i + 1] === '\n') i++;
        i++; continue;
      }
      field += ch; i++;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  function readDelimited(buffer, name, ext) {
    const text = decodeText(buffer);
    const delim = ext === 'tsv' ? '\t' : detectDelimiter(text);
    const rows = parseDelimited(text, delim);
    const sh = newSheet(name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Hoja1');
    rows.forEach((row, r) => row.forEach((txt, c) => {
      const p = N.parseInput(txt);
      if (p.value === null) return;
      sh.cells.push([r, c, p.formula ? "'" + p.value : (p.text ? "'" + p.value : p.value)]);
      if (p.format) sh.styles.set(key(r, c), { fmt: p.format });
    }));
    return { sheets: [sh] };
  }

  async function readFile(file) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    const buffer = await file.arrayBuffer();
    if (ext === 'csv' || ext === 'tsv' || ext === 'txt') return readDelimited(buffer, file.name, ext);
    if (ext === 'xlsx' || ext === 'xlsm') {
      try {
        return await readWithExcelJS(buffer);
      } catch (e) {
        console.warn('ExcelJS no pudo leer el archivo, se usa SheetJS', e);
      }
    }
    return readWithSheetJS(buffer);
  }

  // ---------- Guardado ----------
  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }

  function argb(hex) {
    return 'FF' + hex.replace('#', '').toUpperCase();
  }

  /** src: { sheets: [{ name, rows, cols, cell(r,c) -> {formula, value}, style(r,c), merges, colW, rowH, freezeRow, freezeCol }] } */
  async function writeXlsx(src) {
    const wb = new root.ExcelJS.Workbook();
    wb.creator = 'Coftech Excel';
    src.sheets.forEach(sh => {
      const views = (sh.freezeRow || sh.freezeCol)
        ? [{ state: 'frozen', xSplit: sh.freezeCol ? 1 : 0, ySplit: sh.freezeRow ? 1 : 0 }]
        : [];
      const ws = wb.addWorksheet(sh.name, { views });
      for (let r = 0; r < sh.rows; r++) {
        for (let c = 0; c < sh.cols; c++) {
          const cd = sh.cell(r, c);
          const st = sh.style(r, c);
          if (!cd && !st) continue;
          const cell = ws.getCell(r + 1, c + 1);
          if (cd) {
            if (cd.formula) {
              const v = { formula: F.esToEn(cd.formula.slice(1)) };
              if (cd.value !== null && typeof cd.value !== 'object') v.result = cd.value;
              cell.value = v;
            } else if (cd.value !== null) {
              cell.value = cd.value;
            }
          }
          if (st) applyExcelJSStyle(cell, st);
        }
      }
      Object.keys(sh.colW).forEach(c => { ws.getColumn(+c + 1).width = Math.max(1, (sh.colW[c] - 5) / 7); });
      Object.keys(sh.rowH).forEach(r => { ws.getRow(+r + 1).height = sh.rowH[r] * 0.75; });
      sh.merges.forEach(m => ws.mergeCells(m.r0 + 1, m.c0 + 1, m.r1 + 1, m.c1 + 1));
    });
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function applyExcelJSStyle(cell, st) {
    const font = {};
    if (st.b) font.bold = true;
    if (st.i) font.italic = true;
    if (st.u) font.underline = true;
    if (st.size) font.size = st.size;
    if (st.color) font.color = { argb: argb(st.color) };
    if (Object.keys(font).length) cell.font = Object.assign({ name: 'Calibri', size: 11 }, font);
    if (st.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(st.bg) } };
    const thin = { style: 'thin', color: { argb: 'FF000000' } };
    if (st.bt || st.br || st.bb || st.bl) {
      cell.border = {};
      if (st.bt) cell.border.top = thin;
      if (st.br) cell.border.right = thin;
      if (st.bb) cell.border.bottom = thin;
      if (st.bl) cell.border.left = thin;
    }
    if (st.align) cell.alignment = { horizontal: st.align };
    if (st.fmt) cell.numFmt = st.fmt;
  }

  function writeSheetJS(src, bookType) {
    const X = root.XLSX;
    const wb = X.utils.book_new();
    src.sheets.forEach(sh => {
      const ws = {};
      let maxR = 0, maxC = 0;
      for (let r = 0; r < sh.rows; r++) {
        for (let c = 0; c < sh.cols; c++) {
          const cd = sh.cell(r, c);
          if (!cd) continue;
          const v = cd.value;
          let cell;
          if (v === null || typeof v === 'object') cell = { t: 's', v: '' };
          else if (typeof v === 'number') cell = { t: 'n', v };
          else if (typeof v === 'boolean') cell = { t: 'b', v };
          else cell = { t: 's', v: String(v) };
          if (cd.formula) cell.f = F.esToEn(cd.formula.slice(1));
          const st = sh.style(r, c);
          if (st && st.fmt && cell.t === 'n') cell.z = st.fmt;
          ws[X.utils.encode_cell({ r, c })] = cell;
          maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
        }
      }
      ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
      if (sh.merges.length) ws['!merges'] = sh.merges.map(m => ({ s: { r: m.r0, c: m.c0 }, e: { r: m.r1, c: m.c1 } }));
      const cols = [];
      for (let c = 0; c <= maxC; c++) cols.push({ wpx: sh.colW[c] || DEFAULT_COL_W });
      ws['!cols'] = cols;
      X.utils.book_append_sheet(wb, ws, sh.name);
    });
    const out = X.write(wb, { bookType, type: 'array' });
    const mime = bookType === 'ods' ? 'application/vnd.oasis.opendocument.spreadsheet' : 'application/vnd.ms-excel';
    return new Blob([out], { type: mime });
  }

  /** rows: matriz de textos tal como se ven en pantalla. */
  function writeDelimited(rows, delim) {
    const esc = s => (/["\r\n]/.test(s) || s.includes(delim) ? '"' + s.replace(/"/g, '""') + '"' : s);
    const text = rows.map(row => row.map(esc).join(delim)).join('\r\n');
    const mime = delim === '\t' ? 'text/tab-separated-values' : 'text/csv';
    return new Blob(['﻿' + text], { type: mime + ';charset=utf-8' });
  }

  root.CEIO = {
    readFile, writeXlsx, writeSheetJS, writeDelimited, download, parseDelimited, newSheet,
    DEFAULT_COL_W, DEFAULT_ROW_H
  };
})(typeof window !== 'undefined' ? window : this);
