/* Coftech Excel — interpretación de lo que se escribe y formato de números (estilo argentino). */
(function (root) {
  'use strict';

  const FORMATS = {
    general: 'General',
    numero: '#,##0.00',
    moneda: '"$" #,##0.00',
    porcentaje: '0%',
    fecha: 'dd/mm/yyyy'
  };

  // Excel guarda la "fecha corta" del sistema con códigos en inglés; en Argentina se ve dd/mm/aaaa.
  const DATE_ALIASES = { 'm/d/yy': 'dd/mm/yyyy', 'mm-dd-yy': 'dd/mm/yyyy', 'm/d/yyyy': 'dd/mm/yyyy', 'd/m/yyyy': 'dd/mm/yyyy' };

  function normalizeFormat(fmt) {
    if (!fmt || /^general$/i.test(fmt)) return '';
    return DATE_ALIASES[fmt.toLowerCase()] || fmt;
  }

  function isDateFormat(fmt) {
    if (!fmt) return false;
    const clean = fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
    return /[dy]/i.test(clean) || /m{1,4}.*[d\/y]|[d\/].*m/i.test(clean);
  }

  /** Número de serie de Excel a partir de día/mes/año. */
  function dateSerial(y, m, d) {
    return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  }

  /** Convierte un número escrito al estilo argentino (1.234,56) o internacional (1234.56). */
  function parseNumber(s) {
    s = s.trim();
    if (!/^[+-]?[\d.,]*\d[\d.,]*$/.test(s)) return null;
    let t = s;
    if (t.includes(',')) {
      if ((t.match(/,/g) || []).length > 1) return null;
      t = t.replace(/\./g, '').replace(',', '.');
    } else if (/^[+-]?\d{1,3}(\.\d{3})+$/.test(t)) {
      t = t.replace(/\./g, '');
    } else if ((t.match(/\./g) || []).length > 1) {
      return null;
    }
    const n = Number(t);
    return isFinite(n) ? n : null;
  }

  /**
   * Interpreta lo que el usuario escribe en una celda.
   * Devuelve { value, format } donde value es número, booleano, texto o fórmula ("=...").
   */
  function parseInput(text) {
    if (text == null) return { value: null };
    const raw = String(text);
    const s = raw.trim();
    if (s === '') return { value: null };
    if (s[0] === '=' && s.length > 1) return { value: s, formula: true };
    if (raw[0] === "'") return { value: raw.slice(1), text: true };

    let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(s);
    if (m) {
      const d = +m[1], mo = +m[2];
      let y = +m[3];
      if (m[3].length === 2) y += y < 50 ? 2000 : 1900;
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
        const dt = new Date(Date.UTC(y, mo - 1, d));
        if (dt.getUTCDate() === d) return { value: dateSerial(y, mo, d), format: FORMATS.fecha };
      }
    }

    m = /^([+-]?)\s*\$\s*(.+)$/.exec(s);
    if (m) {
      const n = parseNumber(m[2]);
      if (n !== null) return { value: m[1] === '-' ? -n : n, format: FORMATS.moneda };
    }

    m = /^(.+?)\s*%$/.exec(s);
    if (m) {
      const n = parseNumber(m[1]);
      if (n !== null) {
        const dec = (m[1].split(',')[1] || '').length;
        return { value: n / 100, format: dec ? '0.' + '0'.repeat(dec) + '%' : '0%' };
      }
    }

    const n = parseNumber(s);
    if (n !== null) {
      // "1.500" se interpreta como mil quinientos; se le pone separador de miles para que se vea igual.
      const fmt = /^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s) ? '#,##0' + (s.includes(',') ? '.' + '0'.repeat(s.split(',')[1].length) : '') : undefined;
      return { value: n, format: fmt };
    }

    const up = s.toUpperCase();
    if (up === 'VERDADERO') return { value: true };
    if (up === 'FALSO') return { value: false };
    return { value: raw, text: true };
  }

  const swapSeparators = s => s.replace(/[.,]/g, c => (c === '.' ? ',' : '.'));

  /** Número sin formato: hasta 10 cifras significativas, con coma decimal. */
  function formatGeneral(n) {
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    let s = Math.abs(n) >= 1e11 || (Math.abs(n) < 1e-9 && n !== 0) ? n.toExponential(5) : String(Number(n.toPrecision(10)));
    return s.replace('.', ',');
  }

  /** Texto que se ve en la celda. */
  function display(value, fmt, ssf) {
    if (value === null || value === undefined || value === '') return '';
    if (typeof value === 'boolean') return value ? 'VERDADERO' : 'FALSO';
    if (typeof value === 'number') {
      fmt = normalizeFormat(fmt);
      if (!fmt) return formatGeneral(value);
      try {
        const out = ssf.format(fmt, value);
        return isDateFormat(fmt) ? out : swapSeparators(out);
      } catch (e) {
        return formatGeneral(value);
      }
    }
    return String(value);
  }

  /** Texto que se ve en la barra de fórmulas al editar un valor. */
  function editText(value, fmt, ssf) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'boolean') return value ? 'VERDADERO' : 'FALSO';
    if (typeof value === 'number') {
      fmt = normalizeFormat(fmt);
      if (fmt && isDateFormat(fmt)) return ssf.format('dd/mm/yyyy', value);
      if (fmt && /%/.test(fmt)) return formatGeneral(Number((value * 100).toPrecision(12))) + '%';
      return formatGeneral(value);
    }
    return String(value);
  }

  /** Cambia la cantidad de decimales de un formato (delta = +1 o -1). */
  function changeDecimals(fmt, delta, sample) {
    fmt = normalizeFormat(fmt);
    if (!fmt || isDateFormat(fmt)) {
      if (isDateFormat(fmt)) return fmt;
      let dec = 0;
      if (typeof sample === 'number') {
        const g = formatGeneral(sample);
        dec = g.includes(',') ? g.split(',')[1].length : 0;
      }
      dec = Math.max(0, dec + delta);
      return dec ? '0.' + '0'.repeat(dec) : '0';
    }
    return fmt.split(';').map(section => {
      const m = /0(\.0+)?/.exec(section);
      if (!m) return section;
      const dec = m[1] ? m[1].length - 1 : 0;
      const nd = Math.max(0, dec + delta);
      const repl = nd ? '0.' + '0'.repeat(nd) : '0';
      return section.slice(0, m.index) + repl + section.slice(m.index + m[0].length);
    }).join(';');
  }

  const api = { FORMATS, parseInput, parseNumber, display, editText, changeDecimals, normalizeFormat, isDateFormat, dateSerial, formatGeneral };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CEFormat = api;
})(typeof window !== 'undefined' ? window : this);
