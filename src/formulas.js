/* Coftech Excel — fórmulas: traducción inglés <-> español y ajuste de referencias. */
(function (root) {
  'use strict';

  const ERR_EN_ES = {
    '#DIV/0!': '#¡DIV/0!', '#N/A': '#N/A', '#NAME?': '#¿NOMBRE?', '#NULL!': '#¡NULO!',
    '#NUM!': '#¡NUM!', '#REF!': '#¡REF!', '#VALUE!': '#¡VALOR!'
  };
  const ERR_ES_EN = {};
  Object.keys(ERR_EN_ES).forEach(k => { ERR_ES_EN[ERR_EN_ES[k]] = k; });

  let EN_ES = {};
  let ES_EN = {};

  /** Arma los diccionarios a partir de los idiomas de HyperFormula (canónico -> nombre). */
  function setLanguages(enFunctions, esFunctions) {
    EN_ES = {}; ES_EN = {};
    Object.keys(esFunctions).forEach(key => {
      const en = (enFunctions[key] || key).toUpperCase();
      const es = esFunctions[key].toUpperCase();
      EN_ES[en] = es;
      ES_EN[es] = en;
    });
    EN_ES.TRUE = 'VERDADERO'; EN_ES.FALSE = 'FALSO';
    ES_EN.VERDADERO = 'TRUE'; ES_EN.FALSO = 'FALSE';
  }

  const isIdStart = ch => /[A-Za-z_À-ſ$]/.test(ch);
  const isIdChar = ch => /[A-Za-z0-9_.À-ſ$]/.test(ch);
  const isDigit = ch => ch >= '0' && ch <= '9';

  /**
   * Recorre la fórmula separando literales de texto, nombres de hoja entre comillas,
   * identificadores, números, errores y símbolos.
   */
  function tokenize(f, dec) {
    dec = dec || ',';
    const out = [];
    let i = 0;
    while (i < f.length) {
      const ch = f[i];
      if (ch === '"' || ch === "'") {
        let j = i + 1;
        while (j < f.length) {
          if (f[j] === ch) {
            if (f[j + 1] === ch) { j += 2; continue; }
            break;
          }
          j++;
        }
        out.push({ t: ch === '"' ? 'str' : 'qsheet', v: f.slice(i, j + 1) });
        i = j + 1;
      } else if (ch === '#') {
        const m = /^#[A-Za-z0-9\/!?¡¿]+[!?]?/.exec(f.slice(i));
        const v = m ? m[0] : '#';
        out.push({ t: 'err', v });
        i += v.length;
      } else if (isIdStart(ch)) {
        let j = i + 1;
        while (j < f.length && isIdChar(f[j])) j++;
        out.push({ t: 'id', v: f.slice(i, j) });
        i = j;
      } else if (isDigit(ch) || (ch === dec && isDigit(f[i + 1] || ''))) {
        let j = i;
        let seenDec = false;
        while (j < f.length && (isDigit(f[j]) || f[j] === dec)) {
          if (f[j] === dec) {
            if (seenDec || !isDigit(f[j + 1] || '')) break;
            seenDec = true;
          }
          j++;
        }
        if (/[eE]/.test(f[j] || '') && /[0-9+\-]/.test(f[j + 1] || '')) {
          j += 2;
          while (j < f.length && isDigit(f[j])) j++;
        }
        out.push({ t: 'num', v: f.slice(i, j) });
        i = j;
      } else {
        out.push({ t: 'sym', v: ch });
        i++;
      }
    }
    return out;
  }

  function nextNonSpace(tokens, k) {
    for (let j = k + 1; j < tokens.length; j++) {
      if (tokens[j].v !== ' ') return tokens[j];
    }
    return null;
  }

  /** Fórmula en inglés (sin "=") -> fórmula en español (sin "="). */
  function enToEs(f) {
    const tokens = tokenize(f, '.');
    return tokens.map((tk, k) => {
      if (tk.t === 'id') {
        const next = nextNonSpace(tokens, k);
        let name = tk.v.replace(/^_xlfn\.|^_xlws\./i, '');
        const up = name.toUpperCase();
        if (next && next.v === '(') return EN_ES[up] || name;
        if ((up === 'TRUE' || up === 'FALSE') && !(next && next.v === '!')) return EN_ES[up];
        return tk.v;
      }
      if (tk.t === 'num') return tk.v.replace(/\./g, ',');
      if (tk.t === 'sym' && tk.v === ',') return ';';
      if (tk.t === 'err') return ERR_EN_ES[tk.v.toUpperCase()] || tk.v;
      return tk.v;
    }).join('');
  }

  /** Fórmula en español (sin "=") -> fórmula en inglés (sin "="). */
  function esToEn(f) {
    const tokens = tokenize(f, ',');
    return tokens.map((tk, k) => {
      if (tk.t === 'id') {
        const next = nextNonSpace(tokens, k);
        const up = tk.v.toUpperCase();
        if (next && next.v === '(') return ES_EN[up] || tk.v;
        if ((up === 'VERDADERO' || up === 'FALSO') && !(next && next.v === '!')) return ES_EN[up];
        return tk.v;
      }
      if (tk.t === 'num') return tk.v.replace(/,/g, '.');
      if (tk.t === 'sym' && tk.v === ';') return ',';
      if (tk.t === 'err') return ERR_ES_EN[tk.v.toUpperCase()] || tk.v;
      return tk.v;
    }).join('');
  }

  const MAX_ROW = 1048576, MAX_COL = 16384;

  function colToNum(s) {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function numToCol(n) {
    let s = '';
    n += 1;
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
  const COL_RE = /^(\$?)([A-Za-z]{1,3})$/;
  const ROW_RE = /^(\$?)(\d+)$/;
  const isRowPart = tk => !!tk && (tk.t === 'num' || tk.t === 'id') && ROW_RE.test(tk.v);

  /**
   * Ajusta las referencias relativas de una fórmula (sin "=") como lo hace Excel
   * al copiar o autocompletar. Las referencias con "$" no se mueven.
   */
  function shift(f, dr, dc) {
    if (!dr && !dc) return f;
    const tokens = tokenize(f, ',');
    const out = [];
    for (let k = 0; k < tokens.length; k++) {
      const tk = tokens[k];
      const next = tokens[k + 1];
      if (tk.t === 'id' && !(next && (next.v === '(' || next.v === '!'))) {
        const m = CELL_RE.exec(tk.v);
        if (m && colToNum(m[2]) < MAX_COL) {
          let c = colToNum(m[2]);
          let r = parseInt(m[4], 10) - 1;
          if (!m[1]) c += dc;
          if (!m[3]) r += dr;
          if (c < 0 || r < 0 || c >= MAX_COL || r >= MAX_ROW) { out.push('#¡REF!'); continue; }
          out.push(m[1] + numToCol(c) + m[3] + (r + 1));
          continue;
        }
        // Columnas completas: A:C
        const cm = COL_RE.exec(tk.v);
        const prev = tokens[k - 1], after = tokens[k + 2];
        const isColRange = cm && ((next && next.v === ':' && after && after.t === 'id' && COL_RE.test(after.v)) ||
          (prev && prev.v === ':' && tokens[k - 2] && tokens[k - 2].t === 'id' && COL_RE.test(tokens[k - 2].v)));
        if (isColRange) {
          let c = colToNum(cm[2]);
          if (!cm[1]) c += dc;
          if (c < 0 || c >= MAX_COL) { out.push('#¡REF!'); continue; }
          out.push(cm[1] + numToCol(c));
          continue;
        }
      }
      // Filas completas: 1:3 o $1:$3
      if (isRowPart(tk)) {
        const prev = tokens[k - 1];
        if ((next && next.v === ':' && isRowPart(tokens[k + 2])) || (prev && prev.v === ':' && isRowPart(tokens[k - 2]))) {
          const m = ROW_RE.exec(tk.v);
          let r = parseInt(m[2], 10) - 1;
          if (!m[1]) r += dr;
          if (r < 0 || r >= MAX_ROW) { out.push('#¡REF!'); continue; }
          out.push(m[1] + (r + 1));
          continue;
        }
      }
      out.push(tk.v);
    }
    return out.join('');
  }

  const api = { setLanguages, enToEs, esToEn, shift, colToNum, numToCol, tokenize };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CEFormulas = api;
})(typeof window !== 'undefined' ? window : this);
