/* Alte Formeln aus dem Formel-Editor 3.0 (Word, „Equation.3“) lesen.
   Kette: OLE-Datei (oleObjectN.bin) → Strom „Equation Native“ → MTEF v3 → LaTeX für den Editor.
   Entspricht den Python-Werkzeugen mtef3.py / tolatex.py / polish.py, mit denen alle 1297 Formeln
   aus Papas Manuskripten fehlerfrei umgewandelt wurden. */
(function (root) {
  'use strict';

  /* ---------- OLE-Verbunddatei (Compound File Binary) ---------- */
  function cfbStream(bytes, wanted) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const u32 = o => dv.getUint32(o, true), u16 = o => dv.getUint16(o, true);
    if (bytes.length < 512 || u32(0) !== 0xE011CFD0 || u32(4) !== 0xE11AB1A1) throw new Error('keine OLE-Datei');
    const ssz = 1 << u16(0x1E), mssz = 1 << u16(0x20);
    const cutoff = u32(0x38);
    const sec = id => 512 + id * ssz + (ssz > 512 ? ssz - 512 : 0);
    // Sektoren der FAT über die DIFAT einsammeln
    const fatSecs = [];
    for (let i = 0; i < 109; i++) { const s = u32(0x4C + i * 4); if (s < 0xFFFFFFFA) fatSecs.push(s); }
    let dif = u32(0x44), ndif = u32(0x48);
    while (ndif-- > 0 && dif < 0xFFFFFFFA) {
      const o = sec(dif), n = ssz / 4 - 1;
      for (let i = 0; i < n; i++) { const s = u32(o + i * 4); if (s < 0xFFFFFFFA) fatSecs.push(s); }
      dif = u32(o + n * 4);
    }
    const fat = [];
    for (const s of fatSecs) { const o = sec(s); for (let i = 0; i < ssz / 4; i++) fat.push(u32(o + i * 4)); }
    const chain = (start, table) => { const out = []; let s = start, guard = 0; while (s < 0xFFFFFFFA && guard++ < 1e6) { out.push(s); s = table[s]; } return out; };
    const readChain = (start, size) => {
      const out = new Uint8Array(size); let p = 0;
      for (const s of chain(start, fat)) { const o = sec(s); const n = Math.min(ssz, size - p); if (n <= 0) break; out.set(bytes.subarray(o, o + n), p); p += n; }
      return out;
    };
    // Verzeichnis
    const entries = [];
    for (const s of chain(u32(0x30), fat)) {
      const o0 = sec(s);
      for (let k = 0; k < ssz / 128; k++) {
        const o = o0 + k * 128, nl = u16(o + 0x40);
        let name = '';
        for (let i = 0; i + 2 < nl; i += 2) name += String.fromCharCode(u16(o + i));
        entries.push({name, type: bytes[o + 0x42], start: u32(o + 0x74), size: u32(o + 0x78)});
      }
    }
    const rootE = entries.find(e => e.type === 5);
    const e = entries.find(x => x.type === 2 && x.name === wanted);
    if (!e) throw new Error('Strom „' + wanted + '“ fehlt');
    if (e.size >= cutoff) return readChain(e.start, e.size);
    // kleine Ströme liegen im Mini-Strom
    const mini = readChain(rootE.start, rootE.size);
    const mfat = [];
    if (u32(0x40)) for (const s of chain(u32(0x3C), fat)) { const o = sec(s); for (let i = 0; i < ssz / 4; i++) mfat.push(u32(o + i * 4)); }
    const out = new Uint8Array(e.size); let p = 0;
    for (const s of chain(e.start, mfat)) { const n = Math.min(mssz, e.size - p); if (n <= 0) break; out.set(mini.subarray(s * mssz, s * mssz + n), p); p += n; }
    return out;
  }

  /* ---------- MTEF v3 ---------- */
  const END = 0, LINE = 1, CHAR = 2, TMPL = 3, PILE = 4, MATRIX = 5, EMBELL = 6, RULER = 7, FONT = 8, SIZE = 9;
  function parseMtef(data) {
    const b = data.subarray(28);           // 28 Byte OLE-Kopf überspringen
    if (b[0] !== 3) throw new Error('kein MTEF v3');
    let i = 5;
    const u8 = () => b[i++];
    const u16 = () => { const v = b[i] | (b[i + 1] << 8); i += 2; return v; };
    const s16 = () => { const v = u16(); return v > 0x7fff ? v - 0x10000 : v; };
    const nudge = () => { const dx = u8() - 128, dy = u8() - 128; if (dx === -128 && dy === -128) { s16(); s16(); } };
    const ruler = () => { const n = u8(); for (let k = 0; k < n; k++) { u8(); u16(); } };
    function objlist() {
      const out = [];
      for (;;) {
        if (i >= b.length) return out;
        const tag = u8(), t = tag & 15, o = tag >> 4;
        if (t === END) return out;
        out.push(rec(t, o));
      }
    }
    function rec(t, o) {
      if (t === LINE) { if (o & 8) nudge(); if (o & 4) u8(); if (o & 2) ruler(); if (o & 1) return ['line', []]; return ['line', objlist()]; }
      if (t === CHAR) { if (o & 8) nudge(); const tf = u8() - 128, c = u16(); if (o & 2) objlist(); return ['char', tf, c]; }
      if (t === TMPL) { if (o & 8) nudge(); const sel = u8(), v = u8(), opt = u8(); return ['tmpl', sel, v, opt, objlist()]; }
      if (t === PILE) { if (o & 8) nudge(); const h = u8(); u8(); if (o & 2) { u8(); ruler(); } return ['pile', h, objlist()]; }
      if (t === MATRIX) { if (o & 8) nudge(); u8(); u8(); u8(); const r = u8(), c = u8(); i += Math.floor(((r + 1) * 2 + 7) / 8) + Math.floor(((c + 1) * 2 + 7) / 8); return ['matrix', r, c, objlist()]; }
      if (t === EMBELL) { if (o & 8) nudge(); return ['embell', u8()]; }
      if (t === RULER) { ruler(); return ['ruler']; }
      if (t === FONT) { const tf = u8(); u8(); let name = ''; for (;;) { const c = u8(); if (!c) break; name += String.fromCharCode(c); } return ['font', tf, name]; }
      if (t === SIZE) { const l = u8(); if (l === 101) s16(); else if (l === 100) { u8(); u16(); } else u8(); return ['size']; }
      if (t >= 10 && t <= 14) return ['sz', t];
      throw new Error('unbekannter Eintrag ' + t);
    }
    return objlist();
  }

  /* ---------- MTEF → LaTeX ---------- */
  const GREEK = {0x3b1: '\\alpha', 0x3b2: '\\beta', 0x3b3: '\\gamma', 0x3b4: '\\delta', 0x3b5: '\\varepsilon', 0x3ba: '\\kappa',
    0x3bb: '\\lambda', 0x3bc: '\\mu', 0x3c0: '\\uppi', 0x3c4: '\\tau', 0x3d5: '\\varphi', 0x3c6: '\\varphi', 0x3c1: '\\rho',
    0x3b6: '\\zeta', 0x3c3: '\\sigma', 0x3c9: '\\omega', 0x3b8: '\\theta', 0x3b7: '\\eta', 0x3bd: '\\nu', 0x3be: '\\xi', 0x3c8: '\\psi', 0x3c7: '\\chi',
    0x3a6: '\\Phi', 0x394: '\\Delta', 0x3a3: '\\Sigma', 0x3a9: '\\Omega', 0x393: '\\Gamma', 0x398: '\\Theta', 0x39b: '\\Lambda', 0x3a8: '\\Psi'};
  const SYM = {0x2212: '-', 0x22c5: '\\cdot', 0xd7: '\\times', 0x2248: '\\approx', 0xb1: '\\pm', 0x2217: '\\ast', 0x2264: '\\le', 0x2265: '\\ge',
    0x221e: '\\infty', 0x2211: '\\sum', 0x2202: '\\partial', 0x2192: '\\to', 0x210f: '\\hbar', 0x19b: '\\lambdabar', 0x22ef: '\\cdots',
    0x2026: '\\ldots', 0x2260: '\\ne', 0x221a: '\\surd', 0x222b: '\\int', 0x2207: '\\nabla', 0xeb04: '\\,', 0xeb02: '\\:', 0x2032: "'",
    0x7b: '\\{', 0x7d: '\\}', 0x25: '\\%'};
  const FUNCS = new Set(['sin', 'cos', 'tan', 'cot', 'ln', 'log', 'exp', 'lim', 'max', 'min', 'sinh', 'cosh', 'tanh', 'arcsin', 'arccos', 'arctan']);
  const isAlpha = c => /^\p{L}$/u.test(String.fromCharCode(c));
  const ch = c => GREEK[c] || SYM[c] || String.fromCharCode(c);

  function toLatex(tree) {
    const lines = objs => objs.filter(n => n[0] === 'line');
    const L = n => runs(n[1]);
    function runs(items) {
      let out = '', i = 0;
      while (i < items.length) {
        const n = items[i];
        if (n[0] === 'char' && (n[1] === 1 || n[1] === 2) && isAlpha(n[2])) {
          const tf = n[1]; let s = '';
          while (i < items.length && items[i][0] === 'char' && items[i][1] === tf && isAlpha(items[i][2])) { s += String.fromCharCode(items[i][2]); i++; }
          out += (tf === 2 && FUNCS.has(s)) ? '\\' + s + ' ' : '\\mathrm{' + s + '}';
          continue;
        }
        out += node(n); i++;
      }
      return out;
    }
    function node(n) {
      const k = n[0];
      if (k === 'sz' || k === 'size' || k === 'font' || k === 'ruler' || k === 'embell') return '';
      if (k === 'line') return L(n);
      if (k === 'char') {
        const tf = n[1], c = n[2];
        if (tf === 7) return '\\mathbf{' + ch(c) + '}';
        if (tf === 1 && c === 32) return '\\ ';
        const s = ch(c);
        return s + (s[0] === '\\' && /[a-zA-Z]$/.test(s) ? ' ' : '');
      }
      if (k === 'tmpl') {
        const sel = n[1], v = n[2], objs = n[4], ls = lines(objs), chars = objs.filter(o => o[0] === 'char');
        const Li = j => ls[j] ? L(ls[j]) : '';
        if (sel === 15) {
          const sub = Li(0), sup = Li(1);
          let s = '';
          if ((v === 1 || v === 2) && sub) s += '_{' + sub + '}';
          if ((v === 0 || v === 2) && sup) s += '^{' + sup + '}';
          return s;
        }
        if (sel === 14 || sel === 255) return '\\frac{' + Li(0) + '}{' + Li(1) + '}';
        if (sel === 13) { const idx = Li(1); return '\\sqrt' + (idx ? '[' + idx + ']' : '') + '{' + Li(0) + '}'; }
        const FEN = {0: ['\\langle', '\\rangle'], 1: ['(', ')'], 2: ['\\{', '\\}'], 3: ['[', ']'], 4: ['|', '|'], 5: ['\\|', '\\|']};
        if (FEN[sel]) {
          let [l, r] = FEN[sel];
          if (v === 1) r = '.';
          if (v === 2) l = '.';
          return '\\left' + l + ' ' + Li(0) + ' \\right' + r;
        }
        if (sel === 29) {
          let s = chars.length ? ch(chars[chars.length - 1][2]) : '\\sum';
          if (Li(1)) s += '_{' + Li(1) + '}';
          if (Li(2)) s += '^{' + Li(2) + '}';
          return s + ' ' + Li(0);
        }
        return '{' + objs.map(node).join('') + '}';
      }
      if (k === 'pile') return lines(n[2]).map(L).join('\\\\');
      return '';
    }
    return tree.map(node).join('').trim();
  }

  // Beschreibende Indizes aufrecht, Einheit s nach Zehnerpotenz aufrecht (wie polish.py)
  const WORDS = 'Pl|mag|el|weak|strong|proton|electron|pr|eff|total|tot|rot|point|kin|pot|max|min';
  function polish(t) {
    t = t.replace(/_\{((?:[^{}]|\{[^{}]*\})*)\}/g, (m, inner) =>
      '_{' + inner.replace(new RegExp('(?<![\\\\A-Za-z])(' + WORDS + ')(?![A-Za-z])', 'g'), '\\mathrm{$1}') + '}');
    t = t.replace(/(10\^\{-?\d+\})\s*s\b/g, '$1\\,\\mathrm{s}');
    return t;
  }

  function oleToLatex(bytes) { return polish(toLatex(parseMtef(cfbStream(bytes, 'Equation Native')))); }

  const api = {cfbStream, parseMtef, toLatex, polish, oleToLatex};
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Eqn3 = api;
})(typeof self !== 'undefined' ? self : this);
