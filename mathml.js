/* Formeln aus dem heutigen MathType (Add-in für Word 365) lesen. MathType fügt Bilder ein und legt die Formel
   als MathML in den Alternativtext: {"mathml":"<math …>…</math>","origin":"MathType Legacy",…}.
   Kette: MathML → Baum → LaTeX für den Editor. Nutzt das XML-Lesen und die Zeichentabellen aus omml.js. */
(function (root) {
  'use strict';
  const O = root.Omml || (typeof require === 'function' ? require('./omml.js') : null);

  // MathML aus dem Alternativtext eines Bildes, sonst null
  function fromAlt(alt) {
    const s = String(alt || '').trim();
    if (!/^\{[\s\S]*"mathml"/.test(s)) return null;
    try { const j = JSON.parse(s); return typeof j.mathml === 'string' && /<math[\s>]/.test(j.mathml) ? j.mathml : null; } catch (e) { return null; }
  }
  // Schriftgröße aus style="font-size:12px" (MathType meint damit Punkt)
  function sizeOf(mathml) {
    const m = String(mathml).match(/font-size\s*:\s*([\d.]+)\s*(px|pt)/);
    return m ? +m[1] : null;
  }

  const local = n => n.name ? n.name.replace(/^.*:/, '') : '';
  const els = n => n.kids.filter(k => k.name);
  const textIn = n => n.text !== undefined ? n.text : n.kids.map(textIn).join('');
  const FENCE = {'(': '(', ')': ')', '[': '[', ']': ']', '{': '\\{', '}': '\\}', '|': '|', '‖': '\\|', '⟨': '\\langle', '⟩': '\\rangle',
    '〈': '\\langle', '〉': '\\rangle', '⌊': '\\lfloor', '⌋': '\\rfloor', '⌈': '\\lceil', '⌉': '\\rceil', '': '.'};
  const fence = c => FENCE[c] !== undefined ? FENCE[c] : c;
  const OVER = {'¯': 'overline', '‾': 'overline', '_': 'overline', '→': 'vec', '⃗': 'vec', '^': 'hat', 'ˆ': 'hat', '˙': 'dot', '.': 'dot',
    '¨': 'ddot', '~': 'tilde', '˜': 'tilde', 'ˇ': 'check', '˘': 'breve', '⏞': 'overbrace', '←': 'overleftarrow', '↔': 'overleftrightarrow'};
  const UNDER = {'¯': 'underline', '_': 'underline', '‾': 'underline', '⏟': 'underbrace'};
  const BIG = /^(\\sum|\\prod|\\coprod|\\int|\\iint|\\iiint|\\oint|\\bigcup|\\bigcap|\\lim|\\max|\\min)\s*$/;
  // Grundzeichen für Hoch- und Tiefstellung: ein Zeichen oder Befehl bleibt so, sonst in Klammern
  const base = s => /^(\\[A-Za-z]+\s*|[A-Za-z0-9]|\{[\s\S]*\})$/.test(s.trim()) ? s.trim() : '{' + s + '}';
  const grp = s => '{' + s.trim() + '}';

  function toLatex(mathml) {
    const tree = O.parse(String(mathml));
    const math = (function find(n) { if (local(n) === 'math') return n; for (const k of els(n)) { const f = find(k); if (f) return f; } return null; })(tree);
    if (!math) throw new Error('keine Formel');
    return row(els(math)).replace(/\s+$/, '');
  }
  // Mehrere Knoten hintereinander; Klammer-Paare aus mo werden \left…\right, wenn sie hohe Inhalte umschließen
  function row(list) {
    const out = [];
    for (let i = 0; i < list.length; i++) {
      const n = list[i];
      // ( Tabelle ) → Matrix mit passenden Klammern
      if (local(n) === 'mo' && list[i + 1] && local(list[i + 1]) === 'mtable' && list[i + 2] && local(list[i + 2]) === 'mo') {
        const env = {'()': 'pmatrix', '[]': 'bmatrix', '||': 'vmatrix', '{}': 'Bmatrix', '‖‖': 'Vmatrix'}[textIn(n).trim() + textIn(list[i + 2]).trim()];
        if (env) { out.push(table(list[i + 1], env)); i += 2; continue; }
      }
      if (local(n) === 'mo' && textIn(n).trim() === '{' && list[i + 1] && local(list[i + 1]) === 'mtable' && i + 2 === list.length) {
        out.push(table(list[i + 1], 'cases')); i += 1; continue;
      }
      out.push(node(n));
    }
    return out.join('');
  }
  function table(t, env) {
    const rows = els(t).filter(r => /^m(labeled)?tr$/.test(local(r)));
    const body = rows.map(r => els(r).filter(c => local(c) === 'mtd').map(c => row(els(c)).trim()).join('&')).join('\\\\');
    if (env === 'cases') return '\\begin{cases}' + body + '\\end{cases}';
    // Ausgerichtete Gleichungen: Tabelle mit columnalign right left
    if (!env && /right\s+left/.test(t.attrs.columnalign || '')) return '\\begin{aligned}' + body + '\\end{aligned}';
    return '\\begin{' + (env || 'matrix') + '}' + body + '\\end{' + (env || 'matrix') + '}';
  }
  function chars(t, variant) {
    if (!t) return '';
    const tt = t.replace(/⁢|⁡|⁣|⁤/g, '');      // unsichtbares Mal, Funktionsanwendung
    return O.runText(tt, variant === 'normal' ? 'p' : variant === 'bold' ? 'b' : '', false);
  }
  function node(n) {
    if (!n.name) return '';
    const k = local(n), c = els(n), a = n.attrs || {};
    const arg = i => c[i] ? node(c[i]) : '';
    switch (k) {
      case 'mrow': case 'mstyle': case 'mpadded': case 'semantics': case 'menclose': case 'math': {
        if (k === 'semantics') {
          const tex = c.find(x => local(x) === 'annotation' && /tex/i.test(x.attrs.encoding || ''));
          if (tex) return textIn(tex).trim();
          return c[0] ? node(c[0]) : '';
        }
        const inner = row(c.filter(x => !/^annotation/.test(local(x))));
        const col = a.mathcolor && /^#?[0-9a-f]{6}$/i.test(a.mathcolor) && !/^#?0{6}$/.test(a.mathcolor) ? a.mathcolor.replace(/^#?/, '#').toUpperCase() : null;
        return col ? '\\textcolor{' + col + '}{' + inner + '}' : (k === 'mrow' && c.length > 1 ? inner : inner);
      }
      case 'mi': {
        const t = textIn(n).trim();
        if (t.length > 1 && /^[A-Za-z]+$/.test(t)) return O.runText(t, 'p', false);   // sin, cos, Wörter
        if (t === 'π') return '\\uppi ';                                              // Kreiszahl aufrecht, wie im übrigen Editor
        return chars(t, a.mathvariant);
      }
      case 'mn': return chars(textIn(n).trim(), a.mathvariant);
      case 'mo': {
        const t = textIn(n).trim();
        if (a.stretchy === 'true' && FENCE[t] !== undefined) return fence(t);
        return chars(t);
      }
      case 'mtext': {
        const t = textIn(n);
        if (!t.trim()) return t.length ? '\\ ' : '';
        return O.runText(t, '', true);
      }
      case 'ms': return O.runText('"' + textIn(n) + '"', '', true);
      case 'mspace': {
        const w = parseFloat(a.width || '0');
        return w >= 0.9 ? '\\quad ' : w >= 0.25 ? '\\ ' : w > 0.1 ? '\\, ' : '';
      }
      case 'msub': return base(arg(0)) + '_' + grp(arg(1));
      case 'msup': return base(arg(0)) + '^' + grp(arg(1));
      case 'msubsup': return base(arg(0)) + '_' + grp(arg(1)) + '^' + grp(arg(2));
      case 'mfrac': {
        if (/^0(\D|$)/.test(a.linethickness || '')) return '{' + arg(0) + '\\atop ' + arg(1) + '}';
        return '\\frac{' + arg(0) + '}{' + arg(1) + '}';
      }
      case 'msqrt': return '\\sqrt{' + row(c) + '}';
      case 'mroot': return '\\sqrt[' + arg(1) + ']{' + arg(0) + '}';
      case 'mfenced': {
        const open = a.open !== undefined ? a.open : '(', close = a.close !== undefined ? a.close : ')';
        const sep = a.separators !== undefined ? a.separators.trim() : ',';
        const inner = c.map(node).join(sep ? chars(sep[0]) : '');
        return '\\left' + fence(open) + inner + '\\right' + fence(close);
      }
      case 'mover': case 'munder': case 'munderover': {
        const b = arg(0);
        const accent = s => { const t = textIn(s).trim(); return t; };
        if (k === 'mover' && c[1] && OVER[accent(c[1])]) return '\\' + OVER[accent(c[1])] + '{' + b + '}';
        if (k === 'munder' && c[1] && UNDER[accent(c[1])]) return '\\' + UNDER[accent(c[1])] + '{' + b + '}';
        const big = BIG.test(b);
        const lim = big ? '\\limits' : '';
        if (k === 'mover') return big ? b.trim() + lim + '^' + grp(arg(1)) : '\\overset{' + arg(1) + '}{' + b + '}';
        if (k === 'munder') return big ? b.trim() + lim + '_' + grp(arg(1)) : '\\underset{' + arg(1) + '}{' + b + '}';
        return big ? b.trim() + lim + '_' + grp(arg(1)) + '^' + grp(arg(2)) : '\\overset{' + arg(2) + '}{\\underset{' + arg(1) + '}{' + b + '}}';
      }
      case 'mmultiscripts': {
        // Grundzeichen, dann Paare (tief, hoch); nach <mprescripts/> die vorgestellten Paare
        const i = c.findIndex(x => local(x) === 'mprescripts');
        const post = (i < 0 ? c.slice(1) : c.slice(1, i)), pre = i < 0 ? [] : c.slice(i + 1);
        const pair = (l, j) => {
          const sub = l[j] && local(l[j]) !== 'none' ? node(l[j]) : '', sup = l[j + 1] && local(l[j + 1]) !== 'none' ? node(l[j + 1]) : '';
          return (sub ? '_' + grp(sub) : '') + (sup ? '^' + grp(sup) : '');
        };
        let s = '';
        for (let j = 0; j < pre.length; j += 2) s += '{}' + pair(pre, j);
        s += base(arg(0));
        for (let j = 0; j < post.length; j += 2) s += pair(post, j);
        return s;
      }
      case 'mtable': return table(n, null);
      case 'mphantom': return '\\phantom{' + row(c) + '}';
      case 'none': case 'mprescripts': case 'annotation': case 'annotation-xml': return '';
      case 'merror': return '';
      default: return row(c);
    }
  }

  const api = {fromAlt, toLatex, sizeOf};
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.MathML = api;
})(typeof self !== 'undefined' ? self : this);
