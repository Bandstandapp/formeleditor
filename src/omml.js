/* Formeln aus dem heutigen Formel-Editor von Word (Office Math, „OMML“, <m:oMath>) lesen.
   Kette: OMML-XML → Baum → LaTeX für den Editor. Ohne DOMParser, damit es auch im Prüflauf (Node) läuft. */
(function (root) {
  'use strict';

  /* ---------- kleines XML-Lesen ---------- */
  const unxml = s => s.replace(/&#x([0-9a-fA-F]+);/g, (m, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(+d))
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  function parse(xml) {
    const top = {name: '#root', attrs: {}, kids: []}, stack = [top];
    const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<(\/?)([A-Za-z_][\w.:-]*)((?:\s+[\w.:-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
    let m;
    while ((m = re.exec(xml))) {
      const cur = stack[stack.length - 1];
      if (m[5] !== undefined) { cur.kids.push({text: unxml(m[5])}); continue; }
      if (!m[2]) continue;
      if (m[1]) { if (stack.length > 1) stack.pop(); continue; }
      const attrs = {};
      (m[3] || '').replace(/([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g, (x, k, a, b) => { attrs[k] = unxml(a !== undefined ? a : b); });
      const el = {name: m[2], attrs, kids: []};
      cur.kids.push(el);
      if (!m[4]) stack.push(el);
    }
    return top;
  }
  const local = n => n.name ? n.name.replace(/^.*:/, '') : '';
  const kids = (n, name) => n.kids.filter(k => k.name && (!name || local(k) === name));
  const kid = (n, name) => (n ? n.kids.find(k => k.name && local(k) === name) : null) || null;
  const val = (n, name) => { const k = kid(n, name); if (!k) return null; const v = Object.keys(k.attrs).find(a => a.replace(/^.*:/, '') === 'val'); return v ? k.attrs[v] : ''; };
  const on = (n, name) => { const v = val(n, name); return v !== null && !/^(0|off|false)$/.test(v); };

  /* ---------- Zeichen ---------- */
  const GREEK = {α: 'alpha', β: 'beta', γ: 'gamma', δ: 'delta', ε: 'varepsilon', ϵ: 'epsilon', ζ: 'zeta', η: 'eta', θ: 'theta', ϑ: 'vartheta',
    ι: 'iota', κ: 'kappa', λ: 'lambda', μ: 'mu', ν: 'nu', ξ: 'xi', π: 'pi', ϖ: 'varpi', ρ: 'rho', ϱ: 'varrho', σ: 'sigma', ς: 'varsigma',
    τ: 'tau', υ: 'upsilon', φ: 'varphi', ϕ: 'phi', χ: 'chi', ψ: 'psi', ω: 'omega',
    Γ: 'Gamma', Δ: 'Delta', Θ: 'Theta', Λ: 'Lambda', Ξ: 'Xi', Π: 'Pi', Σ: 'Sigma', Υ: 'Upsilon', Φ: 'Phi', Ψ: 'Psi', Ω: 'Omega'};
  const SYM = {'−': '-', '∙': '\\cdot', '⋅': '\\cdot', '·': '\\cdot', '×': '\\times', '÷': '\\div', '±': '\\pm', '∓': '\\mp', '∗': '\\ast',
    '≈': '\\approx', '≠': '\\ne', '≡': '\\equiv', '≤': '\\le', '≥': '\\ge', '≪': '\\ll', '≫': '\\gg', '∼': '\\sim', '≃': '\\simeq', '≅': '\\cong',
    '∝': '\\propto', '∞': '\\infty', '∂': '\\partial', '∇': '\\nabla', '→': '\\to', '←': '\\leftarrow', '↔': '\\leftrightarrow',
    '⇒': '\\Rightarrow', '⇐': '\\Leftarrow', '⇔': '\\Leftrightarrow', '↦': '\\mapsto', '∈': '\\in', '∉': '\\notin', '⊂': '\\subset',
    '⊆': '\\subseteq', '∪': '\\cup', '∩': '\\cap', '∅': '\\emptyset', '∀': '\\forall', '∃': '\\exists', '¬': '\\neg', '∧': '\\wedge',
    '∨': '\\vee', '⊥': '\\perp', '∥': '\\parallel', '∠': '\\angle', '°': '^{\\circ}', '′': "'", '″': "''", '‴': "'''",
    '⋯': '\\cdots', '…': '\\ldots', '⋮': '\\vdots', '⋱': '\\ddots', 'ℏ': '\\hbar', 'ƛ': '\\lambdabar', 'ℓ': '\\ell', 'ℜ': '\\Re', 'ℑ': '\\Im',
    'ℵ': '\\aleph', '∑': '\\sum', '∏': '\\prod', '∐': '\\coprod', '∫': '\\int', '∬': '\\iint', '∭': '\\iiint', '∮': '\\oint',
    '⟨': '\\langle', '⟩': '\\rangle', '〈': '\\langle', '〉': '\\rangle', '⊗': '\\otimes', '⊕': '\\oplus', '∘': '\\circ', '†': '\\dagger',
    '{': '\\{', '}': '\\}', '%': '\\%', '#': '\\#', '$': '\\$', '_': '\\_', '‖': '\\|', '\u00a0': '\\ ', '\u2009': '\\,', '\u2005': '\\:',
    '\u2004': '\\;', '\u2003': '\\quad', '\u2002': '\\enspace', '\u200b': ''};
  const FUNCS = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'sinh', 'cosh', 'tanh', 'coth', 'arcsin', 'arccos', 'arctan',
    'ln', 'log', 'lg', 'exp', 'lim', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'ker', 'arg', 'deg', 'gcd', 'Pr', 'liminf', 'limsup']);
  const cmd = c => '\\' + c + ' ';
  const isLatin = c => /^[A-Za-z]$/.test(c);

  // Text eines Laufs: sty = p (aufrecht), b, i, bi; nor = normaler Text
  function runText(t, sty, nor) {
    if (nor) return '\\text{' + t.replace(/[\\{}$&#^_%~]/g, c => '\\' + c) + '}';
    const chars = Array.from(t);
    let out = '';
    for (let i = 0; i < chars.length; i++) {
      const c = chars[i];
      if (isLatin(c)) {
        let w = c;
        while (i + 1 < chars.length && isLatin(chars[i + 1])) w += chars[++i];
        if (sty === 'p') out += FUNCS.has(w) ? cmd(w) : '\\mathrm{' + w + '}';
        else out += w;
        continue;
      }
      if (c === ',' && /\d/.test(chars[i - 1] || '') && /^\d\d/.test(chars.slice(i + 1, i + 3).join(''))) { out += '{,}'; continue; }   // 0,636 als Dezimalzahl (R₄,₃ bleibt)
      if (c === ' ') { out += '\\ '; continue; }
      if (GREEK[c]) {
        // Aufrechte kleine griechische Buchstaben (π als Kreiszahl) wie im übrigen Editor
        out += (sty === 'p' && c === 'π') ? '\\uppi ' : (sty === 'p' && /[α-ω]/.test(c)) ? '\\mathrm{' + cmd(GREEK[c]).trim() + '}' : cmd(GREEK[c]);
        continue;
      }
      if (SYM[c] !== undefined) { const s = SYM[c]; out += s + (/\\[A-Za-z]+$/.test(s) ? ' ' : ''); continue; }
      if (/[\u0300-\u036f]/.test(c)) continue;                     // lose Akzentzeichen
      if (/[\uE000-\uF8FF]/.test(c)) continue;                      // Zeichen aus privaten Schriftbereichen
      out += c;
    }
    return out;
  }

  // Schwarz und fast Schwarz (z. B. aus Webseiten kopiert) bleiben ohne Farbe
  const COLOR_SKIP = /^(auto|[0-3][0-9A-F][0-3][0-9A-F][0-3][0-9A-F])$/i;

  /* ---------- OMML → LaTeX ---------- */
  function toLatex(xml) {
    const tree = typeof xml === 'string' ? parse(xml) : xml;
    // Inhalt eines Arguments (e, num, sub …): seine Kinder hintereinander
    const arg = n => n ? n.kids.map(node).join('') : '';
    // Grundzeichen von Hoch- und Tiefstellung. Word hängt den Index an die ganze Gruppe (½m₀ als (½m)₀);
    // gemeint und gleich aussehend ist er am letzten Zeichen, so bleibt die Formel gut bearbeitbar.
    const group = s => {
      const t = s.trim();
      if (/^(\\[A-Za-z]+\s*|[A-Za-z0-9]|\\.|\{(?:[^{}]|\{[^{}]*\})*\})$/.test(t)) return t;
      if (/^\\(mathrm|mathbf|boldsymbol|text|textcolor|left)/.test(t) && balanced(t)) return t;
      const last = t.match(/(\\[A-Za-z]+ ?|[A-Za-z0-9])$/);
      if (last && t.length > last[0].length && !/\\$/.test(t.slice(0, -last[0].length)) && !/^\\(left|right)/.test(last[0])) return t;
      return '{' + t + '}';
    };
    const balanced = t => {                                            // ein einziger Befehl mit Klammer-Argumenten?
      const m = t.match(/^\\[A-Za-z]+/); if (!m) return false;
      let i = m[0].length, d = 0;
      if (m[0] === '\\left') return /\\right.?$/.test(t) && t.lastIndexOf('\\left') === 0;
      for (; i < t.length; i++) { const c = t[i]; if (c === '{') d++; else if (c === '}') d--; else if (!d && c !== ' ') return false; }
      return d === 0;
    };
    const FENCE = {'(': '(', ')': ')', '[': '[', ']': ']', '{': '\\{', '}': '\\}', '|': '|', '‖': '\\|', '⟨': '\\langle', '⟩': '\\rangle',
      '〈': '\\langle', '〉': '\\rangle', '⌊': '\\lfloor', '⌋': '\\rfloor', '⌈': '\\lceil', '⌉': '\\rceil', '': '.'};
    const fence = c => FENCE[c] !== undefined ? FENCE[c] : (c || '.');
    const ACC = {'\u0302': 'hat', '\u0303': 'tilde', '\u0307': 'dot', '\u0308': 'ddot', '\u20d7': 'vec', '\u20d1': 'vec', '\u0304': 'bar',
      '\u0305': 'overline', '\u00af': 'overline', '\u030c': 'check', '\u0306': 'breve', '\u0301': 'acute', '\u0300': 'grave', '\u20d6': 'overleftarrow',
      '\u20e1': 'overleftrightarrow', '\u02c6': 'hat', '\u02dc': 'tilde', '~': 'tilde', '^': 'hat'};
    function rows(n, name) { return kids(n, name); }

    function node(n) {
      if (!n.name) return '';
      const k = local(n);
      switch (k) {
        case 'r': {
          const rPr = n.kids.find(x => x.name === 'm:rPr');
          const wPr = n.kids.find(x => x.name === 'w:rPr');
          const s = rPr ? val(rPr, 'sty') : null;
          const nor = rPr ? on(rPr, 'nor') : false;
          let t = kids(n, 't').map(x => x.kids.map(c => c.text || '').join('')).join('');
          let out = runText(t, s, nor);
          const blank = /^(\\ |\s)*$/.test(out);
          if (s === 'b' && !blank) out = '\\mathbf{' + out + '}';
          else if (s === 'bi' && !blank) out = '\\boldsymbol{' + out + '}';
          const col = wPr ? val(wPr, 'color') : null;
          if (col && !COLOR_SKIP.test(col) && /^[0-9A-Fa-f]{6}$/.test(col) && !blank) out = '\\textcolor{#' + col.toUpperCase() + '}{' + out + '}';
          return out;
        }
        case 'f': {
          const type = val(kid(n, 'fPr'), 'type');
          const a = arg(kid(n, 'num')), b = arg(kid(n, 'den'));
          if (type === 'lin') return group(a) + '/' + group(b);
          if (type === 'noBar') return '\\genfrac{}{}{0pt}{}{' + a + '}{' + b + '}';
          return '\\frac{' + a + '}{' + b + '}';
        }
        case 'sSub': return group(arg(kid(n, 'e'))) + '_{' + arg(kid(n, 'sub')) + '}';
        case 'sSup': return group(arg(kid(n, 'e'))) + '^{' + arg(kid(n, 'sup')) + '}';
        case 'sSubSup': return group(arg(kid(n, 'e'))) + '_{' + arg(kid(n, 'sub')) + '}^{' + arg(kid(n, 'sup')) + '}';
        case 'sPre': return '{}_{' + arg(kid(n, 'sub')) + '}^{' + arg(kid(n, 'sup')) + '}' + group(arg(kid(n, 'e')));
        case 'rad': {
          const pr = kid(n, 'radPr'), deg = arg(kid(n, 'deg'));
          return '\\sqrt' + (!on(pr, 'degHide') && deg.trim() ? '[' + deg + ']' : '') + '{' + arg(kid(n, 'e')) + '}';
        }
        case 'nary': {
          const pr = kid(n, 'naryPr');
          const c = val(pr, 'chr') || '∫';
          let s = SYM[c] || c;
          const isInt = /int$/.test(s);
          if (isInt && val(pr, 'limLoc') === 'undOvr') s += '\\limits';
          const sub = on(pr, 'subHide') ? '' : arg(kid(n, 'sub')), sup = on(pr, 'supHide') ? '' : arg(kid(n, 'sup'));
          if (sub.trim()) s += '_{' + sub + '}';
          if (sup.trim()) s += '^{' + sup + '}';
          return s + ' ' + arg(kid(n, 'e'));
        }
        case 'd': {
          const pr = kid(n, 'dPr');
          const b = val(pr, 'begChr'), e = val(pr, 'endChr'), sep = val(pr, 'sepChr');
          const l = fence(b === null ? '(' : b), r = fence(e === null ? ')' : e);
          const s = sep === null ? '|' : sep;
          const inner = rows(n, 'e').map(arg).join(s === '|' ? '\\mid ' : (SYM[s] || s));
          return '\\left' + l + (/[A-Za-z]$/.test(l) ? ' ' : '') + inner + '\\right' + r + (/[A-Za-z]$/.test(r) ? ' ' : '');
        }
        case 'func': {
          const name = arg(kid(n, 'fName'));
          const plain = name.replace(/\\mathrm\{([A-Za-z]+)\}/g, '$1').replace(/\\([A-Za-z]+)\s*/g, '$1').trim();
          let f;
          if (FUNCS.has(plain)) f = '\\' + plain + ' ';
          else if (/^\\(lim|max|min|sup|inf)(_|\\limits)/.test(name.trim())) f = name.trim() + ' ';
          else if (/^[A-Za-z]+$/.test(plain)) f = '\\operatorname{' + plain + '}';
          else f = name;
          return f + arg(kid(n, 'e'));
        }
        case 'limLow': {
          const e = arg(kid(n, 'e')).trim(), lim = arg(kid(n, 'lim'));
          if (/^\\(lim|max|min|sup|inf)\s*$/.test(e) || /^\\mathrm\{(lim|max|min)\}$/.test(e)) return '\\' + e.replace(/\\mathrm\{|\}|\\|\s/g, '') + '_{' + lim + '}';
          return '\\underset{' + lim + '}{' + e + '}';
        }
        case 'limUpp': return '\\overset{' + arg(kid(n, 'lim')) + '}{' + arg(kid(n, 'e')) + '}';
        case 'acc': {
          const c = val(kid(n, 'accPr'), 'chr');
          const a = ACC[c === null ? '\u0302' : c] || 'hat';
          return '\\' + a + '{' + arg(kid(n, 'e')) + '}';
        }
        case 'bar': return (val(kid(n, 'barPr'), 'pos') === 'top' ? '\\overline{' : '\\underline{') + arg(kid(n, 'e')) + '}';
        case 'groupChr': {
          const pr = kid(n, 'groupChrPr'), c = val(pr, 'chr'), pos = val(pr, 'pos');
          const top = c === '⏞' || (c === null ? pos === 'top' : pos === 'top' && c !== '⏟');
          if (c && !/[⏞⏟]/.test(c)) return (pos === 'top' ? '\\overset{' : '\\underset{') + (SYM[c] || c) + '}{' + arg(kid(n, 'e')) + '}';
          return (top ? '\\overbrace{' : '\\underbrace{') + arg(kid(n, 'e')) + '}';
        }
        case 'borderBox': return '\\boxed{' + arg(kid(n, 'e')) + '}';
        case 'box': return arg(kid(n, 'e'));
        case 'phant': {
          const pr = kid(n, 'phantPr');
          return on(pr, 'show') ? arg(kid(n, 'e')) : '\\phantom{' + arg(kid(n, 'e')) + '}';
        }
        case 'eqArr': {
          const ls = rows(n, 'e').map(arg).map(t => t.replace(/\\#\s*\\left\(.*$|#.*$/, '').trim());
          return lines(ls);
        }
        case 'm': {
          const rs = rows(n, 'mr').map(r => kids(r, 'e').map(arg).join(' & '));
          return '\\begin{matrix}' + rs.join('\\\\') + '\\end{matrix}';
        }
        case 'oMathPara': {
          const ls = kids(n, 'oMath').map(arg).filter(t => t.trim());
          return ls.length > 1 ? lines(ls) : (ls[0] || '');
        }
        case 'oMath': return arg(n);
        case 'e': case 'num': case 'den': case 'sub': case 'sup': case 'deg': case 'lim': case 'fName': return arg(n);
        default:
          if (/Pr$/.test(k)) return '';                                // Eigenschaften (ctrlPr, fPr …) erzeugen nichts
          return arg(n);
      }
    }
    // Mehrere Zeilen: an „&“ (Word-Ausrichtungsmarke) ausrichten, sonst am ersten Relationszeichen
    function lines(ls) {
      if (ls.some(t => /(^|[^\\])&/.test(t))) return '\\begin{aligned}' + ls.join('\\\\') + '\\end{aligned}';
      return '\\begin{aligned}' + ls.map(alignAt).join('\\\\') + '\\end{aligned}';
    }
    return tidy(tree.name === '#root' ? arg(tree) : node(tree));
  }

  // erstes Relationszeichen außerhalb von Klammern bekommt die Ausrichtungsmarke
  function alignAt(t) {
    let d = 0;
    for (let i = 0; i < t.length; i++) {
      const c = t[i];
      if (c === '\\') {
        if (!d && /^\\(approx|le|ge|ne|equiv|sim|simeq|propto)(?![a-zA-Z])/.test(t.slice(i))) return t.slice(0, i) + '&' + t.slice(i);
        i += (t.slice(i).match(/^\\([a-zA-Z]+|.)/) || ['\\'])[0].length - 1; continue;
      }
      if (c === '{') d++; else if (c === '}') d--;
      else if (!d && (c === '=' || c === '<' || c === '>')) return t.slice(0, i) + '&' + t.slice(i);
    }
    return '&' + t;
  }

  // überflüssige Leerzeichen weg, aber nie das Leerzeichen von „\ “
  const tidy = t => t.replace(/\s+/g, ' ').replace(/(?<!\\) (?=[_^}{)\]&,;.!=+\-/|\\])/g, '').replace(/(\\ |\s)+$/, '').trim();

  // Schriftgröße (pt) der Formel aus dem ersten Lauf, sonst null
  function sizeOf(xml) {
    const m = xml.match(/<m:r>[\s\S]*?<w:sz w:val="(\d+)"/);
    return m ? +m[1] / 2 : null;
  }

  // Alle Word-Formeln in einem XML-Stück: abgesetzte Blöcke (m:oMathPara) als eine Formel, sonst jede m:oMath
  const RE_ALL = /<m:oMathPara[\s>][\s\S]*?<\/m:oMathPara>|<m:oMath[\s>][\s\S]*?<\/m:oMath>/g;
  const mText = x => unxml((x.match(/<m:t(?:\s[^>]*)?>[^<]*<\/m:t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join(''));
  function list(xml) {
    const out = [];
    let m;
    RE_ALL.lastIndex = 0;
    while ((m = RE_ALL.exec(xml))) out.push({xml: m[0], start: m.index, end: m.index + m[0].length, display: /^<m:oMathPara/.test(m[0]), text: mText(m[0])});
    return out;
  }

  // Wo stehen die Formeln im Text des Absatzes? Zählt Zeichen so, wie Word sie in paragraph.text liefert
  // (Text, Tabulatoren, Formelzeichen). Liefert je Formel [von, bis) und die Gesamtlänge zum Abgleich.
  function spans(paraXml) {
    const found = list(paraXml);
    const plain = x => {
      let n = 0;
      (x.match(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>|<w:(?:tab|br|cr|noBreakHyphen|sym)\b[^>]*\/?>|<w:delText\b[\s\S]*?<\/w:delText>/g) || []).forEach(t => {
        if (/^<w:t[\s>]/.test(t)) n += Array.from(unxml(t.replace(/<[^>]+>/g, ''))).length;
        else if (!/^<w:delText/.test(t)) n += 1;
      });
      return n;
    };
    let pos = 0, at = 0;
    const out = [];
    found.forEach(f => {
      pos += plain(paraXml.slice(at, f.start));
      const len = Array.from(f.text).length;
      out.push(Object.assign({}, f, {from: pos, to: pos + len}));
      pos += len; at = f.end;
    });
    pos += plain(paraXml.slice(at));
    return {items: out, length: pos};
  }

  const api = {parse, toLatex, sizeOf, alignAt, list, spans, mText, runText};
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Omml = api;
})(typeof self !== 'undefined' ? self : this);
