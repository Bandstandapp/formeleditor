/* Alle Formeln eines Word-Dokuments (.docx) in das neue Format umwandeln.
   Arbeitet auf einer Kopie des Dokuments (JSZip): alte Formeln aus dem Formel-Editor 3.0 werden durch
   Formel-Bilder mit FORMEL-Alternativtext ersetzt, nicht umwandelbare Formeln bleiben stehen und
   werden gelb markiert. Formeln aus dem heutigen Formel-Editor von Word (OMML) werden ebenso umgewandelt. Das Setzen der Formeln übernimmt eine übergebene Funktion render(tex, mode, pt). */
(function (root) {
  const RELNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const FORMATS = {
    'Equation.3': 'Formel-Editor 3.0',
    'Equation.2': 'Formel-Editor 2.0',
    'Equation.DSMT4': 'MathType',
    'Equation.DSMT36': 'MathType',
  };
  const formatName = id => FORMATS[id] || (/DSMT/.test(id) ? 'MathType' : 'Formelobjekt ' + id);
  const NEW = 'Neues Format (Formeleditor)';
  const OMML = 'Word-Formel (heutiger Formel-Editor von Word)';

  const unxml = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const xmlEsc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const attr = (tag, name) => { const m = tag.match(new RegExp('\\s' + name + '="([^"]*)"')); return m ? unxml(m[1]) : null; };
  const textOf = xml => unxml((xml.match(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join(''));

  // Steht die Formel allein im Absatz (höchstens mit Gleichungsnummer), war sie abgesetzt
  const standsAlone = text => text.replace(/\(\s*[0-9][0-9.,a-z]*\s*\)/gi, '').replace(/[\s\u0000-\u001f ￼]/g, '').length <= 1;

  // Reihenfolge der Kinder von w:rPr laut Schema; neue Einträge müssen an die richtige Stelle
  const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss',
    'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight',
    'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];
  function setRPr(rPr, name, val) {
    const el = '<w:' + name + ' w:val="' + val + '"/>';
    if (!rPr) return '<w:rPr>' + el + '</w:rPr>';
    const own = new RegExp('<w:' + name + '(?:\\s[^>]*)?/>');
    if (own.test(rPr)) return rPr.replace(own, el);
    const later = RPR_ORDER.slice(RPR_ORDER.indexOf(name) + 1);
    const re = /<w:([A-Za-z]+)[\s/>]/g;
    let m;
    re.lastIndex = 7;                      // nach „<w:rPr>“
    while ((m = re.exec(rPr))) {
      if (later.includes(m[1])) return rPr.slice(0, m.index) + el + rPr.slice(m.index);
    }
    return rPr.replace(/<\/w:rPr>$/, el + '</w:rPr>');
  }

  // Den letzten Absatz in xml zentrieren (w:jc an die vom Schema verlangte Stelle in w:pPr)
  const PPR_AFTER_JC = ['textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
  function centerLast(xml) {
    const ps = Math.max(xml.lastIndexOf('<w:p>'), xml.lastIndexOf('<w:p '));
    if (ps < 0) return xml;
    const open = xml.indexOf('>', ps) + 1;
    const jc = '<w:jc w:val="center"/>';
    if (!xml.startsWith('<w:pPr>', open)) return xml.slice(0, open) + '<w:pPr>' + jc + '</w:pPr>' + xml.slice(open);
    const end = xml.indexOf('</w:pPr>', open);
    if (end < 0) return xml;
    const re = /<w:([A-Za-z]+)[\s/>]/g;
    re.lastIndex = open + 7;
    let m;
    while ((m = re.exec(xml)) && m.index < end) {
      if (PPR_AFTER_JC.includes(m[1])) return xml.slice(0, m.index) + jc + xml.slice(m.index);
    }
    return xml.slice(0, end) + jc + xml.slice(end);
  }

  // Schriftgröße (pt) aus den Formatvorlagen: Zeichen → Absatzvorlage (mit basedOn) → Standard
  function sizeResolver(styles) {
    const map = {};
    let defPara = null;
    (styles.match(/<w:style\b[\s\S]*?<\/w:style>/g) || []).forEach(s => {
      const open = s.slice(0, s.indexOf('>'));
      const id = attr(open, 'w:styleId');
      const based = s.match(/<w:basedOn w:val="([^"]+)"/);
      const rPr = s.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
      const sz = rPr && rPr[0].match(/<w:sz w:val="(\d+)"/);
      map[id] = {based: based && based[1], sz: sz ? +sz[1] / 2 : null};
      if (attr(open, 'w:type') === 'paragraph' && /^(1|true)$/.test(attr(open, 'w:default') || '')) defPara = id;
    });
    const dd = styles.match(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/);
    const ddSz = dd && dd[0].match(/<w:sz w:val="(\d+)"/);
    const docDefault = ddSz ? +ddSz[1] / 2 : 10;
    const chain = id => { for (let i = 0; id && map[id] && i < 20; i++, id = map[id].based) if (map[id].sz) return map[id].sz; return null; };
    return (runRPr, pStyle) => {
      const own = runRPr && runRPr.match(/<w:sz w:val="(\d+)"/);
      if (own) return +own[1] / 2;
      const rs = runRPr && runRPr.match(/<w:rStyle w:val="([^"]+)"/);
      return (rs && chain(rs[1])) || chain(pStyle || defPara) || chain(defPara) || docDefault;
    };
  }

  function drawingXml(id, ridPng, ridSvg, r, meta) {
    const emu = pt => Math.round(pt * 12700);
    const cx = emu(r.wpt), cy = emu(r.hpt), d = xmlEsc(meta);
    return '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:docPr id="' + id + '" name="Formel ' + id + '" descr="' + d + '"/>' +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="0" name="formel' + id + '.svg" descr="' + d + '"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + ridPng + '"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
      '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="' + ridSvg + '"/></a:ext></a:extLst></a:blip>' +
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
  }

  const b64ToBytes = b64 => { const s = atob(b64), u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; };

  const FONT_NAMES = {stix2: 'STIX Two', termes: 'Termes', newcm: 'Computer Modern', pagella: 'Pagella', schola: 'Schola', bonum: 'Bonum', fira: 'Fira', dejavu: 'DejaVu'};
  const OLD_FONT = 'Schrift des alten Formel-Editors';

  // Alle Formeln eines Dokuments finden. Liefert die Teile (XML) und je Formel einen Eintrag in Dokumentreihenfolge.
  async function scan(zip) {
    const read = async name => (zip.file(name) ? zip.file(name).async('string') : null);
    const sizeOf = sizeResolver((await read('word/styles.xml')) || '');
    const partNames = Object.keys(zip.files).filter(n => /^word\/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$/.test(n))
      .sort((a, b) => (a === 'word/document.xml' ? -1 : b === 'word/document.xml' ? 1 : a.localeCompare(b)));
    const parts = [], items = [];
    for (const name of partNames) {
      const xml = await read(name);
      const part = {name, xml};
      parts.push(part);
      const found = [];
      let m;
      const reObj = /<w:object\b[\s\S]*?<\/w:object>/g;
      while ((m = reObj.exec(xml))) {
        const prog = (m[0].match(/<o:OLEObject\b[^>]*\bProgID="([^"]+)"/) || [])[1];
        if (!prog || !/^Equation\./.test(prog)) continue;       // andere eingebettete Objekte (Tabellen …) nicht anfassen
        found.push({kind: 'old', part, start: m.index, end: m.index + m[0].length, obj: m[0], prog, format: formatName(prog)});
      }
      const reDr = /<w:drawing>[\s\S]*?<\/w:drawing>/g;
      while ((m = reDr.exec(xml))) {
        const dp = m[0].match(/<wp:docPr\b[^>]*>/);
        const descr = dp && attr(dp[0], 'descr');
        if (!descr || !/^FORMEL:/.test(descr)) continue;
        let meta = null;
        try { meta = JSON.parse(descr.slice(7)); } catch (e) {}
        if (!meta || typeof meta.tex !== 'string') continue;
        found.push({kind: 'new', part, start: m.index, end: m.index + m[0].length, obj: m[0], meta, format: NEW});
      }
      // Word-Formeln: ein abgesetzter Block (m:oMathPara) zählt als eine Formel, sonst jede m:oMath einzeln
      const reMath = /<m:oMathPara[\s>][\s\S]*?<\/m:oMathPara>|<m:oMath[\s>][\s\S]*?<\/m:oMath>/g;
      while ((m = reMath.exec(xml))) found.push({kind: 'omml', part, start: m.index, end: m.index + m[0].length, obj: m[0], display: /^<m:oMathPara/.test(m[0]), format: OMML});
      found.sort((a, b) => a.start - b.start);
      for (const f of found) {
        const rs = Math.max(xml.lastIndexOf('<w:r>', f.start), xml.lastIndexOf('<w:r ', f.start));
        const ps = Math.max(xml.lastIndexOf('<w:p>', f.start), xml.lastIndexOf('<w:p ', f.start));
        const pe = xml.indexOf('</w:p>', f.end);
        f.rs = rs; f.reEnd = xml.indexOf('</w:r>', f.end);
        f.runOpen = rs >= 0 ? xml.slice(rs, xml.indexOf('>', rs) + 1) : '';
        const before = rs >= 0 ? xml.slice(rs + f.runOpen.length, f.start) : '';
        const rPrM = before.match(/^\s*(<w:rPr>[\s\S]*?<\/w:rPr>)\s*$/);
        f.rPr = rPrM ? rPrM[1] : '';
        // Word-Formeln stehen nicht in einem w:r, sondern direkt im Absatz (oder in einem Link, Feld …)
        if (f.kind === 'omml') f.rs = -1;
        f.simple = f.kind === 'omml' ? true : rs >= 0 && (rPrM !== null || /^\s*$/.test(before)) && /^\s*$/.test(xml.slice(f.end, f.reEnd));
        f.para = xml.slice(ps, pe);
        const pStyle = (f.para.match(/<w:pPr>[\s\S]*?<w:pStyle w:val="([^"]+)"/) || [])[1];
        f.textPt = Math.round(sizeOf(f.rPr, pStyle) * 2) / 2;
        const ctx = textOf(xml.slice(ps, f.start)).trim();
        f.where = ctx ? 'nach „…' + ctx.slice(-40) + '“' : 'am Absatzanfang';
        if (f.kind === 'new') { f.font = FONT_NAMES[f.meta.font] || f.meta.font || '?'; f.pt = +f.meta.pt || f.textPt; }
        else if (f.kind === 'old') { f.font = OLD_FONT; f.pt = f.textPt; }
        else { f.font = 'Cambria Math (Word)'; f.pt = (root.Omml && root.Omml.sizeOf(f.obj)) || f.textPt; }
        items.push(f);
      }
    }
    items.forEach((f, i) => { f.number = i + 1; });
    return {parts, items, read};
  }

  // Übersicht: Anzahl, Formate, Schriften, Schriftgrößen
  async function analyze(zip) {
    const {items} = await scan(zip);
    const tally = key => items.reduce((o, f) => { o[f[key]] = (o[f[key]] || 0) + 1; return o; }, {});
    return {total: items.length, formats: tally('format'), fonts: tally('font'), sizes: tally('pt')};
  }

  // zip: JSZip-Objekt des Dokuments (wird verändert). render(tex, mode, pt, pad, font) → {dataUrl|png, svg, wpt, hpt, dpt}
  // opts: font/pt (Standard), applyFont/applyPt (Standard auch auf vorhandene neue Formeln anwenden), onProgress
  async function convertDocx(zip, render, opts) {
    opts = opts || {};
    const stdFont = opts.font || 'stix2';
    const progress = opts.onProgress || (() => {});
    const {parts, items, read} = await scan(zip);
    const report = {formats: {}, total: items.length, converted: 0, reformatted: 0, problems: [], unchanged: 0};
    items.forEach(f => { report.formats[f.format] = (report.formats[f.format] || 0) + 1; });

    const usedIds = new Set();
    parts.forEach(p => (p.xml.match(/<wp:docPr\b[^>]*>/g) || []).forEach(t => usedIds.add(+attr(t, 'id'))));
    let nextId = 1000;
    const newId = () => { while (usedIds.has(nextId)) nextId++; usedIds.add(nextId); return nextId++; };

    const relsOf = {};
    async function rels(part) {
      if (relsOf[part.name]) return relsOf[part.name];
      const relName = part.name.replace(/^word\//, 'word/_rels/') + '.rels';
      const xml = (await read(relName)) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
      const map = {};
      (xml.match(/<Relationship\b[^>]*>/g) || []).forEach(t => { map[attr(t, 'Id')] = attr(t, 'Target'); });
      return (relsOf[part.name] = {name: relName, xml, map, add: ''});
    }
    const resolve = target => {
      const out = [];
      ('word/' + target.replace(/^\//, '').replace(/^word\//, '')).split('/').forEach(s => { if (s === '..') out.pop(); else if (s !== '.' && s) out.push(s); });
      return out.join('/');
    };

    // Neue Formeln mit leeren Kästchen (Abstände aus älteren Umwandlungen) werden immer repariert
    const broken = f => f.kind === 'new' && /[\uE000-\uF8FF]/.test(f.meta.tex);
    const jobs = items.filter(f => f.kind === 'old' || f.kind === 'omml' || broken(f) || (f.kind === 'new' && (opts.applyFont || opts.applyPt)));
    report.unchanged = items.length - jobs.length;
    let media = 0;
    for (let k = 0; k < jobs.length; k++) {
      const j = jobs[k], xml = j.part.xml;
      progress(k + 1, jobs.length);
      let tex = null, why = null, mode, pt, font, pad;
      if (j.kind === 'new') {
        tex = root.Eqn3 ? root.Eqn3.clean(j.meta.tex) : j.meta.tex; mode = ['lin', 'kom'].includes(j.meta.mode) ? j.meta.mode : 'sci'; pad = j.meta.pad || 0;
        font = opts.applyFont ? stdFont : (j.meta.font || stdFont);
        pt = opts.applyPt && opts.pt ? +opts.pt : (+j.meta.pt || j.textPt);
      } else if (j.kind === 'omml') {
        font = stdFont;
        pt = opts.applyPt && opts.pt ? +opts.pt : j.pt;
        // Abgesetzt wie in Word: Brüche übereinander; im Fließtext klein wie Words eigene Formeln
        const alone = j.display || standsAlone(textOf(j.para));
        mode = alone ? 'sci' : 'kom';
        pad = alone ? 0 : 0.12;
        try {
          tex = root.Omml.toLatex(j.obj);
          if (!tex || !tex.trim()) throw Object.assign(new Error('leer'), {userText: 'Formel ist leer'});
        } catch (e) { why = e.userText || 'Word-Formel ließ sich nicht lesen'; tex = null; }
      } else {
        font = stdFont;
        pt = opts.applyPt && opts.pt ? +opts.pt : j.textPt;
        // Der alte Formel-Editor setzt Brüche immer übereinander, auch im Fließtext; so bleibt es
        mode = 'sci';
        // Im Fließtext etwas Luft links und rechts, wie beim alten Formelobjekt (dort fehlen oft die Leerzeichen)
        pad = standsAlone(textOf(j.para)) ? 0 : 0.12;
        if (j.prog !== 'Equation.3') why = j.format + ' wird nicht unterstützt';
        else {
          try {
            const rid = attr(j.obj.match(/<o:OLEObject\b[^>]*>/)[0], 'r:id');
            const target = (await rels(j.part)).map[rid];
            if (!target) throw Object.assign(new Error('rel'), {userText: 'Formeldaten fehlen'});
            const f = zip.file(resolve(target));
            if (!f) throw Object.assign(new Error('part'), {userText: 'Formeldaten fehlen'});
            tex = root.Eqn3.oleToLatex(await f.async('uint8array'));
            if (!tex || !tex.trim()) throw Object.assign(new Error('leer'), {userText: 'Formel ist leer'});
          } catch (e) { why = e.userText || 'Formeldaten beschädigt, nicht lesbar'; tex = null; }
        }
      }
      let img = null;
      if (tex) {
        try { img = await render(tex, mode, pt, pad, font); } catch (e) { why = 'Neue Darstellung fehlgeschlagen'; }
      }
      if (img && j.simple) {
        const r = await rels(j.part);
        const n = ++media;
        const ridPng = 'rIdFe' + n + 'p', ridSvg = 'rIdFe' + n + 's';
        zip.file('word/media/fe' + n + '.png', img.png || b64ToBytes(img.dataUrl.split(',')[1]));
        zip.file('word/media/fe' + n + '.svg', img.svg);
        r.add += '<Relationship Id="' + ridPng + '" Type="' + RELNS + '/image" Target="media/fe' + n + '.png"/>' +
                 '<Relationship Id="' + ridSvg + '" Type="' + RELNS + '/image" Target="media/fe' + n + '.svg"/>';
        const m = {tex, mode, font, pt};
        if (pad) m.pad = pad;
        if (j.kind === 'omml') {
          // Word-Formel steht direkt im Absatz: an ihre Stelle kommt ein eigener Lauf mit dem Bild
          j.out = '<w:r>' + setRPr('', 'position', -Math.round(img.dpt * 2)) + drawingXml(newId(), ridPng, ridSvg, img, 'FORMEL:' + JSON.stringify(m)) + '</w:r>';
          j.from = j.start; j.to = j.end;
          // Ein abgesetzter Word-Block war zentriert; allein im Absatz bleibt er das
          if (j.display && !/<w:jc\b/.test(j.para.slice(0, j.para.indexOf('>') + 1) + (j.para.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0])) j.center = true;
          report.converted++;
          continue;
        }
        j.out = j.runOpen + setRPr(j.rPr, 'position', -Math.round(img.dpt * 2)) + drawingXml(newId(), ridPng, ridSvg, img, 'FORMEL:' + JSON.stringify(m)) + '</w:r>';
        j.from = j.rs; j.to = j.reEnd + 6;
        if (j.kind === 'old') report.converted++; else report.reformatted++;
      } else {
        if (!why) why = 'Formel steht an einer ungewöhnlichen Stelle';
        report.problems.push({number: j.number, where: j.where, why, tex});
        if (j.kind === 'omml') continue;              // bleibt als Word-Formel stehen
        if (j.simple) { j.out = j.runOpen + setRPr(j.rPr, 'highlight', 'yellow') + xml.slice(j.start, j.reEnd + 6); j.from = j.rs; j.to = j.reEnd + 6; }
      }
    }

    // Teile neu zusammensetzen
    for (const part of parts) {
      const mine = jobs.filter(j => j.part === part && j.out).sort((a, b) => a.from - b.from);
      if (!mine.length) continue;
      let out = '', pos = 0;
      mine.forEach(j => { const pre = part.xml.slice(pos, j.from); out += (j.center ? centerLast(pre) : pre) + j.out; pos = j.to; });
      zip.file(part.name, out + part.xml.slice(pos));
      const r = relsOf[part.name];
      if (r && r.add) zip.file(r.name, r.xml.replace('</Relationships>', r.add + '</Relationships>'));
    }
    if (media) {
      let ct = await read('[Content_Types].xml');
      if (!/Extension="png"/i.test(ct)) ct = ct.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default ');
      if (!/Extension="svg"/i.test(ct)) ct = ct.replace('<Default ', '<Default Extension="svg" ContentType="image/svg+xml"/><Default ');
      zip.file('[Content_Types].xml', ct);
    }
    report.changed = media;
    return report;
  }

  const api = {convertDocx, analyze, scan, standsAlone, setRPr, centerLast, drawingXml, sizeResolver, formatName, FONT_NAMES};
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.FeConvert = api;
})(typeof window !== 'undefined' ? window : globalThis);
