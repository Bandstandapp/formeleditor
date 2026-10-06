/* Alle Formeln eines Word-Dokuments (.docx) in das neue Format umwandeln.
   Arbeitet auf einer Kopie des Dokuments (JSZip): alte Formeln aus dem Formel-Editor 3.0 werden durch
   Formel-Bilder mit FORMEL-Alternativtext ersetzt, nicht umwandelbare Formeln bleiben stehen und
   werden gelb markiert. Das Setzen der Formeln übernimmt eine übergebene Funktion render(tex, mode, pt). */
(function (root) {
  const RELNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const FORMATS = {
    'Equation.3': 'Formel-Editor 3.0',
    'Equation.2': 'Formel-Editor 2.0',
    'Equation.DSMT4': 'MathType',
    'Equation.DSMT36': 'MathType',
  };
  const formatName = id => FORMATS[id] || (/DSMT/.test(id) ? 'MathType' : 'Formelobjekt ' + id);
  const NEW = 'Neuer Formeleditor (schon umgewandelt)';
  const OMML = 'Word-Formel (bleibt unverändert)';

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

  // zip: JSZip-Objekt des Dokuments (wird verändert). render(tex, mode, pt) → {dataUrl|png, svg, wpt, hpt, dpt}
  async function convertDocx(zip, render, opts) {
    opts = opts || {};
    const font = opts.font || 'stix2';
    const progress = opts.onProgress || (() => {});
    const read = async name => (zip.file(name) ? zip.file(name).async('string') : null);
    const sizeOf = sizeResolver((await read('word/styles.xml')) || '');
    const report = {formats: {}, total: 0, converted: 0, problems: [], unchanged: 0};
    const count = name => { report.formats[name] = (report.formats[name] || 0) + 1; };

    const partNames = Object.keys(zip.files).filter(n => /^word\/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$/.test(n))
      .sort((a, b) => (a === 'word/document.xml' ? -1 : b === 'word/document.xml' ? 1 : a.localeCompare(b)));
    const parts = [];
    const usedIds = new Set();
    for (const name of partNames) {
      const xml = await read(name);
      (xml.match(/<wp:docPr\b[^>]*>/g) || []).forEach(t => usedIds.add(+attr(t, 'id')));
      parts.push({name, xml});
    }
    let nextId = 1000;
    const newId = () => { while (usedIds.has(nextId)) nextId++; usedIds.add(nextId); return nextId++; };

    // Erst alle Formeln zählen, damit der Fortschritt stimmt
    const jobs = [];
    for (const part of parts) {
      const {xml} = part;
      (xml.match(/<wp:docPr\b[^>]*>/g) || []).forEach(t => { if (/^FORMEL:/.test(attr(t, 'descr') || '')) { count(NEW); report.unchanged++; } });
      (xml.match(/<m:oMath[\s>]/g) || []).forEach(() => { count(OMML); report.unchanged++; });
      const re = /<w:object\b[\s\S]*?<\/w:object>/g;
      let m;
      while ((m = re.exec(xml))) {
        const prog = (m[0].match(/<o:OLEObject\b[^>]*\bProgID="([^"]+)"/) || [])[1];
        if (!prog || !/^Equation\./.test(prog)) continue;       // andere eingebettete Objekte (Tabellen …) nicht anfassen
        jobs.push({part, start: m.index, end: m.index + m[0].length, obj: m[0], prog});
      }
    }
    report.total = jobs.length + report.unchanged;

    const relsOf = {};
    async function rels(part) {
      if (relsOf[part.name]) return relsOf[part.name];
      const relName = part.name.replace(/^word\//, 'word/_rels/') + '.rels';
      const xml = (await read(relName)) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
      const map = {};
      (xml.match(/<Relationship\b[^>]*>/g) || []).forEach(t => { map[attr(t, 'Id')] = attr(t, 'Target'); });
      return (relsOf[part.name] = {name: relName, xml, map, add: ''});
    }
    const resolve = (target) => {
      const segs = ('word/' + target.replace(/^\//, '').replace(/^word\//, '')).split('/');
      const out = [];
      segs.forEach(s => { if (s === '..') out.pop(); else if (s !== '.' && s) out.push(s); });
      return out.join('/');
    };

    let media = 0;
    for (let k = 0; k < jobs.length; k++) {
      const j = jobs[k], xml = j.part.xml;
      progress(k + 1, jobs.length);
      count(formatName(j.prog));
      const rs = Math.max(xml.lastIndexOf('<w:r>', j.start), xml.lastIndexOf('<w:r ', j.start));
      const reEnd = xml.indexOf('</w:r>', j.end);
      const runOpen = xml.slice(rs, xml.indexOf('>', rs) + 1);
      const before = xml.slice(runOpen.length + rs, j.start);
      const rPrM = before.match(/^\s*(<w:rPr>[\s\S]*?<\/w:rPr>)\s*$/);
      const simple = rPrM !== null || /^\s*$/.test(before);
      const rPr = rPrM ? rPrM[1] : '';
      const after = xml.slice(j.end, reEnd);
      const ps = Math.max(xml.lastIndexOf('<w:p>', j.start), xml.lastIndexOf('<w:p ', j.start));
      const pe = xml.indexOf('</w:p>', j.end);
      const para = xml.slice(ps, pe);
      const pStyle = (para.match(/<w:pPr>[\s\S]*?<w:pStyle w:val="([^"]+)"/) || [])[1];
      const ctx = textOf(xml.slice(ps, j.start)).trim();
      const where = ctx ? 'nach „…' + ctx.slice(-40) + '“' : 'am Absatzanfang';
      j.number = k + 1;
      let tex = null, why = null;
      if (j.prog !== 'Equation.3') why = formatName(j.prog) + ' wird nicht unterstützt';
      else {
        try {
          const rid = attr(j.obj.match(/<o:OLEObject\b[^>]*>/)[0], 'r:id');
          const r = await rels(j.part);
          const target = r.map[rid];
          if (!target) throw Object.assign(new Error('rel'), {userText: 'Formeldaten fehlen'});
          const f = zip.file(resolve(target));
          if (!f) throw Object.assign(new Error('part'), {userText: 'Formeldaten fehlen'});
          tex = root.Eqn3.oleToLatex(await f.async('uint8array'));
          if (!tex || !tex.trim()) throw Object.assign(new Error('leer'), {userText: 'Formel ist leer'});
        } catch (e) { why = e.userText || 'Formeldaten beschädigt, nicht lesbar'; }
      }
      const pt = Math.round(sizeOf(rPr, pStyle) * 2) / 2;
      const mode = standsAlone(textOf(para)) ? 'sci' : 'lin';
      // Im Fließtext etwas Luft links und rechts, wie beim alten Formelobjekt (dort fehlen oft die Leerzeichen)
      const pad = mode === 'lin' ? 0.12 : 0;
      let img = null;
      if (tex) {
        try { img = await render(tex, mode, pt, pad); } catch (e) { why = 'Neue Darstellung fehlgeschlagen'; }
      }
      if (img && simple && /^\s*$/.test(after)) {
        const r = await rels(j.part);
        const n = ++media;
        const ridPng = 'rIdFe' + n + 'p', ridSvg = 'rIdFe' + n + 's';
        const png = img.png || b64ToBytes(img.dataUrl.split(',')[1]);
        zip.file('word/media/fe' + n + '.png', png);
        zip.file('word/media/fe' + n + '.svg', img.svg);
        r.add += '<Relationship Id="' + ridPng + '" Type="' + RELNS + '/image" Target="media/fe' + n + '.png"/>' +
                 '<Relationship Id="' + ridSvg + '" Type="' + RELNS + '/image" Target="media/fe' + n + '.svg"/>';
        const meta = 'FORMEL:' + JSON.stringify(pad ? {tex, mode, font, pt, pad} : {tex, mode, font, pt});
        j.out = runOpen + setRPr(rPr, 'position', -Math.round(img.dpt * 2)) + drawingXml(newId(), ridPng, ridSvg, img, meta) + '</w:r>';
        j.from = rs; j.to = reEnd + 6;
        report.converted++;
      } else {
        if (!why) why = 'Formel steht an einer ungewöhnlichen Stelle';
        report.problems.push({number: j.number, where, why, tex});
        if (simple) { j.out = runOpen + setRPr(rPr, 'highlight', 'yellow') + xml.slice(j.start, reEnd + 6); j.from = rs; j.to = reEnd + 6; }
      }
    }

    // Teile neu zusammensetzen
    for (const part of parts) {
      const mine = jobs.filter(j => j.part === part && j.out).sort((a, b) => a.from - b.from);
      if (!mine.length) continue;
      let out = '', pos = 0;
      mine.forEach(j => { out += part.xml.slice(pos, j.from) + j.out; pos = j.to; });
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
    return report;
  }

  const api = {convertDocx, standsAlone, setRPr, sizeResolver, formatName};
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.FeConvert = api;
})(typeof window !== 'undefined' ? window : globalThis);
