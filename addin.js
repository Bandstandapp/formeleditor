/* Formeleditor als Word-Add-in: fügt Formeln als Bild (SVG mit PNG-Ersatz) an der Cursorposition ein
   und lädt eine angeklickte Formel zum Bearbeiten zurück. Die Formel selbst steckt im Alternativtext
   des Bildes („FORMEL:{…}“), genau wie beim Kopieren aus der Web-Version. */
(function () {
  const $ = id => document.getElementById(id);
  const status = $('copymsg');
  const pasteMsg = $('pastemsg');
  let current = null;            // Meta der zuletzt im Dokument angeklickten Formel
  let wordReady = false;

  // Bedienelemente der Web-Version auf das Add-in umstellen
  const oldCopy = $('copy');
  const insertBtn = oldCopy.cloneNode(true);   // Klon ohne den Kopieren-Handler
  oldCopy.replaceWith(insertBtn);
  insertBtn.textContent = 'In Word einfügen';
  // Derselbe Knopf noch einmal direkt unter der Eingabe, mit eigener Rückmeldung
  const insertBtn2 = insertBtn.cloneNode(true);
  insertBtn2.id = 'copy2';
  const status2 = document.createElement('p');
  status2.className = 'hint'; status2.setAttribute('role', 'status');
  const eingabeRow = document.querySelector('section[aria-label="Eingabe"] .row');
  const row2 = document.createElement('div');
  row2.className = 'row';
  row2.appendChild(insertBtn2);
  eingabeRow.after(row2, status2);
  const oldLoad = $('fromword');
  const loadBtn = oldLoad.cloneNode(true);
  oldLoad.replaceWith(loadBtn);
  loadBtn.textContent = 'Markierte Formel bearbeiten';
  pasteMsg.textContent = 'Zum Ändern einer Formel: sie im Dokument anklicken. Sie erscheint dann hier automatisch.';
  status.textContent = 'Formel bauen, dann „In Word einfügen“. Sie landet an der Stelle des Cursors.';

  const say = (el, text, ok) => { el.textContent = text; el.classList.toggle('ok', !!ok); };
  const xmlEsc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const b64utf8 = t => { const u = new TextEncoder().encode(t); let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };

  function setEditing(meta) {
    current = meta;
    insertBtn.textContent = insertBtn2.textContent = meta ? 'Formel in Word ersetzen' : 'In Word einfügen';
  }

  // Ein Absatz mit einem einzigen Bild; Word übernimmt es beim Einfügen in die laufende Zeile.
  function ooxml(r, meta) {
    const emu = pt => Math.round(pt * 12700);
    const cx = emu(r.wpt), cy = emu(r.hpt);
    const png = r.dataUrl.split(',')[1];
    const svg = b64utf8(r.svg);
    const drop = -Math.round(r.dpt * 2);    // Grundlinie: Bild um die Unterlänge absenken (Halbpunkte)
    const d = xmlEsc(meta);
    const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + rel + '/officeDocument" Target="word/document.xml"/></Relationships>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rIdFPng" Type="' + rel + '/image" Target="media/formel.png"/>' +
      '<Relationship Id="rIdFSvg" Type="' + rel + '/image" Target="media/formel.svg"/>' +
      '</Relationships></pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="' + rel + '"' +
      ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
      ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:rPr><w:position w:val="' + drop + '"/></w:rPr><w:drawing>' +
      '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:docPr id="1" name="Formel" descr="' + d + '"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
      '<pic:nvPicPr><pic:cNvPr id="0" name="formel.svg" descr="' + d + '"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rIdFPng"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
      '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdFSvg"/></a:ext></a:extLst></a:blip>' +
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/media/formel.png" pkg:contentType="image/png" pkg:compression="store"><pkg:binaryData>' + png + '</pkg:binaryData></pkg:part>' +
      '<pkg:part pkg:name="/word/media/formel.svg" pkg:contentType="image/svg+xml" pkg:compression="store"><pkg:binaryData>' + svg + '</pkg:binaryData></pkg:part>' +
      '</pkg:package>';
  }

  // Die Binärdatei der alten Formel aus dem OOXML der Markierung holen
  function oleBytes(flat) {
    const doc = new DOMParser().parseFromString(flat, 'application/xml');
    const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const parts = Array.from(doc.getElementsByTagNameNS(PKG, 'part'));
    const partData = name => {
      const p = parts.find(x => x.getAttributeNS(PKG, 'name') === name);
      const d = p && p.getElementsByTagNameNS(PKG, 'binaryData')[0];
      return d ? d.textContent.replace(/\s+/g, '') : null;
    };
    let b64 = null;
    const ole = Array.from(doc.getElementsByTagNameNS('urn:schemas-microsoft-com:office:office', 'OLEObject'))
      .find(o => /^Equation\.3/.test(o.getAttribute('ProgID') || ''));
    if (ole) {
      const rid = ole.getAttributeNS(R, 'id');
      const rels = parts.find(x => /\/word\/_rels\/document\.xml\.rels$/.test(x.getAttributeNS(PKG, 'name')));
      const rel = rels && Array.from(rels.getElementsByTagName('Relationship')).find(r => r.getAttribute('Id') === rid);
      if (rel) b64 = partData('/word/' + rel.getAttribute('Target').replace(/^\/?(word\/)?/, ''));
    }
    if (!b64) {
      const p = parts.find(x => /\.bin$/.test(x.getAttributeNS(PKG, 'name')));
      if (p) b64 = partData(p.getAttributeNS(PKG, 'name'));
    }
    if (!b64) throw new Error('Formeldaten fehlen');
    const bin = atob(b64), u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }

  // Steht die Formel allein im Absatz (höchstens mit Gleichungsnummer), war sie abgesetzt
  const standsAlone = text => text.replace(/\(\s*[0-9][0-9.,a-z]*\s*\)/gi, '').replace(/[\s\u0000-\u001f ￼]/g, '').length <= 1;

  // Formel in der aktuellen Markierung: neue (Bild mit FORMEL-Text) oder alte (Formel-Editor 3.0)
  async function selectedFormula(ctx) {
    const sel = ctx.document.getSelection();
    const pics = sel.inlinePictures;
    pics.load('items/altTextDescription,items/altTextTitle');
    sel.load('text');
    await ctx.sync();
    for (const p of pics.items) {
      const m = findMeta(null, (p.altTextDescription || '') + ' ' + (p.altTextTitle || ''));
      if (m) return {kind: 'new', pic: p, meta: m};
    }
    // Nur kurze Markierungen prüfen: ein angeklicktes Formelobjekt hat (fast) keinen Text
    if (sel.text.replace(/\s/g, '').length > 2) return null;
    const ox = sel.getOoxml();
    await ctx.sync();
    if (!/Equation\.3/.test(ox.value)) return null;
    let tex;
    try { tex = Eqn3.oleToLatex(oleBytes(ox.value)); }
    catch (e) { const err = new Error(e.message); err.oldFormula = true; throw err; }
    const para = sel.paragraphs.getFirst();
    para.load('text');
    sel.font.load('size');
    await ctx.sync();
    const pt = Math.round(sel.font.size || sizePt);
    return {kind: 'old', range: sel, meta: {tex, mode: standsAlone(para.text) ? 'sci' : 'lin', font, pt: pt > 5 && pt < 40 ? pt : sizePt}};
  }

  async function checkSelection(manual) {
    if (!wordReady) return;
    try {
      await Word.run(async ctx => {
        const f = await selectedFormula(ctx);
        if (f) {
          loadMeta(f.meta);
          setEditing(f.meta);
          say(pasteMsg, f.kind === 'old'
            ? 'Alte Formel übernommen. Bitte kurz prüfen, bei Bedarf ändern und dann „Formel in Word ersetzen“ drücken. Danach ist sie im neuen Format.'
            : 'Formel aus dem Dokument geladen. Ändern und dann „Formel in Word ersetzen“ drücken.', true);
        } else {
          setEditing(null);
          if (manual) say(pasteMsg, 'Im Dokument ist keine Formel aus diesem Editor markiert. Bitte die Formel einmal anklicken.');
        }
      });
    } catch (e) {
      setEditing(null);
      if (e.oldFormula) say(pasteMsg, 'Diese alte Formel ließ sich nicht lesen (' + e.message + '). Sie muss neu eingegeben werden.');
      else if (manual) say(pasteMsg, 'Word hat die Markierung nicht herausgegeben: ' + e.message);
    }
  }

  let selTimer;
  function onSelectionChanged() { clearTimeout(selTimer); selTimer = setTimeout(() => checkSelection(false), 150); }

  const sayStatus = (text, ok) => { say(status, text, ok); say(status2, text, ok); };

  async function insertIntoWord() {
    const src = mf ? mf.getValue('latex') : '';
    if (!src.trim()) { sayStatus('Erst eine Formel eingeben.'); return; }
    if (!wordReady) { sayStatus('Das geht nur, wenn der Editor in Word geöffnet ist.'); return; }
    insertBtn.disabled = insertBtn2.disabled = true;
    sayStatus('Formel wird eingefügt …');
    try {
      const pad = current && current.pad ? current.pad : 0;   // Luft um umgewandelte Formeln im Fließtext beibehalten
      const meta = 'FORMEL:' + JSON.stringify(pad ? {tex: src, mode, font, pt: sizePt, pad} : {tex: src, mode, font, pt: sizePt});
      const r = await frame.contentWindow.__png(sizePt, 0, pad);
      const xml = ooxml(r, meta);
      let replaced = false;
      await Word.run(async ctx => {
        const f = current ? await selectedFormula(ctx) : null;
        if (f && f.kind === 'new') { f.pic.getRange('Whole').insertOoxml(xml, 'Replace'); replaced = true; }
        else if (f) { f.range.insertOoxml(xml, 'Replace'); replaced = true; }
        else ctx.document.getSelection().insertOoxml(xml, 'Replace');
        await ctx.sync();
      });
      setEditing(null);
      sayStatus(replaced ? 'Formel im Dokument ersetzt.' : 'Formel eingefügt.', true);
    } catch (e) {
      sayStatus('Einfügen hat nicht geklappt: ' + (e && e.message ? e.message : e));
    } finally {
      insertBtn.disabled = insertBtn2.disabled = false;
    }
  }

  insertBtn.addEventListener('click', insertIntoWord);
  insertBtn2.addEventListener('click', insertIntoWord);
  loadBtn.addEventListener('click', () => checkSelection(true));
  // „Neue Formel“ trennt den Editor von der zuletzt angeklickten Formel
  $('clear').addEventListener('click', () => setEditing(null));

  Office.onReady(info => {
    if (info.host !== Office.HostType.Word) { say(status, 'Dieses Add-in läuft nur in Word.'); return; }
    wordReady = true;
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    checkSelection(false);
  });
})();
