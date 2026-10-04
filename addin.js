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
    insertBtn.textContent = meta ? 'Formel in Word ersetzen' : 'In Word einfügen';
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

  // Erste Formel-Grafik in der aktuellen Markierung (oder null)
  async function selectedFormula(ctx) {
    const pics = ctx.document.getSelection().inlinePictures;
    pics.load('items/altTextDescription,items/altTextTitle');
    await ctx.sync();
    for (const p of pics.items) {
      const m = findMeta(null, (p.altTextDescription || '') + ' ' + (p.altTextTitle || ''));
      if (m) return {pic: p, meta: m};
    }
    return null;
  }

  async function checkSelection(manual) {
    if (!wordReady) return;
    try {
      await Word.run(async ctx => {
        const f = await selectedFormula(ctx);
        if (f) {
          loadMeta(f.meta);
          setEditing(f.meta);
          say(pasteMsg, 'Formel aus dem Dokument geladen. Ändern und dann „Formel in Word ersetzen“ drücken.', true);
        } else {
          setEditing(null);
          if (manual) say(pasteMsg, 'Im Dokument ist keine Formel aus diesem Editor markiert. Bitte die Formel einmal anklicken.');
        }
      });
    } catch (e) {
      if (manual) say(pasteMsg, 'Word hat die Markierung nicht herausgegeben: ' + e.message);
    }
  }

  let selTimer;
  function onSelectionChanged() { clearTimeout(selTimer); selTimer = setTimeout(() => checkSelection(false), 150); }

  async function insertIntoWord() {
    const src = mf ? mf.getValue('latex') : '';
    if (!src.trim()) { say(status, 'Erst eine Formel eingeben.'); return; }
    if (!wordReady) { say(status, 'Das geht nur, wenn der Editor in Word geöffnet ist.'); return; }
    insertBtn.disabled = true;
    say(status, 'Formel wird eingefügt …');
    try {
      const meta = 'FORMEL:' + JSON.stringify({tex: src, mode, font, pt: sizePt});
      const r = await frame.contentWindow.__png(sizePt);
      const xml = ooxml(r, meta);
      let replaced = false;
      await Word.run(async ctx => {
        const f = current ? await selectedFormula(ctx) : null;
        if (f) { f.pic.getRange('Whole').insertOoxml(xml, 'Replace'); replaced = true; }
        else ctx.document.getSelection().insertOoxml(xml, 'Replace');
        await ctx.sync();
      });
      setEditing(null);
      say(status, replaced ? 'Formel im Dokument ersetzt.' : 'Formel eingefügt.', true);
    } catch (e) {
      say(status, 'Einfügen hat nicht geklappt: ' + (e && e.message ? e.message : e));
    } finally {
      insertBtn.disabled = false;
    }
  }

  insertBtn.addEventListener('click', insertIntoWord);
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
