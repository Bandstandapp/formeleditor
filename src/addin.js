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
  row2.classList.add('fe-insrow');
  // Abbrechen verwirft alle Änderungen seit dem Laden bzw. dem letzten Einfügen
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button'; cancelBtn.className = 'btn fe-cancel'; cancelBtn.textContent = 'Abbrechen';
  cancelBtn.disabled = true;
  row2.append(cancelBtn, insertBtn2);
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

  // Rahmen um die Eingabe: grün = so wie im Dokument, orange = geändert, aber noch nicht in Word eingefügt
  let saved = null, savedState = null;
  const snap = () => JSON.stringify([mf ? mf.getValue('latex').replace(/\s+/g, '') : '', mode, font, sizePt]);   // Leerzeichen im Quelltext zählen nicht
  function markSaved() {
    saved = snap();
    savedState = {tex: mf.getValue('latex'), mode, font, pt: sizePt};
    updateFrame();
  }
  function cancelEdit() {
    if (savedState) {
      loadMeta(savedState);
      say(pasteMsg, 'Änderungen verworfen.', true);
    } else {
      mf.setValue('', {silenceNotifications: true});
      show();
      say(pasteMsg, '');
    }
    updateFrame();
  }
  function updateFrame() {
    if (!mf) return;
    const empty = !mf.getValue('latex').trim();
    const st = saved === null ? (empty ? '' : 'changed') : (snap() === saved ? 'saved' : 'changed');
    if (mf.dataset.state !== st) { if (st) mf.dataset.state = st; else delete mf.dataset.state; }
    // Wer eine neue Formel beginnt, statt den alten Entwurf zu wählen, verwirft ihn
    if (draftOffer && st === 'changed' && saved === null && !adoptNext && mf.getValue('latex').trim()) closeDraft();
    saveDraft(st);
    cancelBtn.disabled = st !== 'changed';
    // Unveränderte Formel aus dem Dokument: Der Knopf übernimmt sie nur (alte Formeln danach im neuen Format)
    const label = !current ? 'In Word einfügen' : st === 'saved' ? 'Ohne Änderung übernehmen' : 'Formel in Word ersetzen';
    if (insertBtn.textContent !== label) insertBtn.textContent = insertBtn2.textContent = label;
  }
  setInterval(updateFrame, 300);

  // Entwurf: Was noch nicht in Word ist, wird laufend gesichert. Schließt sich Word vorher (Absturz,
  // Seitenleiste zu, Aktualisieren), bietet der Editor die Formel beim nächsten Start wieder an.
  const DRAFT = 'fe-draft';
  let draftOffer = null, lastDraft = null, adoptNext = false;
  try { draftOffer = JSON.parse(localStorage.getItem(DRAFT) || 'null'); } catch (e) {}
  if (!draftOffer || typeof draftOffer.tex !== 'string' || !draftOffer.tex.trim()) draftOffer = null;
  function saveDraft(st) {
    if (draftOffer) return;                 // erst entscheiden lassen, sonst wäre der alte Entwurf weg
    const d = st === 'changed' ? JSON.stringify({tex: mf.getValue('latex'), mode, font, pt: sizePt, edit: !!current, t: Date.now()}) : '';
    if (d === lastDraft) return;
    lastDraft = d;
    try { if (d) localStorage.setItem(DRAFT, d); else localStorage.removeItem(DRAFT); } catch (e) {}
  }
  const draftBox = document.createElement('div');
  draftBox.className = 'fe-draft'; draftBox.hidden = true; draftBox.setAttribute('role', 'status');
  draftBox.innerHTML = '<p><b>Diese Formel war noch nicht in Word:</b></p><div class="fe-draft-f"></div>' +
    '<div class="row"><button class="btn" type="button" data-a="ja">Wiederherstellen</button><button class="btn" type="button" data-a="nein">Verwerfen</button></div>';
  $('mfwrap').before(draftBox);
  function closeDraft() { draftOffer = null; draftBox.hidden = true; lastDraft = null; }
  function offerDraft() {
    if (!draftOffer) return;
    const f = draftBox.querySelector('.fe-draft-f');
    f.textContent = '';
    texToSvg(draftOffer.tex).then(h => { f.innerHTML = h; }).catch(() => { f.textContent = draftOffer.tex; });
    draftBox.hidden = false;
  }
  draftBox.addEventListener('click', e => {
    const a = e.target.closest('button') && e.target.closest('button').dataset.a;
    if (!a) return;
    const d = draftOffer;
    closeDraft();
    if (a === 'nein') { try { localStorage.removeItem(DRAFT); } catch (err) {} return; }
    setEditing(null); saved = savedState = null;
    loadMeta(d);
    adoptNext = !!d.edit;
    say(pasteMsg, d.edit
      ? 'Formel wiederhergestellt. Soll sie eine Formel im Dokument ersetzen, diese jetzt im Dokument anklicken. Die Eingabe bleibt dabei erhalten.'
      : 'Formel wiederhergestellt. Mit „In Word einfügen“ kommt sie an die Stelle des Cursors.', true);
    updateFrame();
  });

  function setEditing(meta) {
    current = meta;
    insertBtn.textContent = insertBtn2.textContent = meta ? 'Formel in Word ersetzen' : 'In Word einfügen';
    if (typeof updateFrame === 'function') updateFrame();
  }

  // Ein Absatz mit einem einzigen Bild; Word übernimmt es beim Einfügen in die laufende Zeile.
  // num: Gleichungsnummer wie „4.13“; sie steht rechtsbündig am Tabstopp aus dem Lineal, ohne Lineal am rechten Rand
  // pPr: Absatzeigenschaften (Einzüge aus dem Lineal) für eine Formel im eigenen Absatz, sonst leer
  function ooxml(r, meta, num, pPr) {
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
      ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p>' + (pPr || '') + '<w:r><w:rPr><w:position w:val="' + drop + '"/></w:rPr><w:drawing>' +
      '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:docPr id="1" name="Formel" descr="' + d + '"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
      '<pic:nvPicPr><pic:cNvPr id="0" name="formel.svg" descr="' + d + '"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rIdFPng"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
      '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdFSvg"/></a:ext></a:extLst></a:blip>' +
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
      // Mit Lineal: Tabulator zum rechten Tabstopp aus pPr; sonst absoluter Tabulator am rechten Rand
      (num ? '<w:r>' + (pPr ? '<w:tab/>' : '<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/>') +
        '<w:t xml:space="preserve">(' + xmlEsc(num) + ')</w:t></w:r>' : '') +
      '</w:p></w:body></w:document></pkg:xmlData></pkg:part>' +
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
    return {kind: 'old', range: sel, meta: {tex, mode: 'sci', pad: standsAlone(para.text) ? 0 : 0.12, font, pt: pt > 5 && pt < 40 ? pt : sizePt}};
  }

  async function checkSelection(manual) {
    if (!wordReady) return;
    try {
      await Word.run(async ctx => {
        const f = await selectedFormula(ctx);
        if (f && adoptNext) {
          // Wiederhergestellter Entwurf: die angeklickte Formel wird sein Ziel, die Eingabe bleibt
          adoptNext = false;
          setEditing(f.meta);
          saved = JSON.stringify([f.meta.tex.replace(/\s+/g, ''), f.meta.mode, f.meta.font, f.meta.pt]);
          savedState = f.meta;
          updateFrame();
          say(pasteMsg, 'Die angeklickte Formel wird durch die wiederhergestellte ersetzt, sobald „Formel in Word ersetzen“ gedrückt wird. „Abbrechen“ zeigt wieder die Formel aus dem Dokument.', true);
        } else if (f) {
          loadMeta(f.meta);
          setEditing(f.meta);
          markSaved();
          say(pasteMsg, f.kind === 'old'
            ? 'Alte Formel übernommen. Bitte kurz prüfen, bei Bedarf ändern und dann den großen Knopf drücken. Danach ist sie im neuen Format.'
            : 'Formel aus dem Dokument geladen. Nach dem Ändern „Formel in Word ersetzen“ drücken.', true);
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

  window.FeWord = {ooxml, editing: () => !!current};   // für bulk.js und nummern.js

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
      // Neue abgesetzte Formel auf Wunsch mit Gleichungsnummer (nummern.js)
      const num = !current && window.FeNummer ? FeNummer.wanted() : null;
      let replaced = false, numText = '';
      // Neue Formel im eigenen Absatz kommt gleich mit den Einzügen aus dem Lineal (nummern.js, lineal.js)
      await Word.run(async ctx => {
        const f = current ? await selectedFormula(ctx) : null;
        if (f) {
          const xml = ooxml(r, meta, null);
          const at = f.kind === 'new' ? f.pic.getRange('Whole').insertOoxml(xml, 'Replace') : f.range.insertOoxml(xml, 'Replace');
          replaced = true;
          if (window.FeNummer) await FeNummer.afterReplace(ctx, at);
        }
        else if (num) {
          const at = await FeNummer.beforeInsert(ctx, num);
          await FeNummer.afterInsert(ctx, at.insertOoxml(ooxml(r, meta, num, FeNummer.pPr()), 'Replace'), num);
        }
        else {
          const pPr = window.FeNummer ? await FeNummer.placeUnnumbered(ctx) : '';
          const at = ctx.document.getSelection().insertOoxml(ooxml(r, meta, null, pPr), 'Replace');
          if (window.FeNummer) await FeNummer.afterInsert(ctx, at, null);
        }
        await ctx.sync();
      });
      if (num) numText = FeNummer.done(num);
      setEditing(null);
      markSaved();
      sayStatus((replaced ? 'Formel im Dokument ersetzt.' : num ? 'Formel mit Nummer (' + num + ') eingefügt.' : 'Formel eingefügt.') + (numText ? ' ' + numText : ''), true);
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
  $('clear').addEventListener('click', () => { setEditing(null); saved = savedState = null; updateFrame(); });
  cancelBtn.addEventListener('click', cancelEdit);

  Office.onReady(info => {
    if (info.host !== Office.HostType.Word) { say(status, 'Dieses Add-in läuft nur in Word.'); return; }
    wordReady = true;
    offerDraft();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    checkSelection(false);
  });
})();
