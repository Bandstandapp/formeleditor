/* „Alle Formeln in neues Format umwandeln“: holt das ganze Dokument, wandelt alle alten Formeln um
   (FeConvert) und öffnet das Ergebnis als neues Dokument. Danach zeigt eine Übersicht, was gefunden
   und umgewandelt wurde; Formeln mit Problemen sind im neuen Dokument gelb markiert. */
(function () {
  const sec = document.createElement('section');
  sec.className = 'panel'; sec.setAttribute('aria-label', 'Dokument'); sec.style.order = '5';
  sec.innerHTML = `<p class="label">Ganzes Dokument</p>
  <p class="hint">Wandelt alle alten Formeln dieses Dokuments in einem Schritt um, in der oben gewählten Schrift. Das Ergebnis öffnet sich als neues Dokument. Das Original bleibt unverändert.</p>
  <div class="row"><button class="btn primary" type="button" id="fe-all">Alle Formeln in neues Format umwandeln</button></div>
  <progress id="fe-prog" max="1" value="0" hidden style="width:100%"></progress>
  <p class="hint" id="fe-allmsg" role="status"></p>
  <div id="fe-report" hidden></div>`;
  (document.querySelector('.wrap') || document.body).appendChild(sec);
  const css = document.createElement('style');
  css.textContent = `
#fe-report { border: 2px solid var(--accent); border-radius: var(--r); padding: 10px 12px; display: grid; gap: 8px; }
#fe-report h3 { margin: 0; font-family: var(--f-head); font-size: var(--s-2); }
#fe-report table { border-collapse: collapse; width: 100%; }
#fe-report td { padding: 3px 0; vertical-align: top; }
#fe-report td.n { text-align: right; font-weight: 700; padding-left: 10px; white-space: nowrap; }
#fe-report tr.sum td { border-top: 1px solid var(--line); padding-top: 6px; }
#fe-report td.sub { padding-left: 14px; color: var(--muted); }
#fe-report mark { background: #fff200; color: #111; padding: 0 3px; }
#fe-report ul { margin: 0; padding-left: 18px; }`;
  document.head.appendChild(css);

  const $ = id => document.getElementById(id);
  const btn = $('fe-all'), prog = $('fe-prog'), msg = $('fe-allmsg'), box = $('fe-report');
  const esc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  let jszip = null;

  function loadJSZip() {
    if (window.JSZip) return Promise.resolve(window.JSZip);
    if (jszip) return jszip;
    return (jszip = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      s.onload = () => res(window.JSZip);
      s.onerror = () => { jszip = null; rej(new Error('Hilfsprogramm zum Entpacken lädt nicht. Bitte Internet prüfen.')); };
      document.head.appendChild(s);
    }));
  }

  // Das ganze Dokument als .docx-Bytes, so wie es gerade in Word steht
  function documentBytes() {
    return new Promise((res, rej) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, {sliceSize: 4194304}, r => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return rej(new Error(r.error.message));
        const file = r.value, chunks = [];
        const next = i => file.getSliceAsync(i, s => {
          if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); return rej(new Error(s.error.message)); }
          chunks.push(s.value.data);
          if (i + 1 < file.sliceCount) return next(i + 1);
          file.closeAsync();
          const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
          let pos = 0;
          chunks.forEach(c => { out.set(c, pos); pos += c.length; });
          res(out);
        });
        next(0);
      });
    });
  }

  // Formel setzen wie beim normalen Einfügen, nur mit kleinerem Ersatzbild (die SVG ist maßgeblich)
  async function renderFormula(tex, m, pt, padEm) {
    const w = frame.contentWindow;
    for (let i = 0; i < 200 && !w.__mjReady; i++) await sleep(100);
    const disp = render(tex, m, true);
    w.__show(disp, disp);
    return w.__png(pt, 300, padEm);
  }

  function showReport(rep) {
    const fmts = Object.entries(rep.formats);
    let h = '<h3>Ergebnis der Umwandlung</h3><table>';
    h += '<tr><td>Gefundene Formeln</td><td class="n">' + rep.total + '</td></tr>';
    h += '<tr><td>Unterschiedliche Ausgangsformate</td><td class="n">' + fmts.length + '</td></tr>';
    fmts.forEach(([k, v]) => { h += '<tr><td class="sub">' + esc(k) + '</td><td class="n">' + v + '</td></tr>'; });
    h += '<tr class="sum"><td>Erfolgreich umgewandelt</td><td class="n">' + rep.converted + '</td></tr>';
    h += '<tr><td>Mit Problemen, <mark>gelb markiert</mark></td><td class="n">' + rep.problems.length + '</td></tr>';
    if (rep.unchanged) h += '<tr><td>Unverändert gelassen</td><td class="n">' + rep.unchanged + '</td></tr>';
    h += '</table>';
    if (rep.problems.length) {
      h += '<p class="hint">Diese Formeln bitte im neuen Dokument neu eingeben (anklicken, im Editor eingeben, „Formel in Word ersetzen“):</p><ul>';
      rep.problems.forEach(p => { h += '<li>Formel ' + p.number + ' (' + esc(p.where) + '): ' + esc(p.why) + '</li>'; });
      h += '</ul>';
    }
    box.innerHTML = h;
    box.hidden = false;
  }

  async function convertAll() {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.document) { msg.textContent = 'Das geht nur, wenn der Editor in Word geöffnet ist.'; return; }
    btn.disabled = true; box.hidden = true;
    prog.hidden = false; prog.value = 0;
    try {
      msg.textContent = 'Dokument wird gelesen …';
      const [JSZip, bytes] = await Promise.all([loadJSZip(), documentBytes()]);
      const zip = await JSZip.loadAsync(bytes);
      const rep = await FeConvert.convertDocx(zip, renderFormula, {
        font,
        onProgress: (k, n) => { prog.max = n; prog.value = k; msg.textContent = 'Formel ' + k + ' von ' + n + ' wird umgewandelt …'; },
      });
      if (!rep.total) { msg.textContent = 'In diesem Dokument wurden keine Formeln gefunden.'; return; }
      if (!rep.converted) {
        msg.textContent = 'Es gab nichts umzuwandeln, deshalb wurde kein neues Dokument erstellt.';
        showReport(rep);
        return;
      }
      msg.textContent = 'Neues Dokument wird geöffnet …';
      const b64 = await zip.generateAsync({type: 'base64', compression: 'DEFLATE'});
      await Word.run(async ctx => {
        const d = ctx.application.createDocument(b64);
        await ctx.sync();
        d.open();
        await ctx.sync();
      });
      msg.textContent = 'Fertig. Das umgewandelte Dokument ist in einem neuen Fenster geöffnet. Bitte dort mit „Ablage“ → „Speichern unter …“ unter neuem Namen sichern.';
      showReport(rep);
    } catch (e) {
      msg.textContent = 'Die Umwandlung hat nicht geklappt: ' + (e && e.message ? e.message : e);
    } finally {
      btn.disabled = false; prog.hidden = true;
      try { show(); } catch (e) {}       // Vorschau wieder auf die Formel im Editor stellen
    }
  }

  btn.addEventListener('click', convertAll);
})();
