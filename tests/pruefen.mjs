// Prüflauf vor jeder neuen Version: setzt Formeln mit genau dem Code, den der Editor benutzt
// (editor.html aus dem Repo, MathJax und MathLive aus vendor/), und vergleicht mit dem letzten guten Stand.
//
//   node pruefen.mjs                         neutrale Beispiele (beispiele.json), Vergleich mit stand-beispiele.json
//   node pruefen.mjs A.docx B.docx --stand stand.json --bericht bericht.html
//                                            alle Formeln aus Word-Dokumenten (alte und neue), z. B. Papas Manuskripte
//   --neu                                    aktuellen Stand als neuen guten Stand speichern (nach gewollten Änderungen)
//
// Geprüft wird je Formel:
//   1. Sie lässt sich setzen (kein MathJax-Fehler), wissenschaftlich und einzeilig.
//   2. Öffnen im Editor ändert sie nicht: MathLive liest den Quelltext ein und gibt ihn zurück,
//      das gesetzte Bild muss gleich bleiben (sonst ändert schon „Ohne Änderung übernehmen“ die Formel).
//   3. Das gesetzte Bild ist dasselbe wie beim letzten guten Stand.
// Ergebnis: Zusammenfassung in der Konsole, auf Wunsch ein HTML-Bericht mit Vorher/Nachher-Bildern.
// Rückgabewert 1, wenn etwas nicht stimmt (für die automatische Prüfung auf GitHub).
import {chromium} from 'playwright';
import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import {createRequire} from 'module';
import {fileURLToPath} from 'url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.dirname(HERE);
const require = createRequire(import.meta.url);
const JSZip = require('jszip');
const Eqn3 = require(path.join(REPO, 'eqn3.js'));
const Omml = require(path.join(REPO, 'omml.js'));
globalThis.Omml = Omml;
globalThis.MathML = require(path.join(REPO, 'mathml.js'));
const FeConvert = require(path.join(REPO, 'convert.js'));

const args = process.argv.slice(2);
const opt = name => { const i = args.indexOf(name); if (i < 0) return null; const v = args[i + 1]; args.splice(i, 2); return v; };
const flag = name => { const i = args.indexOf(name); if (i < 0) return false; args.splice(i, 1); return true; };
const update = flag('--neu');
const reportPath = opt('--bericht');
let standPath = opt('--stand');
const docs = args.filter(a => /\.docx$/i.test(a));
if (!standPath) standPath = docs.length ? null : path.join(HERE, 'stand-beispiele.json');
if (docs.length && !standPath) { console.error('Bei Word-Dokumenten bitte --stand <datei.json> angeben (außerhalb des Repos, die Formeln sind privat).'); process.exit(2); }

// ---------- Formeln einsammeln ----------
async function fromDocx(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const {items} = await FeConvert.scan(zip);
  const out = [];
  const relsCache = {};
  const rels = async part => {
    const name = part.replace(/^word\//, 'word/_rels/') + '.rels';
    if (!relsCache[name]) relsCache[name] = (zip.file(name) ? await zip.file(name).async('string') : '');
    return relsCache[name];
  };
  for (const f of items) {
    const id = path.basename(file) + ' #' + f.number;
    if (f.kind === 'new') { out.push({id, tex: f.meta.tex, mode: f.meta.mode || 'sci', where: f.where}); continue; }
    if (f.kind === 'omml') {                                   // Formeln aus dem heutigen Formel-Editor von Word
      try { out.push({id, tex: Omml.toLatex(f.obj), mode: 'sci', where: f.where}); }
      catch (e) { out.push({id, tex: null, error: 'nicht lesbar: ' + e.message, where: f.where}); }
      continue;
    }
    if (f.kind === 'mt') {                                     // Formeln aus dem heutigen MathType
      try { out.push({id, tex: MathML.toLatex(f.mathml), mode: 'sci', where: f.where}); }
      catch (e) { out.push({id, tex: null, error: 'nicht lesbar: ' + e.message, where: f.where}); }
      continue;
    }
    if (f.kind !== 'old') continue;
    const rid = (f.obj.match(/<o:OLEObject\b[^>]*\br:id="([^"]+)"/) || [])[1];
    const rel = rid && (await rels(f.part.name)).match(new RegExp('<Relationship [^>]*Id="' + rid + '"[^>]*>'));
    const target = rel && (rel[0].match(/Target="([^"]+)"/) || [])[1];
    try {
      const bin = await zip.file('word/' + target.replace(/^\/?(word\/)?/, '')).async('uint8array');
      out.push({id, tex: Eqn3.oleToLatex(bin), mode: 'sci', where: f.where});
    } catch (e) { out.push({id, tex: null, error: 'nicht lesbar: ' + e.message, where: f.where}); }
  }
  return out;
}
let formulas = [];
if (docs.length) for (const d of docs) formulas.push(...await fromDocx(d));
else formulas = JSON.parse(fs.readFileSync(path.join(HERE, 'beispiele.json'), 'utf8')).map((x, i) => ({id: 'Beispiel ' + (i + 1) + ': ' + x.name, tex: x.tex, mode: x.mode || 'sci'}));

// ---------- Setzen im Browser ----------
const types = {'.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css', '.json': 'application/json', '.woff2': 'font/woff2'};
const launch = {};
if (process.env.CHROMIUM) launch.executablePath = process.env.CHROMIUM;
const browser = await chromium.launch(launch);
const page = await browser.newPage();
page.on('pageerror', e => console.log('Seitenfehler:', e.message));
await page.route('**/*', route => {
  const u = new URL(route.request().url());
  if (u.hostname !== 'editor.test') return route.abort();     // nichts von fremden Servern
  const f = path.join(REPO, decodeURIComponent(u.pathname));
  if (!f.startsWith(REPO) || !fs.existsSync(f) || !fs.statSync(f).isFile()) return route.fulfill({status: 404, body: ''});
  route.fulfill({path: f, contentType: types[path.extname(f)] || 'application/octet-stream'});
});
await page.goto('https://editor.test/editor.html');
await page.waitForFunction(() => window.mf && document.getElementById('out').contentWindow.__mjReady === true, null, {timeout: 60000});

const results = [];
for (const f of formulas) {
  if (f.tex == null) { results.push({...f, problems: [f.error]}); continue; }
  const r = await page.evaluate(async ({tex, mode}) => {
    const MJ = document.getElementById('out').contentWindow.MathJax;
    // Was man sieht: jedes Zeichen und jeder Strich mit seiner Lage auf dem Blatt. Der innere Aufbau
    // des Bildes darf sich ändern, solange das Ergebnis gleich aussieht.
    const look = svg => {
      const host = document.createElement('div');
      host.style.cssText = 'position:absolute;left:-9999px;top:0';
      const c = svg.cloneNode(true); host.appendChild(c); document.body.appendChild(host);
      // Maßstab 1:1 zum viewBox, damit Schriftgröße und Seitenstil der Testseite nicht mitzählen
      const vb = (c.getAttribute('viewBox') || '0 0 1 1').split(/\s+/);
      c.removeAttribute('style'); c.setAttribute('width', vb[2]); c.setAttribute('height', vb[3]);
      const r = n => Math.round(n * 10) / 10;
      const parts = [c.getAttribute('viewBox')];
      c.querySelectorAll('use, rect, path, line, polygon').forEach(el => {
        if (el.closest('defs')) return;
        const m = el.getCTM && el.getCTM();
        const box = el.tagName === 'rect' ? [el.getAttribute('x'), el.getAttribute('y'), el.getAttribute('width'), el.getAttribute('height')].join(',') : '';
        const fill = el.getAttribute('fill') || (el.closest('[fill]') || {getAttribute: () => ''}).getAttribute('fill') || '';
        parts.push([el.tagName, (el.getAttribute('xlink:href') || el.getAttribute('href') || el.getAttribute('d') || '').replace(/MJX-\d+-/g, 'MJX-'), box, fill,
          m ? [m.a, m.b, m.c, m.d, m.e, m.f].map(r).join(',') : ''].join('|'));
      });
      host.remove();
      return parts.join('\n');
    };
    const set = async t => {
      try {
        const node = await MJ.tex2svgPromise(t, {display: false});
        const err = node.querySelector('[data-mjx-error]');
        const svg = node.querySelector('svg');
        return {svg: svg ? svg.outerHTML.replace(/MJX-\d+-/g, 'MJX-') : '', look: svg ? look(svg) : '', err: err ? err.getAttribute('data-mjx-error') : null};
      } catch (e) { return {svg: '', err: String(e.message || e)}; }
    };
    const main = await set(render(tex, mode, true));
    const lin = await set(render(tex, 'lin', false));
    mf.setValue(tex, {silenceNotifications: true});
    const back = mf.getValue('latex');
    const again = await set(render(back, mode, true));
    return {svg: main.svg, look: main.look, err: main.err, linErr: lin.err, back, roundtrip: again.look === main.look};
  }, {tex: f.tex, mode: ['sci', 'kom', 'lin'].includes(f.mode) ? f.mode : 'sci'});
  const problems = [];
  if (r.err) problems.push('lässt sich nicht setzen: ' + r.err);
  if (r.linErr) problems.push('einzeilig nicht setzbar: ' + r.linErr);
  if (!r.err && !r.roundtrip) problems.push('ändert sich beim Öffnen im Editor (' + r.back + ')');
  results.push({...f, svg: r.svg, hash: crypto.createHash('sha1').update(r.look).digest('hex').slice(0, 16), problems});
}
await browser.close();

// ---------- Vergleich mit dem letzten guten Stand ----------
const stand = standPath && fs.existsSync(standPath) ? JSON.parse(fs.readFileSync(standPath, 'utf8')) : null;
const old = new Map(stand ? stand.formeln.map(x => [x.id + '\u0000' + x.tex, x]) : []);
let changed = 0, fresh = 0;
for (const x of results) {
  if (!x.svg) continue;
  const o = old.get(x.id + '\u0000' + x.tex);
  if (!o) { fresh++; continue; }
  if (o.hash !== x.hash) { changed++; x.before = o.svg; x.problems.push('sieht anders aus als beim letzten guten Stand'); }
}
const bad = results.filter(x => x.problems.length);
const roundtrip = results.filter(x => x.problems.some(p => p.startsWith('ändert sich'))).length;
console.log('Formeln geprüft:            ' + results.length);
console.log('Lassen sich nicht setzen:   ' + results.filter(x => x.problems.some(p => /setz|lesbar/.test(p))).length);
console.log('Ändern sich beim Öffnen:    ' + roundtrip);
console.log('Anders als guter Stand:     ' + (stand ? changed : '– (noch kein guter Stand gespeichert)'));
if (stand && fresh) console.log('Neu (nicht im guten Stand): ' + fresh);
bad.slice(0, 30).forEach(x => console.log('  ' + x.id + (x.where ? ' (' + x.where + ')' : '') + ': ' + x.problems.join('; ')));
if (bad.length > 30) console.log('  … und ' + (bad.length - 30) + ' weitere, siehe Bericht');

if (reportPath) {
  const esc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const img = svg => svg ? '<div class="f">' + svg.replace(/currentColor/g, '#111') + '</div>' : '<div class="f">–</div>';
  const rows = bad.map(x => '<tr><td><b>' + esc(x.id) + '</b><br><small>' + esc(x.where || '') + '</small></td><td>' + esc(x.problems.join('; ')) +
    '<br><code>' + esc(x.tex || '') + '</code></td><td>' + (x.before ? img(x.before) : '') + '</td><td>' + img(x.svg) + '</td></tr>').join('\n');
  fs.writeFileSync(reportPath, '<!doctype html><meta charset="utf-8"><title>Prüflauf Formeleditor</title><style>body{font:15px system-ui,sans-serif;margin:24px;color:#111}' +
    'table{border-collapse:collapse;width:100%}td{border-top:1px solid #ccc;padding:8px;vertical-align:top}.f{font-size:22px;overflow-x:auto;max-width:420px}' +
    'code{font-size:12px;color:#555;word-break:break-all}</style><h1>Prüflauf Formeleditor</h1><p>' + results.length + ' Formeln geprüft, ' + bad.length +
    ' mit Auffälligkeiten.</p><table><tr><th>Formel</th><th>Befund</th><th>Vorher</th><th>Jetzt</th></tr>' + rows + '</table>');
  console.log('Bericht: ' + reportPath);
}
if (update) {
  fs.writeFileSync(standPath, JSON.stringify({datum: new Date().toISOString().slice(0, 10), formeln: results.filter(x => x.svg).map(x => ({id: x.id, tex: x.tex, hash: x.hash, svg: x.svg}))}));
  console.log('Neuer guter Stand gespeichert: ' + standPath);
}
const failed = results.some(x => x.problems.some(p => !p.startsWith('nicht lesbar'))) && !update;
process.exit(failed ? 1 : 0);
