// Hilfsskript: setzt Formeln (JSON-Liste {name, tex, mode, font}) mit dem echten Editor und speichert PNGs.
//   node zeigen.mjs liste.json zielordner
import {chromium} from 'playwright';
import fs from 'fs';
import path from 'path';
import {fileURLToPath} from 'url';
const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.dirname(HERE);
const [listFile, outDir] = process.argv.slice(2);
const list = JSON.parse(fs.readFileSync(listFile, 'utf8'));
const types = {'.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css', '.json': 'application/json', '.woff2': 'font/woff2'};
const launch = {};
if (process.env.CHROMIUM) launch.executablePath = process.env.CHROMIUM;
const browser = await chromium.launch(launch);
const page = await browser.newPage();
page.on('pageerror', e => console.log('Seitenfehler:', e.message));
await page.route('**/*', route => {
  const u = new URL(route.request().url());
  if (u.hostname !== 'editor.test') return route.abort();
  const f = path.join(REPO, decodeURIComponent(u.pathname));
  if (!f.startsWith(REPO) || !fs.existsSync(f) || !fs.statSync(f).isFile()) return route.fulfill({status: 404, body: ''});
  route.fulfill({path: f, contentType: types[path.extname(f)] || 'application/octet-stream'});
});
await page.goto('https://editor.test/editor.html');
await page.waitForFunction(() => window.mf && document.getElementById('out').contentWindow.__mjReady === true, null, {timeout: 60000});
fs.mkdirSync(outDir, {recursive: true});
for (const x of list) {
  const url = await page.evaluate(async ({tex, mode, font: fid}) => {
    const f = document.createElement('iframe');
    f.style.cssText = 'position:absolute;left:-10000px;top:0;width:900px;height:400px;border:0;';
    f.srcdoc = frameDoc(fid);
    document.body.appendChild(f);
    await new Promise(r => { f.onload = r; });
    const w = f.contentWindow;
    for (let i = 0; i < 300 && !w.__mjReady; i++) await new Promise(r => setTimeout(r, 100));
    const d = render(tex, mode, true);
    w.__show(d, d, mode);
    const r = await w.__png(12, 300, 0);
    f.remove();
    return r.dataUrl;
  }, {tex: x.tex, mode: x.mode || 'sci', font: x.font || 'stix2'});
  fs.writeFileSync(path.join(outDir, x.name + '.png'), Buffer.from(url.split(',')[1], 'base64'));
}
await browser.close();
