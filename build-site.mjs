import { mkdir, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { join, relative, sep } from 'node:path';

const distRoot = 'dist';
const routePrefix = '/Survey_breakdowm';
const mime = {
  '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.svg': 'image/svg+xml',
  '.woff': 'font/woff', '.woff2': 'font/woff2'
};
async function walk(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const source = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...await walk(source));
    else files.push(source);
  }
  return files;
}
const required = ['index.html', 'style.css', 'chart-rules.js', 'data-dictionary.js', 'linked-survey.js', 'script.js', 'data-io.js', 'survey-core.js', 'report-worker.js', 'data-worker.js'];
const sources = [];
for (const file of required) { await readFile(file); sources.push(file); }
for (const directory of ['assets', 'vendor']) { try { sources.push(...await walk(directory)); } catch { /* optional */ } }
await rm(distRoot, { recursive: true, force: true });
await mkdir(join(distRoot, 'server'), { recursive: true });
const assetEntries = [];
for (const source of sources) {
  const route = `/${relative('.', source).split(sep).join('/')}`;
  const content = await readFile(source);
  const extension = route.slice(route.lastIndexOf('.')).toLowerCase();
  const entry = { content: content.toString('base64'), contentType: mime[extension] || 'application/octet-stream' };
  const target = join(distRoot, route.slice(1).replaceAll('/', sep));
  await mkdir(join(target, '..'), { recursive: true });
  await writeFile(target, content);
  assetEntries.push([route, entry]);
  if (route === '/index.html') assetEntries.push(['/', entry]);
}
const prefixed = assetEntries.map(([route, value]) => [`${routePrefix}${route === '/' ? '' : route}`, value]);
const allEntries = [...assetEntries, ...prefixed];
const worker = `const assets = new Map(${JSON.stringify(allEntries)});
function decodeBase64(value) { const raw = atob(value); const bytes = new Uint8Array(raw.length); for (let i = 0; i < raw.length; i += 1) bytes[i] = raw.charCodeAt(i); return bytes; }
function notFound() { return new Response('Not found', { status: 404, headers: { 'content-type': 'text/plain; charset=utf-8' } }); }
export default { async fetch(request) { const url = new URL(request.url); const pathname = url.pathname.replace(/\\/+$/, '') || '/'; const asset = assets.get(pathname) || assets.get(pathname + '/'); if (!asset) return notFound(); return new Response(decodeBase64(asset.content), { headers: { 'content-type': asset.contentType, 'cache-control': 'no-store', 'content-security-policy': "default-src 'self'; script-src 'self'; worker-src 'self' blob:; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self' https://docs.google.com https://*.googleusercontent.com;" } }); } };\n`;
await writeFile(join(distRoot, 'server', 'index.js'), worker, 'utf8');
