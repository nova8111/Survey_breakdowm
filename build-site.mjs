import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';

const files = {
  '/': { source: 'index.html', contentType: 'text/html; charset=utf-8' },
  '/index.html': { source: 'index.html', contentType: 'text/html; charset=utf-8' },
  '/style.css': { source: 'style.css', contentType: 'text/css; charset=utf-8' },
  '/chart-rules.js': { source: 'chart-rules.js', contentType: 'application/javascript; charset=utf-8' },
  '/data-dictionary.js': { source: 'data-dictionary.js', contentType: 'application/javascript; charset=utf-8' },
  '/linked-survey.js': { source: 'linked-survey.js', contentType: 'application/javascript; charset=utf-8' },
  '/script.js': { source: 'script.js', contentType: 'application/javascript; charset=utf-8' },
  '/data-worker.js': { source: 'data-worker.js', contentType: 'application/javascript; charset=utf-8' }
};

await rm('dist', { recursive: true, force: true });
await mkdir('dist/server', { recursive: true });

const assetEntries = await Promise.all(
  Object.entries(files).map(async ([route, asset]) => {
    const content = await readFile(asset.source, 'utf8');
    return [route, { content, contentType: asset.contentType }];
  })
);

const worker = `const assets = new Map(${JSON.stringify(assetEntries)});

function notFound() {
  return new Response('Not found', {
    status: 404,
    headers: { 'content-type': 'text/plain; charset=utf-8' }
  });
}

export default {
  async fetch(request) {
    const url = new URL(request.url);
    const pathname = url.pathname.endsWith('/') && url.pathname !== '/' ? url.pathname.slice(0, -1) : url.pathname;
    const asset = assets.get(pathname);

    if (!asset) return notFound();

    return new Response(asset.content, {
      headers: {
        'content-type': asset.contentType,
        'cache-control': 'no-store'
      }
    });
  }
};
`;

await writeFile('dist/server/index.js', worker);
