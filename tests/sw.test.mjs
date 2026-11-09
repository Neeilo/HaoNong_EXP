import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync, statSync, existsSync } from 'node:fs';
import { join, relative } from 'node:path';

const docs = new URL('../docs/', import.meta.url).pathname;
const read = (f) => readFileSync(join(docs, f), 'utf8');

function walk(dir) {
  return readdirSync(dir).flatMap((n) => {
    const p = join(dir, n);
    return statSync(p).isDirectory() ? walk(p) : [p];
  });
}

const sw = read('sw.js');
const assets = [...sw.match(/const ASSETS = \[([\s\S]*?)\];/)[1].matchAll(/'([^']+)'/g)].map((m) => m[1]);

test('sw.js 預先快取清單涵蓋 docs/ 內每個檔案（除了 sw.js 與授權檔）', () => {
  const onDisk = walk(docs).map((p) => relative(docs, p)).filter((p) => p !== 'sw.js' && !p.startsWith('vendor/LICENSE'));
  const listed = assets.filter((a) => a !== './');
  assert.deepEqual([...listed].sort(), [...onDisk].sort());
});

test('預先快取的每個檔案都存在', () => {
  for (const a of assets) if (a !== './') assert.ok(existsSync(join(docs, a)), a);
});

test('sw.js VERSION 與 app.js APP_VERSION 一致', () => {
  const swV = sw.match(/const VERSION = '([^']+)'/)[1];
  const appV = read('app.js').match(/const APP_VERSION = '([^']+)'/)[1];
  assert.equal(swV, appV);
});

test('manifest 的圖示都存在，且 start_url／scope 為相對路徑（可部署在子路徑）', () => {
  const m = JSON.parse(read('manifest.webmanifest'));
  for (const i of m.icons) assert.ok(existsSync(join(docs, i.src)), i.src);
  assert.equal(m.start_url, './');
  assert.equal(m.scope, './');
  assert.ok(m.icons.some((i) => i.purpose === 'maskable'));
  assert.ok(m.icons.some((i) => i.sizes === '192x192') && m.icons.some((i) => i.sizes === '512x512'));
});
