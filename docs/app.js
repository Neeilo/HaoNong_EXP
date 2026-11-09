import {
  FileResult, processSheet, combineResults, timestamp,
  outputFileName, errorReportFileName, buildErrorReport,
  RECIPIENT_SHEET_NAME, PICKLIST_SHEET_NAME, LOG_SHEET_NAME,
  RECIPIENT_SHEET_HEADERS, PICKLIST_SHEET_HEADERS, LOG_SHEET_HEADERS,
} from './core.js';
import { readWorkbook, buildOutputWorkbook } from './xlsx-io.js';

const APP_VERSION = '1.0.0'; // 必須與 sw.js 的 VERSION 一致（tests/sw.test.mjs 會檢查）

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const PREVIEW_ROW_LIMIT = 100;

const $ = (id) => document.getElementById(id);
const els = {
  input: $('file-input'), dropzone: $('dropzone'), list: $('file-list'),
  run: $('run-btn'), clear: $('clear-btn'), status: $('status'),
  result: $('result'), resultBody: $('result-body'),
  install: $('install-btn'), offline: $('offline-badge'), iosHint: $('ios-hint'),
  updateBanner: $('update-banner'), updateBtn: $('update-btn'), version: $('version'),
};

/** 建立 DOM 的小工具。文字一律用 textContent / 文字節點，不會被當成 HTML。 */
function h(tag, props = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v === false || v === null || v === undefined) continue;
    if (k === 'class') node.className = v;
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else node.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat()) {
    if (c === false || c === null || c === undefined) continue; // 讓 `cond && h(...)` 可以直接使用
    node.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return node;
}

// ------------------------------------------------------------------ 檔案清單

/** @type {File[]} */
let files = [];
let busy = false;

const fileKey = (f) => `${f.name}|${f.size}|${f.lastModified}`;
const isExcel = (f) => /\.(xlsx|xlsm|xls)$/i.test(f.name);
const formatSize = (n) => (n >= 1048576 ? `${(n / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(n / 1024))} KB`);

function setStatus(text, isError = false) {
  els.status.textContent = text;
  els.status.classList.toggle('error', isError);
}

function addFiles(incoming) {
  const picked = [...incoming];
  const accepted = picked.filter(isExcel);
  const ignored = picked.length - accepted.length;
  const known = new Set(files.map(fileKey));
  for (const f of accepted) if (!known.has(fileKey(f))) files.push(f);
  renderFileList();
  setStatus(ignored ? `已略過 ${ignored} 個不是 Excel 的檔案。` : '', ignored > 0);
}

function removeFile(index) {
  files.splice(index, 1);
  renderFileList();
  setStatus('');
}

function renderFileList() {
  els.list.replaceChildren(
    ...files.map((f, i) => h('li', {},
      h('span', { class: 'file-name' }, f.name, h('span', { class: 'file-size' }, formatSize(f.size))),
      h('button', { class: 'btn', type: 'button', 'aria-label': `移除 ${f.name}`, onclick: () => removeFile(i), disabled: busy }, '移除'))),
  );
  els.run.disabled = busy || files.length === 0;
  els.clear.hidden = files.length === 0;
  els.clear.disabled = busy;
}

els.input.addEventListener('change', () => {
  addFiles(els.input.files);
  els.input.value = ''; // 允許再次選到同一個檔案
});
els.clear.addEventListener('click', () => {
  files = [];
  renderFileList();
  setStatus('');
  els.result.hidden = true;
});

for (const type of ['dragenter', 'dragover']) {
  els.dropzone.addEventListener(type, (e) => { e.preventDefault(); els.dropzone.classList.add('drag'); });
}
for (const type of ['dragleave', 'drop']) {
  els.dropzone.addEventListener(type, (e) => { e.preventDefault(); els.dropzone.classList.remove('drag'); });
}
els.dropzone.addEventListener('drop', (e) => { if (e.dataTransfer?.files) addFiles(e.dataTransfer.files); });

// ------------------------------------------------------------------ 處理流程

async function processFile(file) {
  try {
    const { sheetTitle, grid } = await readWorkbook(file);
    return processSheet(file.name, grid, sheetTitle);
  } catch (err) {
    const res = new FileResult(file.name);
    res.errors.push(`【${file.name}】無法讀取檔案：${err?.message || err}`);
    return res;
  }
}

async function run() {
  if (busy || files.length === 0) return;
  busy = true;
  renderFileList();
  els.result.hidden = true;

  try {
    const results = [];
    for (const [i, file] of files.entries()) {
      setStatus(`處理中（${i + 1}/${files.length}）：${file.name}`);
      results.push(await processFile(file));
    }

    const combined = combineResults(results);
    const ts = timestamp();

    if (combined.errors.length) {
      setStatus('');
      showErrors(combined.errors, ts);
    } else {
      setStatus('產生輸出檔中…');
      const buffer = await buildOutputWorkbook(combined);
      const blob = new Blob([buffer], { type: XLSX_MIME });
      setStatus('');
      showSuccess(combined, blob, outputFileName(ts));
    }
  } catch (err) {
    console.error(err);
    setStatus(`發生未預期的錯誤：${err?.message || err}`, true);
  } finally {
    busy = false;
    renderFileList();
  }
}
els.run.addEventListener('click', run);

// ------------------------------------------------------------------ 結果畫面

function showResult(...nodes) {
  els.resultBody.replaceChildren(...nodes);
  els.result.hidden = false;
  els.result.scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth', block: 'start' });
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = h('a', { href: url, download: name });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

function showErrors(errors, ts) {
  // 與原工具相同：有任何問題就不產生輸出檔，只提供錯誤報告。UTF-8 BOM 讓記事本等工具正確顯示中文。
  const report = new Blob(['﻿', buildErrorReport(errors, ts)], { type: 'text/plain;charset=utf-8' });
  showResult(
    h('div', { class: 'alert', role: 'alert' },
      h('strong', {}, `發現 ${errors.length} 個問題，未產生任何輸出檔。`),
      h('ul', { class: 'error-list' }, errors.map((e) => h('li', {}, e)))),
    h('div', { class: 'actions' },
      h('button', { class: 'btn primary', type: 'button', onclick: () => downloadBlob(report, errorReportFileName(ts)) }, '下載錯誤報告'),
      h('p', { class: 'hint' }, '修正檔案後，移除舊檔再重新選擇即可。')),
  );
}

function renderTable(headers, rows, { limit = Infinity, wrapCols = [] } = {}) {
  const shown = rows.slice(0, limit);
  const isNum = (v) => typeof v === 'number';
  const table = h('table', {},
    h('thead', {}, h('tr', {}, headers.map((t) => h('th', { scope: 'col' }, t)))),
    h('tbody', {}, shown.map((row) => h('tr', {}, row.map((v, c) =>
      h('td', { class: [isNum(v) ? 'num' : '', wrapCols.includes(c) ? 'wrap' : ''].join(' ').trim() || false },
        v === null || v === undefined ? '' : v))))));
  const nodes = [h('div', { class: 'table-wrap' }, table)];
  if (rows.length > shown.length) {
    nodes.push(h('p', { class: 'note' }, `僅預覽前 ${shown.length} 列（共 ${rows.length} 列），完整內容請下載 Excel。`));
  }
  if (rows.length === 0) nodes.push(h('p', { class: 'note' }, '（沒有資料）'));
  return nodes;
}

function showSuccess(combined, blob, fileName) {
  const { summary, rows, picklist, fileStats } = combined;
  const file = new File([blob], fileName, { type: XLSX_MIME });
  const canShare = !!(navigator.canShare && navigator.canShare({ files: [file] }));

  const tabs = [
    { id: 'picklist', label: PICKLIST_SHEET_NAME, count: picklist.length,
      body: () => renderTable(PICKLIST_SHEET_HEADERS, picklist, { wrapCols: [0] }) },
    { id: 'recipients', label: RECIPIENT_SHEET_NAME, count: rows.length,
      body: () => renderTable(RECIPIENT_SHEET_HEADERS, rows, { limit: PREVIEW_ROW_LIMIT }) },
    { id: 'log', label: LOG_SHEET_NAME, count: fileStats.length,
      body: () => renderTable(LOG_SHEET_HEADERS, fileStats) },
  ];

  const panel = h('div', { role: 'tabpanel', id: 'tab-panel' });
  const tablist = h('div', { class: 'tabs', role: 'tablist', 'aria-label': '結果預覽' });
  const tabButtons = tabs.map((t) => h('button', {
    class: 'tab', role: 'tab', type: 'button', id: `tab-${t.id}`, 'aria-controls': 'tab-panel',
    onclick: () => select(t.id),
  }, t.label, h('span', { class: 'tab-count' }, `(${t.count})`)));
  tablist.append(...tabButtons);

  function select(id, focus = false) {
    tabs.forEach((t, i) => {
      const on = t.id === id;
      tabButtons[i].setAttribute('aria-selected', String(on));
      tabButtons[i].tabIndex = on ? 0 : -1;
      if (on) {
        panel.setAttribute('aria-labelledby', `tab-${t.id}`);
        panel.replaceChildren(...t.body());
        if (focus) tabButtons[i].focus();
      }
    });
  }
  tablist.addEventListener('keydown', (e) => {
    const cur = tabButtons.findIndex((b) => b.getAttribute('aria-selected') === 'true');
    const next = { ArrowRight: cur + 1, ArrowLeft: cur - 1, Home: 0, End: tabs.length - 1 }[e.key];
    if (next === undefined) return;
    e.preventDefault();
    select(tabs[(next + tabs.length) % tabs.length].id, true);
  });
  select('picklist');

  showResult(
    h('div', { class: 'stats' },
      stat(summary.fileCount, '檔案'), stat(summary.orderCount, '訂單'), stat(summary.itemCount, '品項列')),
    h('div', { class: 'actions' },
      h('button', { class: 'btn primary', type: 'button', onclick: () => downloadBlob(blob, fileName) }, '下載 Excel'),
      canShare && h('button', {
        class: 'btn', type: 'button',
        onclick: () => navigator.share({ files: [file], title: fileName }).catch((e) => { if (e.name !== 'AbortError') console.error(e); }),
      }, '分享／傳送')),
    h('p', { class: 'note' }, fileName),
    tablist, panel,
  );
}

const stat = (n, label) => h('div', { class: 'stat' }, h('b', {}, n), h('span', {}, label));

// ------------------------------------------------------------------ PWA：安裝、離線、更新

els.version.textContent = `v${APP_VERSION}`;

let deferredInstall = null;
window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  deferredInstall = e;
  els.install.hidden = false;
});
els.install.addEventListener('click', async () => {
  if (!deferredInstall) return;
  deferredInstall.prompt();
  await deferredInstall.userChoice;
  deferredInstall = null;
  els.install.hidden = true;
});
window.addEventListener('appinstalled', () => { els.install.hidden = true; });

const isStandalone = matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;
const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
els.iosHint.hidden = !(isIOS && !isStandalone);

const syncOnline = () => { els.offline.hidden = navigator.onLine; };
window.addEventListener('online', syncOnline);
window.addEventListener('offline', syncOnline);
syncOnline();

if ('serviceWorker' in navigator) {
  let refreshing = false;
  let userRequestedUpdate = false;

  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (refreshing || !userRequestedUpdate) return; // 只在使用者按下「立即更新」後才重新載入，避免打斷處理中的工作
    refreshing = true;
    location.reload();
  });

  const offerUpdate = (worker) => {
    els.updateBanner.hidden = false;
    els.updateBtn.onclick = () => {
      userRequestedUpdate = true;
      worker.postMessage('SKIP_WAITING');
    };
  };

  navigator.serviceWorker.register('./sw.js').then((reg) => {
    if (reg.waiting && navigator.serviceWorker.controller) offerUpdate(reg.waiting);
    reg.addEventListener('updatefound', () => {
      const incoming = reg.installing;
      incoming?.addEventListener('statechange', () => {
        if (incoming.state === 'installed' && navigator.serviceWorker.controller) offerUpdate(incoming);
      });
    });
  }).catch((err) => console.warn('Service worker 註冊失敗：', err));
}

renderFileList();
