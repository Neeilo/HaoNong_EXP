import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  processSheet, combineResults, buildPicklist, parseNumber, normalizePhone,
  stripEnglishLetters, splitBeforeStar, pyRepr, buildErrorReport, timestamp,
  REQUIRED_FIELDS, MONEY_KEYS,
} from '../docs/core.js';

// ---- 黃金標準：fixtures/golden.json 的「預期輸出」是用原本 order_tool.exe 的 Python 邏輯，
// 對 52 組隨機生成的 CRM 檔案（含全形數字、重複表頭、缺欄位、各種異常資料）實際跑出來的。
test('輸出與原 order_tool.exe 的邏輯逐項一致（黃金標準）', () => {
  const scenarios = JSON.parse(readFileSync(new URL('./fixtures/golden.json', import.meta.url), 'utf8'));
  assert.ok(scenarios.length >= 50);
  for (const s of scenarios) {
    const got = combineResults(s.files.map((f) => processSheet(f.name, f.grid, f.sheet)));
    const actual = { errors: got.errors };
    if (!s.expected.errors.length) {
      Object.assign(actual, { rows: got.rows, picklist: got.picklist, fileStats: got.fileStats, summary: got.summary });
    }
    assert.deepEqual(actual, s.expected, `scenario ${s.id}`);
  }
});

// ---- 小工具
test('parseNumber：空白／可轉／無法辨識，且認得全形數字與千分位', () => {
  assert.deepEqual(parseNumber(null), ['blank', null]);
  assert.deepEqual(parseNumber('  '), ['blank', null]);
  assert.deepEqual(parseNumber('1,200'), ['ok', 1200]);
  assert.deepEqual(parseNumber('１２'), ['ok', 12]);
  assert.deepEqual(parseNumber(' 2.5 '), ['ok', 2.5]);
  assert.deepEqual(parseNumber('1e2'), ['ok', 100]);
  assert.deepEqual(parseNumber('abc'), ['invalid', 'abc']);
  assert.deepEqual(parseNumber('0x10'), ['invalid', '0x10']);
  assert.deepEqual(parseNumber('nan'), ['invalid', 'nan']); // 與原工具唯一的刻意差異
});

test('normalizePhone：去符號、+886／886／00886 改回 0 開頭', () => {
  assert.equal(normalizePhone('+886-912-345-678'), '0912345678');
  assert.equal(normalizePhone('00886 912 345 678'), '0912345678');
  assert.equal(normalizePhone('(02) 2345-6789'), '0223456789');
  assert.equal(normalizePhone(''), '');
});

test('地址與品項名稱的清理', () => {
  assert.equal(stripEnglishLetters('台北市 Taipei'), '台北市');
  assert.equal(splitBeforeStar('高山茶*禮盒'), '高山茶');
  assert.equal(splitBeforeStar('  蛋 '), '蛋');
});

test('pyRepr 與 Python repr() 一致（錯誤訊息的「原始值」）', () => {
  assert.equal(pyRepr(null), 'None');
  assert.equal(pyRepr('abc'), "'abc'");
  assert.equal(pyRepr("it's"), '"it\'s"');
  assert.equal(pyRepr(12), '12');
});

// ---- 組合行為
const HEADER = REQUIRED_FIELDS.map(([, , h]) => h);
const makeRow = (over) => HEADER.map((h) => (h in over ? over[h] : null));

test('同一張訂單：只有第一列帶收件人與金額，後續非空值覆蓋', () => {
  const grid = [
    HEADER,
    makeRow({ 'Order Number': 'A1', 'Item Name': '茶*盒', Quantity: 2, 'Fulfillment Last Name': '王', Subtotal: 100 }),
    // 後續列的非空值整個覆蓋（含收件人：每列重新組「姓＋名」，非空就取代；金額同理）
    makeRow({ 'Order Number': 'A1', 'Item Name': '蛋', Quantity: 1, 'Fulfillment Last Name': '王', 'Fulfillment First Name': '小明', Subtotal: 250 }),
  ];
  const res = processSheet('f.xlsx', grid, 'Sheet1');
  assert.equal(res.errors.length, 0);
  assert.equal(res.orderCount, 1);
  assert.equal(res.itemRowCount, 2);
  const subtotalCol = 7 + MONEY_KEYS.indexOf('subtotal');
  assert.equal(res.rows[0][2], '王小明');
  assert.equal(res.rows[0][subtotalCol], 250);
  assert.equal(res.rows[1][2], '');
  assert.equal(res.rows[1][subtotalCol], '');
});

test('缺少欄位：報錯並列出實際抓到的表頭，不產生任何輸出', () => {
  const grid = [HEADER.filter((h) => h !== 'Quantity'), ['x']];
  const res = processSheet('bad.xlsx', grid, 'Sheet1');
  assert.match(res.errors[0], /缺少必要欄位：數量/);
  assert.match(res.errors[0], /該檔案實際抓到的表頭：/);
  const combined = combineResults([res]);
  assert.equal(combined.rows.length, 0);
  assert.equal(combined.summary, null);
});

test('撿貨清單：同品項加總、數量 0／空白略過、由多到少排序', () => {
  const row = (name, qty) => ['f', 'o', ...Array(Object.keys(MONEY_KEYS).length + 5).fill(''), name, null, qty];
  const picked = buildPicklist([row('A', 2), row('B', 5), row('A', 4), row('C', 0), row('D', null), row('', 3)]);
  assert.deepEqual(picked, [['A', 6], ['B', 5]]);
});

test('錯誤報告格式與時間戳', () => {
  assert.equal(buildErrorReport(['e1', 'e2'], '20260101_010203'), '訂單彙總／撿貨清單工具－錯誤報告\n產生時間：20260101_010203\n\ne1\n\ne2');
  assert.match(timestamp(new Date(2026, 0, 2, 3, 4, 5)), /^20260102_030405$/);
});
