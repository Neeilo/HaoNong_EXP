/**
 * 訂單彙總／撿貨清單工具 — 核心邏輯（與畫面、與 Excel 函式庫無關）
 *
 * 由原本的 order_tool.exe（Python／openpyxl／tkinter）逐項移植而來，
 * 取代 tools_260210_v4.xlsm 巨集（Run_OrderTool_MultiFiles）。
 *
 * 設計重點（與原工具相同）：
 *  - 用「表頭文字」動態比對欄位，不寫死欄號；CRM 匯出格式改版也不會悄悄錯位。
 *  - 每個檔案各自獨立解析。
 *  - 任何一個檔案有問題 → 不產生輸出檔，只產生錯誤報告。
 *
 * 輸入的 grid 是二維陣列：grid[0] 是表頭列，其後是資料列，
 * 儲存格值為 string / number / boolean / Date / null。
 */

export const APP_TITLE = '訂單彙總／撿貨清單工具';

/** [內部鍵, 中文名稱, CRM 匯出檔的表頭文字] */
export const REQUIRED_FIELDS = [
  ['order_no', '訂單編號', 'Order Number'],
  ['remark', '備註', 'Remark'],
  ['recipient_first', '收件人名', 'Fulfillment First Name'],
  ['recipient_last', '收件人姓', 'Fulfillment Last Name'],
  ['phone', '收件人電話', 'Fulfillment Phone'],
  ['address1', '收件地址', 'Fulfillment Address 1'],
  ['zip', '郵遞區號', 'Fulfillment Zipcode'],
  ['district', '地區', 'Fulfillment District'],
  ['province', '縣市', 'Fulfillment Province'],
  ['item_name', '品項名稱', 'Item Name'],
  ['item_variant', '品項規格', 'Item Variant'],
  ['item_price', '品項單價', 'Item Price'],
  ['qty', '數量', 'Quantity'],
  ['subtotal', '小計', 'Subtotal'],
  ['shipping_fee', '運費', 'Shipping Fee'],
  ['shipping_fee_discount', '運費折扣', 'Shipping Fee Discount'],
  ['order_discount', '訂單折扣', 'Order Discount'],
  ['total_tax', '稅金', 'Total Tax'],
  ['transaction_fee', '手續費', 'Transaction Fee'],
  ['credit_used', '已使用購物金', 'Credit Used'],
  ['total_amount', '訂單總額', 'Total Amount'],
  ['refunded_amount', '已退款金額', 'Refunded Amount'],
  ['total_weight', '總重量', 'Total Weight(grams)'],
];

export const FIELD_LABEL = Object.fromEntries(REQUIRED_FIELDS.map(([key, label]) => [key, label]));

/** 訂單層級的金額欄位（每張訂單只在第一列顯示）。 */
export const MONEY_KEYS = [
  'subtotal',
  'shipping_fee',
  'shipping_fee_discount',
  'order_discount',
  'total_tax',
  'transaction_fee',
  'credit_used',
  'total_amount',
  'refunded_amount',
  'total_weight',
];

export const RECIPIENT_SHEET_NAME = '收件人清單';
export const PICKLIST_SHEET_NAME = '撿貨清單';
export const LOG_SHEET_NAME = '處理紀錄';

export const RECIPIENT_SHEET_HEADERS = [
  '來源檔案', '訂單編號', '收件者', '收件人電話', '郵遞區號', '收件地址', '備註',
  ...MONEY_KEYS.map((k) => FIELD_LABEL[k]),
  '品項名稱（含規格）', '單價', '數量',
];
export const PICKLIST_SHEET_HEADERS = ['品項＋規格', '總數量'];
export const LOG_SHEET_HEADERS = ['檔案名稱', '訂單數', '品項列數'];

// ---------------------------------------------------------------- 文字工具

const pad = (n, w = 2) => String(n).padStart(w, '0');

/** 儲存格值 → 字串（盡量貼近 Python 的 str()）。呼叫前請先排除 null/undefined。 */
export function toText(v) {
  if (v instanceof Date) {
    // ExcelJS 以 UTC 表示「沒有時區的 Excel 時間」，所以用 getUTC*。
    return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())} ` +
      `${pad(v.getUTCHours())}:${pad(v.getUTCMinutes())}:${pad(v.getUTCSeconds())}`;
  }
  if (typeof v === 'boolean') return v ? 'True' : 'False';
  return String(v);
}

/** 對應 Python 的 repr()，用在錯誤訊息裡顯示「原始值」。 */
export function pyRepr(v) {
  if (v === null || v === undefined) return 'None';
  if (typeof v === 'string') {
    const quote = v.includes("'") && !v.includes('"') ? '"' : "'";
    const escaped = v
      .replace(/\\/g, '\\\\')
      .replace(/\n/g, '\\n')
      .replace(/\r/g, '\\r')
      .replace(/\t/g, '\\t')
      .split(quote).join('\\' + quote);
    return quote + escaped + quote;
  }
  return toText(v);
}

export function normalizeHeader(v) {
  if (v === null || v === undefined) return '';
  return toText(v).trim().replace(/\s+/g, ' ');
}

export function trimmed(v) {
  if (v === null || v === undefined) return '';
  return toText(v).trim();
}

/** 移除英文字母（CRM 的縣市／地區欄位會帶英文翻譯，例如「台北市 Taipei」）。 */
export function stripEnglishLetters(s) {
  if (!s) return '';
  return s.replace(/[A-Za-z]/g, '').replace(/\s+/g, ' ').trim();
}

/** 電話：去掉空白、-、()、+；+886／886／00886 開頭改回 0 開頭。 */
export function normalizePhone(s) {
  if (!s) return '';
  for (const ch of [' ', '-', '(', ')', '+']) s = s.split(ch).join('');
  if (s.startsWith('00886')) return '0' + s.slice(5);
  if (s.startsWith('886')) return '0' + s.slice(3);
  return s;
}

/** 品項名稱只取「*」之前的部分。 */
export function splitBeforeStar(s) {
  if (!s) return '';
  const idx = s.indexOf('*');
  return (idx >= 0 ? s.slice(0, idx) : s).trim();
}

// 數字片段允許以單一底線分隔（Python 的 float("1_000") 合法）
const DIGITS = '\\d+(?:_\\d+)*';
const NUMBER_RE = new RegExp(
  `^[+-]?(?:${DIGITS}\\.?(?:${DIGITS})?|\\.${DIGITS})(?:[eE][+-]?${DIGITS})?$`,
);
const UNICODE_DIGIT_RE = /\p{Nd}/gu;
const isDigitChar = (cp) => /^\p{Nd}$/u.test(String.fromCodePoint(cp));

/**
 * 把任何書寫系統的十進位數字（全形「１２」、阿拉伯數字等）換成 ASCII 0-9。
 * Python 的 float() 本來就接受這些，所以原工具也能讀懂全形數字。
 */
function asciiDigits(s) {
  return s.replace(UNICODE_DIGIT_RE, (ch) => {
    const cp = ch.codePointAt(0);
    if (cp >= 48 && cp <= 57) return ch;
    let start = cp; // 同一組數字是連續 10 碼，往回找到該組的 0
    while (isDigitChar(start - 1)) start--;
    return String((cp - start) % 10);
  });
}

/**
 * 回傳 [狀態, 數值]。狀態為 'blank'（空白）/ 'ok'（可轉數字）/ 'invalid'（有值但無法轉數字）。
 *
 * 與原工具的唯一差異：Python 的 float() 會把 "nan"、"inf" 當成數字，
 * 這裡一律視為無法辨識，避免把垃圾資料算進撿貨總數。
 */
export function parseNumber(v) {
  const s = trimmed(v);
  if (s === '') return ['blank', null];
  const cleaned = asciiDigits(s.replace(/,/g, ''));
  if (NUMBER_RE.test(cleaned)) return ['ok', Number(cleaned.replace(/_/g, ''))];
  return ['invalid', s];
}

// ---------------------------------------------------------------- 欄位對應

export function buildHeaderMap(headerRow) {
  const map = new Map();
  (headerRow || []).forEach((raw, i) => {
    const header = normalizeHeader(raw);
    if (header && !map.has(header)) map.set(header, i + 1); // 重複表頭取第一個
  });
  return map;
}

export function resolveRequiredColumns(headerMap) {
  const cols = {};
  const missing = [];
  for (const [key, label, headerText] of REQUIRED_FIELDS) {
    const col = headerMap.get(headerText);
    if (col === undefined) missing.push(label);
    else cols[key] = col;
  }
  return [cols, missing];
}

export function buildAddress(province, district, addr1) {
  return stripEnglishLetters(trimmed(province)) + stripEnglishLetters(trimmed(district)) + trimmed(addr1);
}

// ---------------------------------------------------------------- 單一檔案

export class FileResult {
  constructor(filename) {
    this.filename = filename;
    this.errors = [];
    this.rows = [];
    this.orderCount = 0;
    this.itemRowCount = 0;
  }
}

/**
 * 處理單一檔案的工作表。
 * @param {string} filename  顯示用檔名
 * @param {Array<Array<any>>} grid  二維陣列，grid[0] 為表頭
 * @param {string} sheetTitle  工作表名稱（錯誤訊息用）
 */
export function processSheet(filename, grid, sheetTitle) {
  const result = new FileResult(filename);

  const headerMap = buildHeaderMap(grid[0]);
  const [cols, missing] = resolveRequiredColumns(headerMap);

  if (missing.length) {
    const found = [...headerMap.keys()].sort().join('、') || '(無)';
    result.errors.push(
      `【${filename}】工作表「${sheetTitle}」缺少必要欄位：${missing.join('、')}\n  該檔案實際抓到的表頭：${found}`,
    );
    return result;
  }

  const lastRow = grid.length;
  if (lastRow < 2) return result;

  const cell = (r, key) => {
    const v = (grid[r - 1] || [])[cols[key] - 1];
    return v === undefined ? null : v;
  };

  // 第一輪：逐列檢查資料、彙整訂單層級資訊
  const states = new Map();
  const orderSeq = [];
  const rowItems = new Map();
  const rowAnomalies = [];

  for (let r = 2; r <= lastRow; r++) {
    const orderNo = trimmed(cell(r, 'order_no'));
    if (orderNo === '') continue;

    const itemNameRaw = trimmed(cell(r, 'item_name'));
    const [qtyStatus, qtyVal] = parseNumber(cell(r, 'qty'));
    const [priceStatus, priceVal] = parseNumber(cell(r, 'item_price'));

    if (itemNameRaw === '' && (qtyStatus !== 'blank' || priceStatus !== 'blank')) {
      rowAnomalies.push(`第 ${r} 列（訂單 ${orderNo}）：品項名稱空白，但單價/數量有值`);
    } else if (itemNameRaw !== '') {
      if (qtyStatus !== 'ok') {
        rowAnomalies.push(`第 ${r} 列（訂單 ${orderNo}）：數量無法辨識為數字，原始值=${pyRepr(cell(r, 'qty'))}`);
      }
      if (priceStatus === 'invalid') {
        rowAnomalies.push(`第 ${r} 列（訂單 ${orderNo}）：單價無法辨識為數字，原始值=${pyRepr(cell(r, 'item_price'))}`);
      }
    }

    rowItems.set(r, [
      splitBeforeStar(itemNameRaw),
      trimmed(cell(r, 'item_variant')),
      priceStatus === 'ok' ? priceVal : null,
      qtyStatus === 'ok' ? qtyVal : null,
    ]);

    const recipient = (trimmed(cell(r, 'recipient_last')) + trimmed(cell(r, 'recipient_first'))).trim();
    const phone = normalizePhone(trimmed(cell(r, 'phone')));
    const remark = trimmed(cell(r, 'remark'));
    const zipc = trimmed(cell(r, 'zip'));
    const address = buildAddress(cell(r, 'province'), cell(r, 'district'), cell(r, 'address1'));

    if (!states.has(orderNo)) {
      orderSeq.push(orderNo);
      states.set(orderNo, {
        recipient, phone, remark, zip: zipc, address,
        money: Object.fromEntries(MONEY_KEYS.map((k) => [k, cell(r, k)])),
      });
      continue;
    }

    // 同一張訂單的後續列：有值的欄位覆蓋前面的值（以最後一個非空值為準）
    const st = states.get(orderNo);
    if (recipient) st.recipient = recipient;
    if (phone) st.phone = phone;
    if (remark) st.remark = remark;
    if (zipc) st.zip = zipc;
    if (address) st.address = address;
    for (const k of MONEY_KEYS) {
      const v = cell(r, k);
      if (trimmed(v) !== '') st.money[k] = v;
    }
  }

  if (rowAnomalies.length) {
    result.errors.push(`【${filename}】資料異常：\n  ` + rowAnomalies.join('\n  '));
    return result;
  }

  const moneyValue = (raw) => {
    const [status, num] = parseNumber(raw);
    return status === 'ok' ? num : trimmed(raw);
  };

  // 第二輪：輸出列。每張訂單只有第一列帶收件人與金額資訊。
  const printed = new Set();
  for (let r = 2; r <= lastRow; r++) {
    const orderNo = trimmed(cell(r, 'order_no'));
    if (orderNo === '') continue;

    const st = states.get(orderNo);
    const isFirst = !printed.has(orderNo);
    printed.add(orderNo);

    const base = isFirst
      ? [result.filename, orderNo, st.recipient, st.phone, st.zip, st.address, st.remark,
        ...MONEY_KEYS.map((k) => moneyValue(st.money[k]))]
      : [result.filename, orderNo, '', '', '', '', '', ...MONEY_KEYS.map(() => '')];

    const [itemClean, variant, priceVal, qtyVal] = rowItems.get(r);
    const itemDisplay = variant ? `${itemClean} | ${variant}` : itemClean;

    result.rows.push([...base, itemDisplay, priceVal, qtyVal]);
    result.itemRowCount += 1;
  }
  result.orderCount = orderSeq.length;
  return result;
}

// ---------------------------------------------------------------- 彙總

/** 撿貨清單：同品項＋規格加總數量，由多到少排序（同數量維持出現順序）。 */
export function buildPicklist(allRows) {
  const totals = new Map();
  for (const row of allRows) {
    const qty = row[row.length - 1];
    const itemDisplay = trimmed(row[row.length - 3]);
    if (itemDisplay === '' || typeof qty !== 'number' || qty === 0) continue;
    totals.set(itemDisplay, (totals.get(itemDisplay) || 0) + qty);
  }
  return [...totals.entries()].sort((a, b) => b[1] - a[1]);
}

/**
 * 合併多個檔案的處理結果。
 * 只要有任何錯誤，就只回傳 errors（不產生輸出）。
 */
export function combineResults(results) {
  const errors = results.flatMap((res) => res.errors);
  if (errors.length) return { errors, rows: [], picklist: [], fileStats: [], summary: null };

  const rows = results.flatMap((res) => res.rows);
  const fileStats = results.map((res) => [res.filename, res.orderCount, res.itemRowCount]);
  return {
    errors,
    rows,
    picklist: buildPicklist(rows),
    fileStats,
    summary: {
      fileCount: results.length,
      orderCount: fileStats.reduce((n, s) => n + s[1], 0),
      itemCount: fileStats.reduce((n, s) => n + s[2], 0),
    },
  };
}

// ---------------------------------------------------------------- 輸出檔名／報告

/** 本地時間 YYYYMMDD_HHMMSS */
export function timestamp(d = new Date()) {
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_` +
    `${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

export const outputFileName = (ts) => `收件人清單_撿貨清單_${ts}.xlsx`;
export const errorReportFileName = (ts) => `error_report_${ts}.txt`;

export function buildErrorReport(errors, ts) {
  return `${APP_TITLE}－錯誤報告\n產生時間：${ts}\n\n` + errors.join('\n\n');
}
