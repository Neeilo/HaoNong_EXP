/**
 * Excel 讀寫層（使用 ExcelJS，由 index.html 先載入成全域 ExcelJS）。
 * 所有處理都在瀏覽器本機完成，檔案不會上傳到任何地方。
 */
import {
  RECIPIENT_SHEET_NAME, PICKLIST_SHEET_NAME, LOG_SHEET_NAME,
  RECIPIENT_SHEET_HEADERS, PICKLIST_SHEET_HEADERS, LOG_SHEET_HEADERS,
} from './core.js';

const getExcelJS = () => {
  if (!globalThis.ExcelJS) throw new Error('Excel 元件尚未載入，請重新整理頁面後再試。');
  return globalThis.ExcelJS;
};

/** ExcelJS 的各種儲存格型別 → 單純的值（對應原工具 openpyxl data_only=True 的行為：公式取計算結果）。 */
export function normalizeCellValue(v) {
  if (v === null || v === undefined) return null;
  if (typeof v !== 'object' || v instanceof Date) return v;
  if ('result' in v) return normalizeCellValue(v.result); // 公式 → 快取的計算結果
  if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
  if ('error' in v) return String(v.error);
  if ('text' in v) return normalizeCellValue(v.text); // 超連結
  return String(v);
}

function sheetToGrid(ws) {
  const grid = [];
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const values = row.values; // 1-based，可能有洞
    const out = [];
    for (let c = 1; c < values.length; c++) out.push(normalizeCellValue(values[c]));
    grid[rowNumber - 1] = out;
  });
  for (let i = 0; i < grid.length; i++) if (!grid[i]) grid[i] = []; // 補上整列都是空的
  return grid;
}

/**
 * 讀取一個 Excel 檔：優先使用名為「Products」的工作表，否則用第一個工作表。
 * @returns {{sheetTitle: string, grid: any[][]}}
 */
export async function readWorkbook(file) {
  const ExcelJS = getExcelJS();
  if (/\.xls$/i.test(file.name)) {
    throw new Error('不支援舊版 .xls 格式，請先在 Excel 另存為 .xlsx 再匯入。');
  }
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(await file.arrayBuffer());
  } catch (err) {
    console.warn(err);
    throw new Error('檔案無法開啟，可能已損毀，或不是有效的 .xlsx / .xlsm 檔。');
  }
  const ws = wb.getWorksheet('Products') || wb.worksheets[0];
  if (!ws) throw new Error('檔案中找不到任何工作表。');
  return { sheetTitle: ws.title, grid: sheetToGrid(ws) };
}

const HEADER_FONT = { name: 'Arial', bold: true };
const BODY_FONT = { name: 'Arial' };

function styleSheet(ws) {
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.font = rowNumber === 1 ? HEADER_FONT : BODY_FONT;
    });
  });
  for (let c = 1; c <= ws.columnCount; c++) ws.getColumn(c).width = 18;
}

/**
 * 產生輸出的 .xlsx（三張工作表）。電話、郵遞區號欄位設為文字格式，避免開頭的 0 被吃掉。
 * @returns {Promise<ArrayBuffer>}
 */
export async function buildOutputWorkbook({ rows, picklist, fileStats }) {
  const ExcelJS = getExcelJS();
  const wb = new ExcelJS.Workbook();

  const ws1 = wb.addWorksheet(RECIPIENT_SHEET_NAME);
  ws1.addRow(RECIPIENT_SHEET_HEADERS);
  for (const row of rows) ws1.addRow(row);
  const phoneCol = RECIPIENT_SHEET_HEADERS.indexOf('收件人電話') + 1;
  const zipCol = RECIPIENT_SHEET_HEADERS.indexOf('郵遞區號') + 1;
  for (let r = 2; r <= ws1.rowCount; r++) {
    ws1.getCell(r, phoneCol).numFmt = '@';
    ws1.getCell(r, zipCol).numFmt = '@';
  }
  styleSheet(ws1);

  const ws2 = wb.addWorksheet(PICKLIST_SHEET_NAME);
  ws2.addRow(PICKLIST_SHEET_HEADERS);
  for (const [key, qty] of picklist) ws2.addRow([key, qty]);
  styleSheet(ws2);

  const ws3 = wb.addWorksheet(LOG_SHEET_NAME);
  ws3.addRow(LOG_SHEET_HEADERS);
  for (const stat of fileStats) ws3.addRow(stat);
  styleSheet(ws3);

  return wb.xlsx.writeBuffer();
}
