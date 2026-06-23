/**
 * Excel bank-statement template + upload parsing.
 * 1:1 port of Backend_Tally/excel_bridge.py.
 */
const ExcelJS = require('exceljs');

const TEMPLATE_HEADERS = ['DATE', 'DESCRIPTION', 'CHEQUE NO.', 'Debit', 'Credit', 'LEDGER'];
const TEMPLATE_SAMPLE_ROWS = [
  ['01/04/25', '50100215223054-TPT-PRA', '0000000250745865', 20000, '', 'Suspense A/c'],
  ['01/04/25', 'IMPS-509120183257-AKAS', '0000509120183257', 27000, '', 'Suspense A/c'],
];

const HEADER_ALIASES = {
  date: ['date'],
  description: ['description', 'narration'],
  chequeNo: ['cheque no.', 'cheque no', 'chequeno', 'chq no', 'chq./ref.no.', 'ref no', 'reference'],
  debit: ['debit'],
  credit: ['credit'],
  ledger: ['ledger'],
};

async function buildTemplateWorkbook() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('BankStatement');
  TEMPLATE_HEADERS.forEach((header, idx) => {
    const cell = ws.getCell(1, idx + 1);
    cell.value = header;
    cell.font = { bold: true };
  });
  TEMPLATE_SAMPLE_ROWS.forEach((rowData, rowIdx) => {
    rowData.forEach((value, colIdx) => {
      ws.getCell(rowIdx + 2, colIdx + 1).value = value === '' ? null : value;
    });
  });
  for (let col = 1; col <= TEMPLATE_HEADERS.length; col++) {
    ws.getColumn(col).width = 22;
  }
  return wb.xlsx.writeBuffer();
}

function normalizeHeader(text) {
  return String(text || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function matchColumns(headers) {
  const normalized = headers.map(normalizeHeader);
  const colMap = {};
  for (const [field, aliases] of Object.entries(HEADER_ALIASES)) {
    for (let idx = 0; idx < normalized.length; idx++) {
      if (aliases.includes(normalized[idx])) { colMap[field] = idx; break; }
    }
  }
  return colMap;
}

const DATE_FORMATS = [
  [/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, 'dmy'],
  [/^(\d{1,2})-(\d{1,2})-(\d{4})$/, 'dmy'],
  [/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/, 'dmy2'],
  [/^(\d{1,2})-(\d{1,2})-(\d{2})$/, 'dmy2'],
  [/^(\d{4})-(\d{1,2})-(\d{1,2})$/, 'ymd'],
];
const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

function pad2(n) { return String(n).padStart(2, '0'); }

function formatDate(value) {
  if (value == null) return '';
  if (value instanceof Date) {
    return `${pad2(value.getDate())}/${pad2(value.getMonth() + 1)}/${value.getFullYear()}`;
  }
  const text = String(value).trim();
  if (!text) return '';
  for (const [regex, kind] of DATE_FORMATS) {
    const m = text.match(regex);
    if (!m) continue;
    let d, mo, y;
    if (kind === 'dmy') { d = +m[1]; mo = +m[2]; y = +m[3]; }
    else if (kind === 'dmy2') { d = +m[1]; mo = +m[2]; y = 2000 + +m[3]; }
    else { y = +m[1]; mo = +m[2]; d = +m[3]; }
    const date = new Date(y, mo - 1, d);
    if (date.getFullYear() === y && date.getMonth() === mo - 1 && date.getDate() === d) {
      return `${pad2(d)}/${pad2(mo)}/${y}`;
    }
  }
  const m2 = text.match(/^(\d{1,2})-([A-Za-z]+)-(\d{2,4})$/);
  if (m2) {
    const d = +m2[1];
    const monIdx = MONTHS.findIndex((mo) => m2[2].toLowerCase().startsWith(mo));
    let y = +m2[3];
    if (y < 100) y += 2000;
    if (monIdx !== -1) {
      const date = new Date(y, monIdx, d);
      if (date.getDate() === d) return `${pad2(d)}/${pad2(monIdx + 1)}/${y}`;
    }
  }
  return text;
}

function parseAmount(value) {
  if (value == null || value === '') return 0.0;
  if (typeof value === 'number') return value;
  const text = String(value).trim();
  if (!text) return 0.0;
  const negative = text.startsWith('-') || (text.startsWith('(') && text.endsWith(')'));
  const cleaned = text.replace(/[^\d.]/g, '');
  if (!cleaned) return 0.0;
  const result = parseFloat(cleaned);
  if (!Number.isFinite(result)) return 0.0;
  return negative ? -result : result;
}

/** Reads an uploaded bank statement Excel file (first row = headers) and
 * returns normalized rows: date, description, chequeNo, debit, credit, ledger. */
async function readExcelRows(content) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(content);
  const ws = wb.worksheets[0];

  const firstRow = ws.getRow(1);
  const headers = [];
  for (let i = 1; i <= ws.columnCount; i++) headers.push(String(firstRow.getCell(i).value ?? '').trim());
  const colMap = matchColumns(headers);

  const cellValue = (values, field) => {
    const idx = colMap[field];
    return idx !== undefined && idx < values.length ? values[idx] : null;
  };

  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const values = [];
    for (let i = 1; i <= ws.columnCount; i++) values.push(row.getCell(i).value);
    if (values.every((v) => v == null || v === '')) continue;

    const dateValue = formatDate(cellValue(values, 'date'));
    if (!dateValue) continue;

    rows.push({
      date: dateValue,
      description: String(cellValue(values, 'description') ?? '').trim(),
      chequeNo: String(cellValue(values, 'chequeNo') ?? '').trim(),
      debit: parseAmount(cellValue(values, 'debit')),
      credit: parseAmount(cellValue(values, 'credit')),
      ledger: String(cellValue(values, 'ledger') ?? '').trim(),
    });
  }
  return rows;
}

module.exports = { TEMPLATE_HEADERS, TEMPLATE_SAMPLE_ROWS, buildTemplateWorkbook, readExcelRows, formatDate, parseAmount };
