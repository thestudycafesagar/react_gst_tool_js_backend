/**
 * GSTR-2B (GST purchase return) -> Tally purchase voucher engine.
 * 1:1 port of Backend_Tally/gstr2b_bridge.py.
 */
const ExcelJS = require('exceljs');
const { panFromGstin } = require('../utils/common');
const { Element, SubElement, setAttr, toCompactXml, toPrettyXml } = require('../utils/elementTree');

const ALLOWED_TAX_RATES = [0.0, 5.0, 12.0, 18.0, 28.0, 40.0];
const TAX_RATE_TOLERANCE = 0.30;

const HEADER_PATTERNS = {
  gstin: ['gstin of supplier'],
  trade_name: ['trade/legal name', 'trade name', 'legal name'],
  invoice_no: ['invoice number', 'note number'],
  invoice_type: ['invoice type', 'note type'],
  invoice_date: ['invoice date', 'note date'],
  invoice_value: ['invoice value', 'note value'],
  place_of_supply: ['place of supply'],
  reverse_charge: ['reverse charge', 'supply attract reverse charge'],
  rate: ['rate(%)', 'rate (%)'],
  taxable_value: ['taxable value'],
  igst: ['integrated tax'],
  cgst: ['central tax'],
  sgst: ['state/ut tax', 'state tax', 'ut tax'],
  cess: ['cess'],
  filing_period: ['gstr-1/iff/gstr-5 period', 'gstr-1/iff period', 'filing period',
    'gstr-1/1a/iff period', 'gstr-1/iff/gstr-1a period', 'gstr-1/iff/1a period'],
  itc_avail: ['itc availability', 'itc avail'],
};

const GSTIN_STATE_MAP = {
  '01': 'Jammu And Kashmir', '02': 'Himachal Pradesh', '03': 'Punjab', '04': 'Chandigarh',
  '05': 'Uttarakhand', '06': 'Haryana', '07': 'Delhi', '08': 'Rajasthan', '09': 'Uttar Pradesh',
  10: 'Bihar', 11: 'Sikkim', 12: 'Arunachal Pradesh', 13: 'Nagaland', 14: 'Manipur',
  15: 'Mizoram', 16: 'Tripura', 17: 'Meghalaya', 18: 'Assam', 19: 'West Bengal',
  20: 'Jharkhand', 21: 'Odisha', 22: 'Chhattisgarh', 23: 'Madhya Pradesh',
  24: 'Gujarat', 25: 'Daman And Diu', 26: 'Dadra And Nagar Haveli And Daman And Diu',
  27: 'Maharashtra', 29: 'Karnataka', 30: 'Goa', 31: 'Lakshadweep', 32: 'Kerala',
  33: 'Tamil Nadu', 34: 'Puducherry', 35: 'Andaman And Nicobar Islands', 36: 'Telangana',
  37: 'Andhra Pradesh', 38: 'Ladakh', 97: 'Other Territory', 99: 'Centre Jurisdiction',
};

function stateFromGstin(gstin) {
  const code = String(gstin || '').trim().toUpperCase().slice(0, 2);
  return GSTIN_STATE_MAP[code] || '';
}

function normalizeStateName(value) {
  let raw = String(value || '').trim();
  if (!raw) return '';
  raw = raw.replace(/_/g, ' ').replace(/-/g, ' ').trim();
  if (/^\d{2}/.test(raw)) raw = raw.slice(2).replace(/^[ -]+/, '');
  const parts = raw.split(/\s+/).filter(Boolean);
  if (!parts.length) return '';
  const smallWords = new Set(['and', 'of', 'ut']);
  return parts.map((w) => (smallWords.has(w.toLowerCase()) ? w.toUpperCase() : w[0].toUpperCase() + w.slice(1).toLowerCase())).join(' ');
}

function pad2(n) { return String(n).padStart(2, '0'); }

function excelSerialToDate(serial) {
  const base = new Date(1899, 11, 30);
  base.setDate(base.getDate() + Math.floor(serial));
  return base;
}

const MONTH_ABBR = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const MONTH_FULL = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];

function monthIndexFromName(name) {
  const lower = name.toLowerCase();
  let idx = MONTH_FULL.findIndex((m) => m === lower);
  if (idx !== -1) return idx;
  idx = MONTH_ABBR.findIndex((m) => m === lower);
  return idx;
}

/** Normalizes any date representation to DD/MM/YYYY — mirrors _normalize_date_str. */
function normalizeDateStr(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date) {
    return `${pad2(value.getDate())}/${pad2(value.getMonth() + 1)}/${value.getFullYear()}`;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    try {
      const d = excelSerialToDate(value);
      return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
    } catch {
      // fall through to text parsing
    }
  }
  const text = String(value).trim();
  if (!text) return '';
  if (/^\d{8}$/.test(text)) {
    const head = text.slice(0, 4);
    if (/^\d{4}$/.test(head) && parseInt(head, 10) >= 1900 && parseInt(head, 10) <= 2100) {
      const y = +text.slice(0, 4), m = +text.slice(4, 6), d = +text.slice(6, 8);
      const dt = new Date(y, m - 1, d);
      if (dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d) {
        return `${pad2(d)}/${pad2(m)}/${y}`;
      }
    } else {
      const d = +text.slice(0, 2), m = +text.slice(2, 4), y = +text.slice(4, 8);
      const dt = new Date(y, m - 1, d);
      if (dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d) {
        return `${pad2(d)}/${pad2(m)}/${y}`;
      }
    }
  }
  let m;
  if ((m = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/))) return `${pad2(+m[1])}/${pad2(+m[2])}/${m[3]}`;
  if ((m = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2})$/))) return `${pad2(+m[1])}/${pad2(+m[2])}/${2000 + +m[3]}`;
  if ((m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) return `${pad2(+m[3])}/${pad2(+m[2])}/${m[1]}`;
  if ((m = text.match(/^(\d{1,2})[/-]([A-Za-z]+)[/-](\d{4})$/))) {
    const mi = monthIndexFromName(m[2]);
    if (mi !== -1) return `${pad2(+m[1])}/${pad2(mi + 1)}/${m[3]}`;
  }
  if ((m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})T/))) return `${pad2(+m[3])}/${pad2(+m[2])}/${m[1]}`;
  return text;
}

/** Converts a normalized DD/MM/YYYY (or raw) date to Tally's YYYYMMDD —
 * mirrors _tally_date. fallbackToday matches the Python kwarg. */
function tallyDate(value, fallbackToday = true) {
  const dateText = normalizeDateStr(value);
  const today = () => {
    const d = new Date();
    return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`;
  };
  if (!dateText) return fallbackToday ? today() : '';
  const parts = dateText.split('/');
  if (parts.length === 3) {
    const [dd, mm, yy] = parts;
    const candidate = `${yy}${mm.padStart(2, '0')}${dd.padStart(2, '0')}`;
    if (/^\d{8}$/.test(candidate)) return candidate;
  }
  const compact = dateText.replace(/\//g, '').replace(/-/g, '');
  if (/^\d{8}$/.test(compact)) return compact;
  return fallbackToday ? today() : '';
}

// ─── Column detection (B2BColumnMapper equivalent) — operates on exceljs worksheets ──

function cellText(cell) {
  const v = cell ? cell.value : null;
  if (v == null) return '';
  if (typeof v === 'object' && 'richText' in v) return v.richText.map((r) => r.text).join('');
  if (typeof v === 'object' && 'result' in v) return String(v.result);
  return String(v);
}

function findHeaderRows(ws) {
  const maxRow = Math.min(15, ws.rowCount);
  for (let rowIdx = 1; rowIdx <= maxRow; rowIdx++) {
    const row = ws.getRow(rowIdx);
    for (let c = 1; c <= ws.columnCount; c++) {
      const text = cellText(row.getCell(c)).trim().toLowerCase();
      if (text.includes('gstin of supplier')) return [rowIdx, rowIdx + 1];
    }
  }
  return [5, 6];
}

/** Returns {columnMap, dataStartRow, headerRow1, headerRow2}. */
function detectColumns(ws) {
  const [headerRow1, headerRow2] = findHeaderRows(ws);
  const headersR1 = {}, headersR2 = {};
  const row1 = ws.getRow(headerRow1), row2 = ws.getRow(headerRow2);
  for (let c = 1; c <= ws.columnCount; c++) {
    const t1 = cellText(row1.getCell(c)).trim().toLowerCase();
    if (t1) headersR1[c - 1] = t1;
    const t2 = cellText(row2.getCell(c)).trim().toLowerCase();
    if (t2) headersR2[c - 1] = t2;
  }

  const allHeaders = {};
  for (let colIdx = 0; colIdx < ws.columnCount; colIdx++) {
    const parts = [headersR1[colIdx], headersR2[colIdx]].filter(Boolean);
    if (parts.length) allHeaders[colIdx] = parts.join(' | ');
  }

  const columnMap = {};
  for (const [field, patterns] of Object.entries(HEADER_PATTERNS)) {
    for (const [colIdxStr, headerText] of Object.entries(allHeaders)) {
      if (patterns.some((p) => headerText.includes(p))) {
        const colIdx = +colIdxStr;
        if (!(field in columnMap)) columnMap[field] = colIdx;
        break;
      }
    }
  }

  const invNoCols = Object.entries(allHeaders)
    .filter(([, h]) => h.includes('invoice number') || h.includes('note number'))
    .map(([c]) => +c).sort((a, b) => a - b);
  const isB2ba = Object.values(allHeaders).some((h) => h.includes('original invoice') || h.includes('revised invoice'))
    || invNoCols.length >= 2;
  if (isB2ba) {
    const invDateCols = Object.entries(allHeaders)
      .filter(([, h]) => h.includes('invoice date') || h.includes('note date'))
      .map(([c]) => +c).sort((a, b) => a - b);
    if (invNoCols.length >= 2) {
      columnMap.orig_invoice_no = invNoCols[0];
      columnMap.invoice_no = invNoCols[1];
    } else if (invNoCols.length) {
      columnMap.orig_invoice_no = invNoCols[0];
    }
    if (invDateCols.length >= 2) {
      columnMap.orig_invoice_date = invDateCols[0];
      columnMap.invoice_date = invDateCols[1];
    } else if (invDateCols.length) {
      columnMap.orig_invoice_date = invDateCols[0];
    }
  }

  const gstinCol = columnMap.gstin !== undefined ? columnMap.gstin : 0;
  let dataStartRow = headerRow2 + 1;
  for (let rowIdx = headerRow2; rowIdx <= Math.min(headerRow2 + 10, ws.rowCount); rowIdx++) {
    const val = ws.getRow(rowIdx).getCell(gstinCol + 1).value;
    if (val && typeof val === 'string' && val.trim().length >= 15) {
      dataStartRow = rowIdx;
      break;
    }
  }

  return { columnMap, dataStartRow, headerRow1, headerRow2 };
}

function safeGet(row, field, columnMap, defaultValue = null) {
  const colIdx = columnMap[field];
  if (colIdx === undefined || colIdx >= row.length) return defaultValue;
  const v = row[colIdx];
  return v == null ? defaultValue : v;
}

function safeFloat(row, field, columnMap, defaultValue = 0.0) {
  const val = safeGet(row, field, columnMap);
  if (val == null) return defaultValue;
  const n = parseFloat(val);
  return Number.isFinite(n) ? n : defaultValue;
}

function safeStr(row, field, columnMap, defaultValue = '') {
  const val = safeGet(row, field, columnMap);
  return val == null ? defaultValue : String(val).trim();
}

/** Parses one B2B/B2BA/CDNR data row into a flat record dict, or null if no GSTIN. */
function parseB2bRow(row, rowIdx, columnMap) {
  const gstin = safeStr(row, 'gstin', columnMap);
  if (!gstin || gstin.length < 15) return null;

  const invDateRaw = safeGet(row, 'invoice_date', columnMap);
  let invDate;
  if (invDateRaw instanceof Date) {
    invDate = `${pad2(invDateRaw.getDate())}/${pad2(invDateRaw.getMonth() + 1)}/${invDateRaw.getFullYear()}`;
  } else if (invDateRaw) {
    invDate = String(invDateRaw);
  } else {
    invDate = '';
  }

  let taxable = safeFloat(row, 'taxable_value', columnMap);
  let igst = safeFloat(row, 'igst', columnMap);
  let cgst = safeFloat(row, 'cgst', columnMap);
  let sgst = safeFloat(row, 'sgst', columnMap);
  let cess = safeFloat(row, 'cess', columnMap);
  let invoiceValue = safeFloat(row, 'invoice_value', columnMap);

  const invoiceType = safeStr(row, 'invoice_type', columnMap, 'Regular');
  if (invoiceType.toLowerCase().includes('credit note')) {
    taxable = taxable ? -Math.abs(taxable) : 0.0;
    igst = igst ? -Math.abs(igst) : 0.0;
    cgst = cgst ? -Math.abs(cgst) : 0.0;
    sgst = sgst ? -Math.abs(sgst) : 0.0;
    cess = cess ? -Math.abs(cess) : 0.0;
    invoiceValue = invoiceValue ? -Math.abs(invoiceValue) : 0.0;
  }

  let rate;
  if ('rate' in columnMap) {
    rate = safeFloat(row, 'rate', columnMap);
  } else {
    const totalTax = Math.abs(igst) + Math.abs(cgst) + Math.abs(sgst);
    rate = (Math.abs(taxable) > 0 && totalTax > 0) ? Math.round((totalTax / Math.abs(taxable)) * 100) : 0;
  }

  return {
    gstin,
    trade_name: safeStr(row, 'trade_name', columnMap),
    invoice_no: safeStr(row, 'invoice_no', columnMap),
    orig_invoice_no: 'orig_invoice_no' in columnMap ? safeStr(row, 'orig_invoice_no', columnMap) : '',
    invoice_type: invoiceType,
    invoice_date: invDate,
    invoice_value: invoiceValue,
    place_of_supply: safeStr(row, 'place_of_supply', columnMap),
    reverse_charge: safeStr(row, 'reverse_charge', columnMap, 'No'),
    rate,
    taxable_value: taxable,
    igst, cgst, sgst, cess,
    filing_period: safeStr(row, 'filing_period', columnMap),
    itc_avail: safeStr(row, 'itc_avail', columnMap),
    row_idx: rowIdx,
  };
}

function parseReadme(ws) {
  const info = { company_gstin: '', company_name: '', trade_name: '', financial_year: '', tax_period: '' };
  const maxRow = Math.min(15, ws.rowCount);
  for (let rowIdx = 1; rowIdx <= maxRow; rowIdx++) {
    const row = ws.getRow(rowIdx);
    const vals = [];
    for (let c = 1; c <= Math.max(3, ws.columnCount); c++) vals.push(row.getCell(c).value);
    if (!vals.length || !vals[0]) continue;
    const label = String(vals[0]).trim();
    const value = vals.length > 2 ? String(vals[2] || '') : '';
    if (label === 'GSTIN') info.company_gstin = value;
    else if (label === 'Legal Name') info.company_name = value;
    else if (label.includes('Trade Name')) info.trade_name = value;
    else if (label === 'Financial Year') info.financial_year = value;
    else if (label === 'Tax Period') info.tax_period = value;
  }
  return info;
}

/** Parses a GSTR-2B Excel export (B2B / B2BA / B2B-CDNR sheets).
 * Returns {success, records, errors, warnings, company_gstin, company_name,
 * trade_name, financial_year, tax_period}. */
async function parseGstr2bExcel(content) {
  const errors = [], warnings = [];
  let wb;
  try {
    wb = new ExcelJS.Workbook();
    await wb.xlsx.load(content);
  } catch (exc) {
    return { success: false, records: [], errors: [`Failed to open file: ${exc.message}`], warnings: [] };
  }

  let readmeInfo = { company_gstin: '', company_name: '', trade_name: '', financial_year: '', tax_period: '' };
  const sheetByUpper = {};
  for (const ws of wb.worksheets) sheetByUpper[ws.name.toUpperCase()] = ws;
  if (sheetByUpper['READ ME']) readmeInfo = parseReadme(sheetByUpper['READ ME']);
  else {
    const readMeSheet = wb.worksheets.find((s) => s.name === 'Read me');
    if (readMeSheet) readmeInfo = parseReadme(readMeSheet);
  }

  const b2bSheets = sheetByUpper.B2B ? [sheetByUpper.B2B] : [];
  const cdnrSheets = sheetByUpper['B2B-CDNR'] ? [sheetByUpper['B2B-CDNR']] : [];
  const b2baSheets = sheetByUpper.B2BA ? [sheetByUpper.B2BA] : [];

  if (!b2bSheets.length && !cdnrSheets.length && !b2baSheets.length) {
    return {
      success: false, records: [],
      errors: ['No B2B, B2B-CDNR, or B2BA sheets found in the uploaded file!'],
      warnings: [], ...readmeInfo,
    };
  }

  const records = [];

  const b2baRecords = [];
  const b2baOrigNos = new Set();
  for (const ws of b2baSheets) {
    const detected = detectColumns(ws);
    const columnMap = detected.columnMap;
    if (!('gstin' in columnMap)) {
      warnings.push(`Could not detect columns in ${ws.name} sheet — skipped.`);
      continue;
    }
    for (let rowIdx = detected.dataStartRow; rowIdx <= ws.rowCount; rowIdx++) {
      const wsRow = ws.getRow(rowIdx);
      const row = [];
      for (let c = 1; c <= ws.columnCount; c++) row.push(wsRow.getCell(c).value);
      const gstinCol = columnMap.gstin !== undefined ? columnMap.gstin : 0;
      if (gstinCol >= row.length || !row[gstinCol]) continue;
      try {
        const record = parseB2bRow(row, rowIdx, columnMap);
        if (record) {
          record.is_amendment = true;
          record.sheet_type = 'B2BA';
          b2baRecords.push(record);
          const origNo = (record.orig_invoice_no || '').trim().toUpperCase();
          const invNo = (record.invoice_no || '').trim().toUpperCase();
          b2baOrigNos.add(origNo || invNo);
        }
      } catch (exc) {
        errors.push(`${ws.name} Row ${rowIdx}: ${exc.message}`);
      }
    }
  }

  if (b2baSheets.length && !b2baRecords.length) {
    warnings.push('B2BA sheet found but no valid amendment records could be parsed.');
  } else if (b2baRecords.length) {
    warnings.push(`B2BA: ${b2baRecords.length} amendment record(s) found. ${b2baOrigNos.size} original invoice(s) will be replaced.`);
  }

  for (const ws of [...b2bSheets, ...cdnrSheets]) {
    const isCdnr = cdnrSheets.includes(ws);
    const detected = detectColumns(ws);
    const columnMap = detected.columnMap;
    if (!('gstin' in columnMap)) {
      warnings.push(`Could not detect 'GSTIN of supplier' column in ${ws.name} sheet!`);
      continue;
    }
    if (!('rate' in columnMap)) {
      warnings.push(`${ws.name}: Optional column not found: rate — will auto-calculate`);
    }

    let skippedAmendments = 0;
    let lastGstin = '', lastTradeName = '';
    const tradeCol = columnMap.trade_name !== undefined ? columnMap.trade_name : 1;
    const gstinCol = columnMap.gstin !== undefined ? columnMap.gstin : 0;

    for (let rowIdx = detected.dataStartRow; rowIdx <= ws.rowCount; rowIdx++) {
      const wsRow = ws.getRow(rowIdx);
      let row = [];
      for (let c = 1; c <= ws.columnCount; c++) row.push(wsRow.getCell(c).value);

      const curGstin = gstinCol < row.length ? String(row[gstinCol] || '').trim() : '';
      if (curGstin.length >= 15) {
        lastGstin = curGstin;
        lastTradeName = tradeCol < row.length ? String(row[tradeCol] || '').trim() : '';
      } else if (lastGstin) {
        row = row.slice();
        while (row.length <= Math.max(gstinCol, tradeCol)) row.push(null);
        row[gstinCol] = lastGstin;
        if (tradeCol < row.length && !String(row[tradeCol] || '').trim()) row[tradeCol] = lastTradeName;
      } else {
        continue;
      }

      try {
        const record = parseB2bRow(row, rowIdx, columnMap);
        if (record) {
          record.sheet_type = isCdnr ? 'CDNR' : 'B2B';
          const invKey = (record.invoice_no || '').trim().toUpperCase();
          if (b2baOrigNos.has(invKey)) skippedAmendments += 1;
          else records.push(record);
        }
      } catch (exc) {
        errors.push(`${ws.name} Row ${rowIdx}: ${exc.message}`);
      }
    }

    if (skippedAmendments) {
      warnings.push(`${ws.name}: ${skippedAmendments} invoice(s) skipped — superseded by B2BA amendment(s).`);
    }
  }

  records.push(...b2baRecords);

  return {
    success: records.length > 0,
    records, errors, warnings, ...readmeInfo,
  };
}

function nearestAllowedTaxRate(rateValue) {
  return ALLOWED_TAX_RATES.reduce((best, r) => (Math.abs(r - rateValue) < Math.abs(best - rateValue) ? r : best));
}

/** Validates GST structure (IGST vs CGST/SGST exclusivity, CGST==SGST,
 * rate within allowed slabs). Returns [validRecords, invalidIssues]. */
function validateTaxConfiguration(records) {
  const validRecords = [], invalidIssues = [];
  for (const rec of records) {
    const taxable = Math.abs(parseFloat(rec.taxable_value || 0.0));
    const igstAmt = Math.abs(parseFloat(rec.igst || 0.0));
    const cgstAmt = Math.abs(parseFloat(rec.cgst || 0.0));
    const sgstAmt = Math.abs(parseFloat(rec.sgst || 0.0));
    const cessAmt = Math.abs(parseFloat(rec.cess || 0.0));

    const hasIgst = igstAmt > 0.009;
    const hasCgst = cgstAmt > 0.009;
    const hasSgst = sgstAmt > 0.009;
    const reasons = [];

    if (hasIgst && (hasCgst || hasSgst)) reasons.push('IGST cannot be present together with CGST/SGST.');
    if (hasCgst !== hasSgst) reasons.push('CGST and SGST must both be present (or both zero).');
    if (hasCgst && hasSgst && Math.abs(cgstAmt - sgstAmt) > 1.0) reasons.push('CGST and SGST amounts are not equal.');
    if (taxable <= 0 && (hasIgst || hasCgst || hasSgst)) reasons.push('Tax amount exists but taxable value is zero.');

    const taxStructure = hasIgst ? 'IGST' : ((hasCgst || hasSgst) ? 'CGST+SGST' : 'No GST');
    let computedRate = 0.0;
    if (taxable > 0) computedRate = hasIgst ? (igstAmt / taxable * 100.0) : ((cgstAmt + sgstAmt) / taxable * 100.0);

    const nearestRate = nearestAllowedTaxRate(computedRate);
    if (!rec.is_multi_rate && Math.abs(computedRate - nearestRate) > TAX_RATE_TOLERANCE) {
      reasons.push(`Computed GST rate ${computedRate.toFixed(2)}% is not in allowed slabs (0, 5, 12, 18, 28, 40).`);
    }

    if (reasons.length) {
      invalidIssues.push({
        row_idx: rec.row_idx || '', invoice_no: rec.invoice_no || '', party_name: rec.trade_name || '',
        taxable_value: taxable, igst: igstAmt, cgst: cgstAmt, sgst: sgstAmt, cess: cessAmt,
        tax_structure: taxStructure, computed_rate: Math.round(computedRate * 10000) / 10000,
        sheet_rate: rec.rate || '', nearest_allowed_rate: nearestRate, issue: reasons.join(' | '),
      });
    } else {
      validRecords.push(rec);
    }
  }
  return [validRecords, invalidIssues];
}

/** Two-pass lookup: direct match against total GST rate, then half-rate
 * match (when the user configured the individual CGST/SGST rate). */
function getGstLedger(gstLedgerRateMap, taxType, rate) {
  if (!gstLedgerRateMap || !Object.keys(gstLedgerRateMap).length) return '';
  let r;
  try { r = Math.round(parseFloat(rate || 0) * 100) / 100; } catch { r = 0.0; }
  if (!Number.isFinite(r)) r = 0.0;
  const halfR = r / 2.0;
  for (const [mappedRate, mapping] of Object.entries(gstLedgerRateMap)) {
    if (mappedRate === 'default') continue;
    const mr = parseFloat(mappedRate);
    if (Number.isFinite(mr) && Math.abs(mr - r) < 0.15) {
      const ledger = (mapping || {})[taxType] || '';
      if (ledger) return ledger;
    }
  }
  for (const [mappedRate, mapping] of Object.entries(gstLedgerRateMap)) {
    if (mappedRate === 'default') continue;
    const mr = parseFloat(mappedRate);
    if (Number.isFinite(mr) && Math.abs(mr - halfR) < 0.15) {
      const ledger = (mapping || {})[taxType] || '';
      if (ledger) return ledger;
    }
  }
  return (gstLedgerRateMap.default || {})[taxType] || '';
}

function getPurchaseLedger(partyLedgerMap, partyName, defaultValue) {
  if (!partyLedgerMap || !Object.keys(partyLedgerMap).length) return defaultValue;
  const key = (partyName || '').toUpperCase().trim();
  return key in partyLedgerMap ? partyLedgerMap[key] : defaultValue;
}

function resolvePurchaseLedger(rec, partyLedgerMap, defaultPurchaseLedger) {
  let ledger = String(rec.purchase_ledger || '').trim();
  if (!ledger) {
    const partyName = rec.trade_name || rec.party_name || '';
    ledger = getPurchaseLedger(partyLedgerMap, partyName, defaultPurchaseLedger);
  }
  ledger = (ledger || defaultPurchaseLedger).trim();

  let rate;
  try { rate = parseFloat(rec.rate || 0); } catch { rate = 0.0; }
  if (!Number.isFinite(rate)) rate = 0.0;
  if (rate > 0 && /\d+(?:\.\d+)?\s*%/.test(ledger)) {
    const rateClean = Math.abs(rate - Math.trunc(rate)) < 0.01 ? Math.trunc(rate) : rate;
    ledger = ledger.replace(/\d+(?:\.\d+)?\s*%/, `${rateClean}%`);
  }
  return ledger;
}

function asFloat(val, defaultValue = 0.0) {
  if (val === null || val === undefined || val === '') return defaultValue;
  const n = parseFloat(val);
  return Number.isFinite(n) ? n : defaultValue;
}

function coalesceText(current, incoming) {
  return String(current || '').trim() ? current : incoming;
}

function isIneligible(value) {
  return ['NO', 'N', 'INELIGIBLE'].includes(String(value || '').trim().toUpperCase());
}

function normalizeInvoiceNo(value) {
  let text = String(value || '').trim();
  if (text.endsWith('.0') && /^\d+$/.test(text.slice(0, -2))) text = text.slice(0, -2);
  return text;
}

/** Combines rows that belong to the same invoice (same gstin/party/invoice
 * no/date/type) into a single voucher record, accumulating per-ledger GST
 * splits so multi-rate invoices still produce one voucher with multiple
 * tax-ledger entries. */
function consolidateInvoiceRecords(records, defaultPurchaseLedger, partyLedgerMap = null, gstLedgerRateMap = null) {
  partyLedgerMap = partyLedgerMap || {};
  gstLedgerRateMap = gstLedgerRateMap || {};
  if (!records || !records.length) return [];

  const grouped = [];
  const keyMap = new Map();

  for (const rec of records) {
    const invoiceNo = normalizeInvoiceNo(rec.invoice_no || rec.supplier_invoice_no || '');
    if (!invoiceNo) {
      grouped.push({ ...rec });
      continue;
    }

    const invoiceTypeKey = String(rec.invoice_type || 'Regular').trim().toLowerCase();
    const sheetTypeKey = String(rec.sheet_type || '').trim().toUpperCase();
    const origInvoiceKey = normalizeInvoiceNo(rec.orig_invoice_no || '');
    const dateKey = normalizeDateStr(rec.invoice_date || rec.supplier_invoice_date || '');

    const key = [
      String(rec.gstin || '').trim().toUpperCase(),
      String(rec.trade_name || rec.party_name || '').trim().toUpperCase(),
      invoiceNo.toUpperCase(), dateKey, invoiceTypeKey, sheetTypeKey, origInvoiceKey,
    ].join('');

    let base = keyMap.get(key);
    if (base === undefined) {
      base = { ...rec };
      base.invoice_no = invoiceNo;
      if (!String(base.supplier_invoice_no || '').trim()) base.supplier_invoice_no = invoiceNo;
      if (origInvoiceKey && !String(base.orig_invoice_no || '').trim()) base.orig_invoice_no = origInvoiceKey;
      const baseLedger = resolvePurchaseLedger(rec, partyLedgerMap, defaultPurchaseLedger);
      base.purchase_ledger = baseLedger;
      base.purchase_ledger_splits = { [baseLedger]: asFloat(rec.taxable_value) };
      for (const f of ['taxable_value', 'igst', 'cgst', 'sgst', 'cess', 'invoice_value']) base[f] = asFloat(base[f]);

      const recRate = rec.rate;
      if (asFloat(rec.cgst)) {
        const cl = (rec.cgst_ledger || '').trim() || getGstLedger(gstLedgerRateMap, 'cgst', recRate) || 'CGST';
        base.cgst_ledger_splits = { [cl]: asFloat(rec.cgst) };
      }
      if (asFloat(rec.sgst)) {
        const sl = (rec.sgst_ledger || '').trim() || getGstLedger(gstLedgerRateMap, 'sgst', recRate) || 'SGST';
        base.sgst_ledger_splits = { [sl]: asFloat(rec.sgst) };
      }
      if (asFloat(rec.igst)) {
        const il = (rec.igst_ledger || '').trim() || getGstLedger(gstLedgerRateMap, 'igst', recRate) || 'IGST';
        base.igst_ledger_splits = { [il]: asFloat(rec.igst) };
      }
      keyMap.set(key, base);
      grouped.push(base);
      continue;
    }

    if (rec.rate !== base.rate) base.is_multi_rate = true;
    base.taxable_value = asFloat(base.taxable_value) + asFloat(rec.taxable_value);
    base.igst = asFloat(base.igst) + asFloat(rec.igst);
    base.cgst = asFloat(base.cgst) + asFloat(rec.cgst);
    base.sgst = asFloat(base.sgst) + asFloat(rec.sgst);
    base.cess = asFloat(base.cess) + asFloat(rec.cess);

    const mrecRate = rec.rate;
    if (asFloat(rec.cgst)) {
      const ck = (rec.cgst_ledger || '').trim() || getGstLedger(gstLedgerRateMap, 'cgst', mrecRate) || 'CGST';
      if (!base.cgst_ledger_splits) base.cgst_ledger_splits = {};
      base.cgst_ledger_splits[ck] = asFloat(base.cgst_ledger_splits[ck], 0.0) + asFloat(rec.cgst);
    }
    if (asFloat(rec.sgst)) {
      const sk = (rec.sgst_ledger || '').trim() || getGstLedger(gstLedgerRateMap, 'sgst', mrecRate) || 'SGST';
      if (!base.sgst_ledger_splits) base.sgst_ledger_splits = {};
      base.sgst_ledger_splits[sk] = asFloat(base.sgst_ledger_splits[sk], 0.0) + asFloat(rec.sgst);
    }
    if (asFloat(rec.igst)) {
      const ik = (rec.igst_ledger || '').trim() || getGstLedger(gstLedgerRateMap, 'igst', mrecRate) || 'IGST';
      if (!base.igst_ledger_splits) base.igst_ledger_splits = {};
      base.igst_ledger_splits[ik] = asFloat(base.igst_ledger_splits[ik], 0.0) + asFloat(rec.igst);
    }

    const splitLedger = resolvePurchaseLedger(rec, partyLedgerMap, defaultPurchaseLedger);
    if (!base.purchase_ledger_splits) base.purchase_ledger_splits = {};
    base.purchase_ledger_splits[splitLedger] = asFloat(base.purchase_ledger_splits[splitLedger], 0.0) + asFloat(rec.taxable_value);

    for (const f of ['voucher_date', 'voucher_no', 'invoice_date', 'invoice_type', 'party_name',
      'party_mailing_name', 'party_address1', 'party_address2', 'party_pincode',
      'party_state', 'place_of_supply', 'purchase_ledger', 'narration',
      'tds_ledger', 'tds_rate']) {
      base[f] = coalesceText(base[f], rec[f]);
    }

    const supplierInvNo = normalizeInvoiceNo(rec.supplier_invoice_no || '');
    if (supplierInvNo) base.supplier_invoice_no = coalesceText(base.supplier_invoice_no, supplierInvNo);
    if (rec.supplier_invoice_date) base.supplier_invoice_date = coalesceText(base.supplier_invoice_date, rec.supplier_invoice_date);

    if (['YES', 'Y'].includes(String(rec.reverse_charge || '').trim().toUpperCase())) base.reverse_charge = 'Yes';

    if (isIneligible(rec.itc_avail) || isIneligible(base.itc_avail)) {
      base.itc_avail = 'Ineligible';
    } else {
      base.itc_avail = coalesceText(base.itc_avail, rec.itc_avail) || 'Yes';
    }

    const baseTdsAmount = asFloat(base.tds_amount, null);
    const recTdsAmount = asFloat(rec.tds_amount, null);
    if (baseTdsAmount === null && recTdsAmount === null) {
      base.tds_amount = '';
    } else {
      base.tds_amount = asFloat(baseTdsAmount, 0.0) + asFloat(recTdsAmount, 0.0);
    }
  }

  for (const rec of grouped) {
    const taxable = asFloat(rec.taxable_value);
    const igst = asFloat(rec.igst), cgst = asFloat(rec.cgst), sgst = asFloat(rec.sgst), cess = asFloat(rec.cess);
    rec.invoice_value = taxable + igst + cgst + sgst + cess;
    const taxTotal = Math.abs(igst) + Math.abs(cgst) + Math.abs(sgst);
    rec.rate = (taxable && taxTotal) ? Math.round((taxTotal / Math.abs(taxable)) * 100 * 100) / 100 : 0;
  }
  return grouped;
}

function fmt2(n) { return Number(n).toFixed(2); }

function addCommonLedgerFlags(node, isParty = 'No') {
  SubElement(node, 'GSTCLASS').text = 'Not Applicable';
  SubElement(node, 'ISDEEMEDPOSITIVE').text = isParty === 'Yes' ? 'No' : 'Yes';
  SubElement(node, 'LEDGERFROMITEM').text = 'No';
  SubElement(node, 'REMOVEZEROENTRIES').text = 'No';
  SubElement(node, 'ISPARTYLEDGER').text = isParty;
  SubElement(node, 'GSTOVERRIDDEN').text = 'No';
  SubElement(node, 'ISGSTASSESSABLEVALUEOVERRIDDEN').text = 'No';
}

function findChild(node, tag) {
  return node.children.find((c) => c.tag === tag);
}

/** Adds GSTREGISTRATION/CMPGSTIN/CMPGSTSTATE; returns [companyGstin, companyState]. */
function companyRegistrationBlock(ctx, voucher, partyGstin) {
  const companyGstin = String(ctx.company_gstin || '').trim().toUpperCase();
  if (!companyGstin) return ['', ''];
  const companyState = ctx.company_registration_state || stateFromGstin(companyGstin);
  const regName = ctx.company_registration_name || (companyState ? `${companyState} Registration` : companyGstin);
  SubElement(voucher, 'GSTTRANSACTIONTYPE').text = partyGstin ? 'Tax Invoice' : 'Unregistered';
  SubElement(voucher, 'GSTREGISTRATIONTYPE').text = 'Regular';
  if (partyGstin) SubElement(voucher, 'PARTYGSTIN').text = partyGstin;
  const gstReg = SubElement(voucher, 'GSTREGISTRATION');
  setAttr(gstReg, 'TAXTYPE', 'GST');
  setAttr(gstReg, 'TAXREGISTRATION', companyGstin);
  gstReg.text = regName;
  SubElement(voucher, 'CMPGSTIN').text = companyGstin;
  SubElement(voucher, 'CMPGSTREGISTRATIONTYPE').text = 'Regular';
  if (companyState) SubElement(voucher, 'CMPGSTSTATE').text = companyState;
  return [companyGstin, companyState];
}

function buildPurchaseVoucherXml(parent, rec, purchaseLedger, narration, voucherDate, roundOffLedger, ctx) {
  const tallyMsg = SubElement(parent, 'TALLYMESSAGE');
  setAttr(tallyMsg, 'xmlns:UDF', 'TallyUDF');
  const voucher = SubElement(tallyMsg, 'VOUCHER');
  setAttr(voucher, 'REMOTEID', '');

  const invTypeLower = String(rec.invoice_type || '').toLowerCase();
  let vchType;
  if (invTypeLower.includes('credit note')) vchType = 'Debit Note';
  else if (invTypeLower.includes('debit note')) vchType = 'Credit Note';
  else vchType = 'Purchase';
  const isDebitNote = vchType === 'Debit Note';
  setAttr(voucher, 'VCHTYPE', vchType);
  setAttr(voucher, 'ACTION', 'Create');
  setAttr(voucher, 'OBJVIEW', 'Invoice Voucher View');

  const actualVoucherDate = rec.voucher_date || rec.invoice_date || voucherDate;
  const tallyDt = tallyDate(actualVoucherDate, true);
  const refDate = tallyDate(rec.supplier_invoice_date || rec.invoice_date || actualVoucherDate, false);

  const partyName = rec.party_name || rec.trade_name;
  const partyLedger = rec.trade_name;
  const partyGstin = String(rec.gstin || '').trim().toUpperCase();
  const partyState = normalizeStateName(rec.party_state || stateFromGstin(partyGstin));
  const partyMailingName = rec.party_mailing_name || partyName;
  const partyAddress1 = rec.party_address1 || '';
  const partyAddress2 = rec.party_address2 || '';
  const partyPincode = String(rec.party_pincode || '').trim();
  const supplierInvoiceNo = rec.supplier_invoice_no || rec.invoice_no;

  if (partyAddress1 || partyAddress2) {
    const addressList = SubElement(voucher, 'ADDRESS.LIST');
    setAttr(addressList, 'TYPE', 'String');
    if (partyAddress1) SubElement(addressList, 'ADDRESS').text = partyAddress1;
    if (partyAddress2) SubElement(addressList, 'ADDRESS').text = partyAddress2;
  }

  SubElement(voucher, 'DATE').text = tallyDt;
  if (refDate) SubElement(voucher, 'REFERENCEDATE').text = refDate;
  SubElement(voucher, 'GSTREGISTRATIONTYPE').text = 'Regular';
  SubElement(voucher, 'VATDEALERTYPE').text = 'Regular';
  if (partyState) SubElement(voucher, 'STATENAME').text = partyState;
  SubElement(voucher, 'COUNTRYOFRESIDENCE').text = 'India';
  if (partyGstin) SubElement(voucher, 'PARTYGSTIN').text = partyGstin;
  const companyStateHint = ctx.company_registration_state || stateFromGstin(ctx.company_gstin || '');
  const placeOfSupply = companyStateHint || normalizeStateName(rec.place_of_supply || '');
  if (placeOfSupply) SubElement(voucher, 'PLACEOFSUPPLY').text = placeOfSupply;
  SubElement(voucher, 'VOUCHERTYPENAME').text = vchType;
  SubElement(voucher, 'PARTYNAME').text = partyName;
  companyRegistrationBlock(ctx, voucher, partyGstin);
  SubElement(voucher, 'PARTYLEDGERNAME').text = partyLedger;
  SubElement(voucher, 'REFERENCE').text = supplierInvoiceNo;
  SubElement(voucher, 'PARTYMAILINGNAME').text = partyMailingName;
  if (partyPincode) SubElement(voucher, 'PARTYPINCODE').text = partyPincode;
  SubElement(voucher, 'BASICBASEPARTYNAME').text = partyName;
  SubElement(voucher, 'PERSISTEDVIEW').text = 'Invoice Voucher View';
  SubElement(voucher, 'VCHENTRYMODE').text = 'Accounting Invoice';
  SubElement(voucher, 'ISINVOICE').text = 'Yes';
  SubElement(voucher, 'ISGSTOVERRIDDEN').text = 'No';
  SubElement(voucher, 'GSTTRANSACTIONTYPE').text = partyGstin ? 'Tax Invoice' : 'Unregistered';
  SubElement(voucher, 'EFFECTIVEDATE').text = tallyDt;
  SubElement(voucher, 'ISELIGIBLEFORITC').text = 'Yes';
  SubElement(voucher, 'NARRATION').text = narration;
  const vchNo = String(rec.voucher_no || '').trim();
  if (vchNo) SubElement(voucher, 'VOUCHERNUMBER').text = vchNo;

  const taxable = Math.round(parseFloat(rec.taxable_value || 0) * 100) / 100;
  const igstAmt = Math.round(parseFloat(rec.igst || 0) * 100) / 100;
  const cgstAmt = Math.round(parseFloat(rec.cgst || 0) * 100) / 100;
  const sgstAmt = Math.round(parseFloat(rec.sgst || 0) * 100) / 100;
  const cessAmt = Math.round(parseFloat(rec.cess || 0) * 100) / 100;

  const tdsLedger = rec.tds_ledger || '';
  let tdsRate;
  try { tdsRate = parseFloat(rec.tds_rate || 0); } catch { tdsRate = 0.0; }
  if (!Number.isFinite(tdsRate)) tdsRate = 0.0;
  let tdsAmount = rec.tds_amount;
  if ((tdsAmount === null || tdsAmount === undefined || tdsAmount === '') && tdsLedger && tdsRate > 0) {
    tdsAmount = Math.round(taxable * tdsRate / 100 * 100) / 100;
  } else {
    const rawTds = Math.round(Math.abs(parseFloat(tdsAmount || 0)) * 100) / 100;
    tdsAmount = Number.isFinite(rawTds) ? (taxable < 0 ? -rawTds : rawTds) : 0.0;
  }

  const totalAmount = Math.round((taxable + igstAmt + cgstAmt + sgstAmt + cessAmt) * 100) / 100;
  if (Math.abs(tdsAmount) > Math.abs(totalAmount)) tdsAmount = totalAmount;

  const splits = rec.purchase_ledger_splits || {};
  const roundedSplits = {};
  for (const [k, v] of Object.entries(splits)) {
    const rv = Math.round(parseFloat(v) * 100) / 100;
    if (rv !== 0) roundedSplits[k] = rv;
  }
  const hasRoundedSplits = Object.keys(roundedSplits).length > 0;
  const purchaseTotal = hasRoundedSplits ? Math.round(Object.values(roundedSplits).reduce((s, v) => s + v, 0) * 100) / 100 : taxable;
  const taxTotal = Math.round((igstAmt + cgstAmt + sgstAmt + cessAmt) * 100) / 100;
  const roBase = Math.round((purchaseTotal + taxTotal) * 100) / 100;
  const tdsVal = Math.round(parseFloat(tdsAmount || 0) * 100) / 100;
  const netPreRo = Math.round((roBase - tdsVal) * 100) / 100;
  const roAmtMain = roundOffLedger ? Math.round((Math.round(netPreRo) - netPreRo) * 100) / 100 : 0.0;
  const partyAmount = Math.round((netPreRo + roAmtMain) * 100) / 100;

  const partyDeemed = isDebitNote ? 'Yes' : 'No';
  const counterDeemed = isDebitNote ? 'No' : 'Yes';

  const pe = SubElement(voucher, 'LEDGERENTRIES.LIST');
  SubElement(pe, 'LEDGERNAME').text = partyLedger;
  addCommonLedgerFlags(pe, 'Yes');
  findChild(pe, 'ISDEEMEDPOSITIVE').text = partyDeemed;
  SubElement(pe, 'AMOUNT').text = fmt2(partyAmount);
  const ba = SubElement(pe, 'BILLALLOCATIONS.LIST');
  SubElement(ba, 'NAME').text = supplierInvoiceNo;
  SubElement(ba, 'BILLTYPE').text = 'New Ref';
  SubElement(ba, 'AMOUNT').text = fmt2(partyAmount);

  if (hasRoundedSplits) {
    for (const [ledgerName, ledgerAmount] of Object.entries(roundedSplits)) {
      const pu = SubElement(voucher, 'LEDGERENTRIES.LIST');
      SubElement(pu, 'LEDGERNAME').text = String(ledgerName);
      addCommonLedgerFlags(pu, 'No');
      findChild(pu, 'ISDEEMEDPOSITIVE').text = counterDeemed;
      SubElement(pu, 'AMOUNT').text = fmt2(-ledgerAmount);
    }
  } else {
    const pu = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(pu, 'LEDGERNAME').text = purchaseLedger;
    addCommonLedgerFlags(pu, 'No');
    findChild(pu, 'ISDEEMEDPOSITIVE').text = counterDeemed;
    SubElement(pu, 'AMOUNT').text = fmt2(-taxable);
  }

  const filterSplits = (obj) => {
    const out = {};
    for (const [k, v] of Object.entries(obj || {})) {
      const rv = Math.round(v * 100) / 100;
      if (Math.round(Math.abs(v) * 100) / 100 > 0) out[k] = rv;
    }
    return out;
  };
  const igstSplits = filterSplits(rec.igst_ledger_splits);
  const cgstSplits = filterSplits(rec.cgst_ledger_splits);
  const sgstSplits = filterSplits(rec.sgst_ledger_splits);

  const gstEntry = (lgr, amt, defaultName) => {
    const e = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(e, 'LEDGERNAME').text = lgr || defaultName;
    addCommonLedgerFlags(e, 'No');
    findChild(e, 'ISDEEMEDPOSITIVE').text = counterDeemed;
    SubElement(e, 'AMOUNT').text = fmt2(-amt);
  };

  if (Object.keys(igstSplits).length || Object.keys(cgstSplits).length || Object.keys(sgstSplits).length) {
    for (const [lgr, amt] of Object.entries(igstSplits)) gstEntry(lgr, amt, 'IGST');
    for (const [lgr, amt] of Object.entries(cgstSplits)) gstEntry(lgr, amt, 'CGST');
    for (const [lgr, amt] of Object.entries(sgstSplits)) gstEntry(lgr, amt, 'SGST');
  } else {
    const recRate = rec.rate;
    const gstLedgerRateMap = ctx.gst_ledger_rate_map || {};
    if (Math.abs(igstAmt) > 0) {
      const led = (rec.igst_ledger || '').trim() || getGstLedger(gstLedgerRateMap, 'igst', recRate) || 'IGST';
      gstEntry(led, igstAmt, 'IGST');
    } else {
      if (Math.abs(cgstAmt) > 0) {
        const led = (rec.cgst_ledger || '').trim() || getGstLedger(gstLedgerRateMap, 'cgst', recRate) || 'CGST';
        gstEntry(led, cgstAmt, 'CGST');
      }
      if (Math.abs(sgstAmt) > 0) {
        const led = (rec.sgst_ledger || '').trim() || getGstLedger(gstLedgerRateMap, 'sgst', recRate) || 'SGST';
        gstEntry(led, sgstAmt, 'SGST');
      }
    }
  }

  if (Math.abs(cessAmt) > 0) {
    const cs = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(cs, 'LEDGERNAME').text = 'Cess';
    addCommonLedgerFlags(cs, 'No');
    findChild(cs, 'ISDEEMEDPOSITIVE').text = counterDeemed;
    SubElement(cs, 'AMOUNT').text = fmt2(-cessAmt);
  }

  if (tdsLedger && Math.abs(tdsAmount) > 0) {
    const te = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(te, 'LEDGERNAME').text = tdsLedger;
    addCommonLedgerFlags(te, 'No');
    findChild(te, 'ISDEEMEDPOSITIVE').text = counterDeemed;
    SubElement(te, 'AMOUNT').text = fmt2(tdsAmount);
  }

  if (roundOffLedger && Math.abs(roAmtMain) >= 0.005) {
    const ro = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(ro, 'LEDGERNAME').text = roundOffLedger;
    addCommonLedgerFlags(ro, 'No');
    if (roAmtMain > 0) {
      findChild(ro, 'ISDEEMEDPOSITIVE').text = counterDeemed;
      SubElement(ro, 'AMOUNT').text = fmt2(-roAmtMain);
    } else {
      findChild(ro, 'ISDEEMEDPOSITIVE').text = partyDeemed;
      SubElement(ro, 'AMOUNT').text = fmt2(Math.abs(roAmtMain));
    }
  }
}

/** ITC-ineligible purchase: full amount, no tax breakup, no TDS. */
function buildJournalVoucherXml(parent, rec, purchaseLedger, narration, voucherDate, roundOffLedger, ctx) {
  const tallyMsg = SubElement(parent, 'TALLYMESSAGE');
  setAttr(tallyMsg, 'xmlns:UDF', 'TallyUDF');
  const voucher = SubElement(tallyMsg, 'VOUCHER');
  setAttr(voucher, 'REMOTEID', '');
  setAttr(voucher, 'VCHTYPE', 'Journal');
  setAttr(voucher, 'ACTION', 'Create');
  setAttr(voucher, 'OBJVIEW', 'Accounting Voucher View');

  const actualVoucherDate = rec.voucher_date || rec.invoice_date || voucherDate;
  const tallyDt = tallyDate(actualVoucherDate, true);
  const supplierInvoiceNo = rec.supplier_invoice_no || rec.invoice_no || '';
  const partyName = rec.party_name || rec.trade_name || '';

  const taxable = Math.round(parseFloat(rec.taxable_value || 0) * 100) / 100;
  const igstAmt = Math.round(parseFloat(rec.igst || 0) * 100) / 100;
  const cgstAmt = Math.round(parseFloat(rec.cgst || 0) * 100) / 100;
  const sgstAmt = Math.round(parseFloat(rec.sgst || 0) * 100) / 100;
  const cessAmt = Math.round(parseFloat(rec.cess || 0) * 100) / 100;

  SubElement(voucher, 'DATE').text = tallyDt;
  SubElement(voucher, 'VOUCHERTYPENAME').text = 'Journal';
  SubElement(voucher, 'PERSISTEDVIEW').text = 'Accounting Voucher View';
  SubElement(voucher, 'VCHENTRYMODE').text = 'Accounting Voucher View';
  SubElement(voucher, 'ISINVOICE').text = 'No';
  SubElement(voucher, 'EFFECTIVEDATE').text = tallyDt;
  SubElement(voucher, 'ISELIGIBLEFORITC').text = 'No';
  SubElement(voucher, 'ISGSTOVERRIDDEN').text = 'No';

  const partyGstin = String(rec.gstin || '').trim().toUpperCase();
  const [companyGstin, companyState] = companyRegistrationBlock(ctx, voucher, partyGstin);
  if (companyGstin && companyState) SubElement(voucher, 'PLACEOFSUPPLY').text = companyState;

  SubElement(voucher, 'NARRATION').text = narration;
  if (supplierInvoiceNo) SubElement(voucher, 'REFERENCE').text = String(supplierInvoiceNo);
  const vchNo = String(rec.voucher_no || '').trim();
  if (vchNo) SubElement(voucher, 'VOUCHERNUMBER').text = vchNo;

  const totalAmount = Math.round((taxable + igstAmt + cgstAmt + sgstAmt + cessAmt) * 100) / 100;
  if (totalAmount === 0) return;
  const roAmt = roundOffLedger ? Math.round((Math.round(totalAmount) - totalAmount) * 100) / 100 : 0.0;
  const roTotal = Math.round((totalAmount + roAmt) * 100) / 100;

  const debitPurchase = SubElement(voucher, 'LEDGERENTRIES.LIST');
  SubElement(debitPurchase, 'LEDGERNAME').text = purchaseLedger;
  addCommonLedgerFlags(debitPurchase, 'No');
  findChild(debitPurchase, 'ISDEEMEDPOSITIVE').text = 'Yes';
  SubElement(debitPurchase, 'AMOUNT').text = fmt2(-totalAmount);

  if (roundOffLedger && Math.abs(roAmt) >= 0.005) {
    const ro = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(ro, 'LEDGERNAME').text = roundOffLedger;
    addCommonLedgerFlags(ro, 'No');
    if (roAmt > 0) {
      findChild(ro, 'ISDEEMEDPOSITIVE').text = 'Yes';
      SubElement(ro, 'AMOUNT').text = fmt2(-roAmt);
    } else {
      findChild(ro, 'ISDEEMEDPOSITIVE').text = 'No';
      SubElement(ro, 'AMOUNT').text = fmt2(Math.abs(roAmt));
    }
  }

  const creditParty = SubElement(voucher, 'LEDGERENTRIES.LIST');
  SubElement(creditParty, 'LEDGERNAME').text = partyName;
  addCommonLedgerFlags(creditParty, 'Yes');
  findChild(creditParty, 'ISDEEMEDPOSITIVE').text = 'No';
  SubElement(creditParty, 'AMOUNT').text = fmt2(roTotal);
  if (supplierInvoiceNo) {
    const ba = SubElement(creditParty, 'BILLALLOCATIONS.LIST');
    SubElement(ba, 'NAME').text = String(supplierInvoiceNo);
    SubElement(ba, 'BILLTYPE').text = 'New Ref';
    SubElement(ba, 'AMOUNT').text = fmt2(roTotal);
  }
}

/** RCM Journal: Dr Expense + Dr GST Inward RCM, Cr Party + Cr GST Outward RCM. */
function buildRcmJournalVoucherXml(parent, rec, purchaseLedger, narration, voucherDate, roundOffLedger, ctx) {
  const rcm = ctx.rcm_ledger_map || {};
  const expenseLedger = rcm.expense || purchaseLedger;
  const partyLedgerName = rec.trade_name || '';
  const partyName = rec.party_name || partyLedgerName;
  const supplierInvoiceNo = rec.supplier_invoice_no || rec.invoice_no || '';

  const taxable = Math.round(parseFloat(rec.taxable_value || 0) * 100) / 100;
  const igstAmt = Math.round(parseFloat(rec.igst || 0) * 100) / 100;
  const cgstAmt = Math.round(parseFloat(rec.cgst || 0) * 100) / 100;
  const sgstAmt = Math.round(parseFloat(rec.sgst || 0) * 100) / 100;
  const isIgst = Math.abs(igstAmt) > 0;
  const hasGst = (Math.abs(igstAmt) + Math.abs(cgstAmt) + Math.abs(sgstAmt)) > 0;
  if (taxable === 0 && !hasGst) return;

  const roAmt = roundOffLedger ? Math.round((Math.round(taxable) - taxable) * 100) / 100 : 0.0;
  const partyCredit = Math.round((taxable + roAmt) * 100) / 100;

  const actualVoucherDate = rec.voucher_date || rec.invoice_date || voucherDate;
  const tallyDt = tallyDate(actualVoucherDate, true);
  const partyGstin = String(rec.gstin || '').trim().toUpperCase();
  const partyState = normalizeStateName(rec.party_state || stateFromGstin(partyGstin));

  const tallyMsg = SubElement(parent, 'TALLYMESSAGE');
  setAttr(tallyMsg, 'xmlns:UDF', 'TallyUDF');
  const voucher = SubElement(tallyMsg, 'VOUCHER');
  setAttr(voucher, 'REMOTEID', '');
  setAttr(voucher, 'VCHTYPE', 'Journal');
  setAttr(voucher, 'ACTION', 'Create');
  setAttr(voucher, 'OBJVIEW', 'Accounting Voucher View');

  SubElement(voucher, 'DATE').text = tallyDt;
  SubElement(voucher, 'VOUCHERTYPENAME').text = 'Journal';
  SubElement(voucher, 'PERSISTEDVIEW').text = 'Accounting Voucher View';
  SubElement(voucher, 'VCHENTRYMODE').text = 'Accounting Voucher View';
  SubElement(voucher, 'ISINVOICE').text = 'No';
  SubElement(voucher, 'EFFECTIVEDATE').text = tallyDt;

  const companyGstin = String(ctx.company_gstin || '').trim().toUpperCase();
  if (hasGst) {
    SubElement(voucher, 'ISGSTOVERRIDDEN').text = 'No';
    const [, companyState] = companyRegistrationBlock(ctx, voucher, partyGstin);
    if (companyGstin && companyState) SubElement(voucher, 'PLACEOFSUPPLY').text = companyState;
  }

  SubElement(voucher, 'PARTYNAME').text = partyName;
  SubElement(voucher, 'ISELIGIBLEFORITC').text = 'Yes';
  SubElement(voucher, 'NARRATION').text = narration;
  if (supplierInvoiceNo) SubElement(voucher, 'REFERENCE').text = String(supplierInvoiceNo);
  const vchNo = String(rec.voucher_no || '').trim();
  if (vchNo) SubElement(voucher, 'VOUCHERNUMBER').text = vchNo;

  const dr = (ledgerName, amount) => {
    const e = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(e, 'LEDGERNAME').text = ledgerName;
    addCommonLedgerFlags(e, 'No');
    findChild(e, 'ISDEEMEDPOSITIVE').text = 'Yes';
    SubElement(e, 'AMOUNT').text = fmt2(-Math.abs(amount));
  };
  const crParty = (ledgerName, amount, invNo = '') => {
    const e = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(e, 'LEDGERNAME').text = ledgerName;
    addCommonLedgerFlags(e, 'Yes');
    findChild(e, 'ISDEEMEDPOSITIVE').text = 'No';
    if (hasGst) {
      SubElement(e, 'GSTREGISTRATIONTYPE').text = 'Regular';
      if (partyGstin) SubElement(e, 'GSTIN').text = partyGstin;
      if (partyState) SubElement(e, 'STATENAME').text = partyState;
      SubElement(e, 'COUNTRYOFRESIDENCE').text = 'India';
    }
    SubElement(e, 'AMOUNT').text = fmt2(Math.abs(amount));
    if (invNo) {
      const ba = SubElement(e, 'BILLALLOCATIONS.LIST');
      SubElement(ba, 'NAME').text = String(invNo);
      SubElement(ba, 'BILLTYPE').text = 'New Ref';
      SubElement(ba, 'AMOUNT').text = fmt2(Math.abs(amount));
    }
  };
  const cr = (ledgerName, amount) => {
    const e = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(e, 'LEDGERNAME').text = ledgerName;
    addCommonLedgerFlags(e, 'No');
    findChild(e, 'ISDEEMEDPOSITIVE').text = 'No';
    SubElement(e, 'AMOUNT').text = fmt2(Math.abs(amount));
  };

  if (taxable) dr(expenseLedger, taxable);

  if (isIgst) {
    if (Math.abs(igstAmt) > 0) dr(rcm.igst_inward || 'IGST Inward RCM', igstAmt);
  } else {
    if (Math.abs(cgstAmt) > 0) dr(rcm.cgst_inward || 'CGST Inward RCM', cgstAmt);
    if (Math.abs(sgstAmt) > 0) dr(rcm.sgst_inward || 'SGST Inward RCM', sgstAmt);
  }

  if (roundOffLedger && Math.abs(roAmt) >= 0.005 && roAmt > 0) {
    const ro = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(ro, 'LEDGERNAME').text = roundOffLedger;
    addCommonLedgerFlags(ro, 'No');
    findChild(ro, 'ISDEEMEDPOSITIVE').text = 'Yes';
    SubElement(ro, 'AMOUNT').text = fmt2(-roAmt);
  }

  if (partyCredit) crParty(partyLedgerName, partyCredit, supplierInvoiceNo);

  if (isIgst) {
    if (Math.abs(igstAmt) > 0) cr(rcm.igst_outward || 'IGST Outward RCM', igstAmt);
  } else {
    if (Math.abs(cgstAmt) > 0) cr(rcm.cgst_outward || 'CGST Outward RCM', cgstAmt);
    if (Math.abs(sgstAmt) > 0) cr(rcm.sgst_outward || 'SGST Outward RCM', sgstAmt);
  }

  if (roundOffLedger && Math.abs(roAmt) >= 0.005 && roAmt < 0) {
    const ro = SubElement(voucher, 'LEDGERENTRIES.LIST');
    SubElement(ro, 'LEDGERNAME').text = roundOffLedger;
    addCommonLedgerFlags(ro, 'No');
    findChild(ro, 'ISDEEMEDPOSITIVE').text = 'No';
    SubElement(ro, 'AMOUNT').text = fmt2(Math.abs(roAmt));
  }
}

function buildVoucherXml(parent, rec, purchaseLedger, narration, voucherDate, roundOffLedger, ctx) {
  if (String(rec.reverse_charge || '').trim().toUpperCase() === 'YES') {
    buildRcmJournalVoucherXml(parent, rec, purchaseLedger, narration, voucherDate, roundOffLedger, ctx);
    return;
  }
  const itcStatus = String(rec.itc_avail || 'Yes').trim().toUpperCase();
  if (['NO', 'N', 'INELIGIBLE'].includes(itcStatus)) {
    buildJournalVoucherXml(parent, rec, purchaseLedger, narration, voucherDate, roundOffLedger, ctx);
    return;
  }
  buildPurchaseVoucherXml(parent, rec, purchaseLedger, narration, voucherDate, roundOffLedger, ctx);
}

function formatNarrationTemplate(template, params) {
  // Mirrors Python's str.format(party=..., inv=..., date=...) for the
  // specific {party}/{inv}/{date} placeholders this template always uses.
  return template.replace(/\{(\w+)\}/g, (match, key) => (key in params ? params[key] : match));
}

function generateGstr2bXml(records, opts = {}) {
  const {
    companyName = '', companyGstin = '', companyRegistrationName = '', companyRegistrationState = '',
    purchaseLedger = 'Purchase Account', narrationTemplate = 'Being purchase from {party} vide Inv {inv} dt {date}',
    roundOffLedger = '', rcmLedgerMap = null, gstLedgerRateMap = null, partyLedgerMap = null,
    useInvoiceNoAsVoucherNo = false,
  } = opts;
  const ctx = {
    company_gstin: companyGstin,
    company_registration_name: companyRegistrationName,
    company_registration_state: companyRegistrationState,
    rcm_ledger_map: rcmLedgerMap || {},
    gst_ledger_rate_map: gstLedgerRateMap || {},
  };
  const consolidated = consolidateInvoiceRecords(records, purchaseLedger, partyLedgerMap, gstLedgerRateMap);

  const envelope = Element('ENVELOPE');
  const header = SubElement(envelope, 'HEADER');
  SubElement(header, 'TALLYREQUEST').text = 'Import Data';
  const body = SubElement(envelope, 'BODY');
  const importData = SubElement(body, 'IMPORTDATA');
  const reqDesc = SubElement(importData, 'REQUESTDESC');
  SubElement(reqDesc, 'REPORTNAME').text = 'Vouchers';
  const staticVars = SubElement(reqDesc, 'STATICVARIABLES');
  SubElement(staticVars, 'SVCURRENTCOMPANY').text = companyName || 'My Company';
  const reqData = SubElement(importData, 'REQUESTDATA');

  const today = new Date();
  const todayStr = `${pad2(today.getDate())}/${pad2(today.getMonth() + 1)}/${today.getFullYear()}`;

  for (const rec of consolidated) {
    const trade = rec.trade_name || rec.party_name || '';
    const inv = rec.invoice_no || '';
    const date_ = rec.invoice_date || '';
    const recLedger = rec.purchase_ledger || getPurchaseLedger(partyLedgerMap || {}, trade, purchaseLedger);
    let recNarration;
    try {
      recNarration = rec.narration || formatNarrationTemplate(narrationTemplate, { party: trade, inv, date: date_ });
    } catch {
      recNarration = `Being purchase from ${trade} vide Inv ${inv} dt ${date_}`;
    }
    if (rec.is_amendment && rec.orig_invoice_no) recNarration += ` [Amends ${rec.orig_invoice_no}]`;
    if (useInvoiceNoAsVoucherNo && !String(rec.voucher_no || '').trim()) rec.voucher_no = inv;
    buildVoucherXml(reqData, rec, recLedger, recNarration, todayStr, roundOffLedger, ctx);
  }

  const prettyBody = toPrettyXml(envelope);
  return `<?xml version="1.0" encoding="UTF-8"?>\n${prettyBody}\n`;
}

/** Single-voucher manual builder — opened one-per-record for every "has
 * stock item = Yes" ITC template row (item mode) AND for every tax-slab-
 * mismatch record when the user chooses to keep mismatched records instead
 * of dropping them (accounting mode). Always VCHTYPE=Purchase. Serializes
 * with ET.tostring's compact format (no pretty-printing), unlike generateGstr2bXml. */
function buildManualPurchaseVoucherXml(
  invNo, dateText, partyName, gstin, taxEntries, opts = {}
) {
  const {
    itemEntries = [], ledgerEntries = [], narration = '', companyName = '', companyGstin = '',
    companyRegistrationName = '', companyRegistrationState = '', useExistingSeries = true,
  } = opts;
  const isItemMode = itemEntries.length > 0;

  const tallyDt = tallyDate(dateText);
  const gstNo = (gstin || '').trim().toUpperCase();
  const partyState = gstNo ? stateFromGstin(gstNo) : '';
  const cmpGstin = (companyGstin || '').trim().toUpperCase();
  const cmpState = companyRegistrationState || (cmpGstin ? stateFromGstin(cmpGstin) : '');
  const regName = companyRegistrationName || (cmpState ? `${cmpState} Registration` : cmpGstin);

  const taxTotal = Math.round(taxEntries.reduce((s, t) => s + parseFloat(t.amount || 0), 0) * 100) / 100;
  let bodyTotal;
  if (isItemMode) {
    bodyTotal = Math.round(itemEntries.reduce((s, it) => s + parseFloat(it.amount || 0), 0) * 100) / 100;
  } else {
    bodyTotal = Math.round(ledgerEntries.reduce((s, e) => s + Math.abs(parseFloat(e.amount || 0)), 0) * 100) / 100;
  }
  const partyAmount = Math.round((bodyTotal + taxTotal) * 100) / 100;

  const envelope = Element('ENVELOPE');
  const header = SubElement(envelope, 'HEADER');
  SubElement(header, 'TALLYREQUEST').text = 'Import Data';
  const body = SubElement(envelope, 'BODY');
  const importData = SubElement(body, 'IMPORTDATA');
  const reqDesc = SubElement(importData, 'REQUESTDESC');
  SubElement(reqDesc, 'REPORTNAME').text = 'Vouchers';
  const sv = SubElement(reqDesc, 'STATICVARIABLES');
  SubElement(sv, 'SVEXPORTFORMAT').text = '$$SysName:XML';
  if (companyName) SubElement(sv, 'SVCURRENTCOMPANY').text = companyName;
  const reqData = SubElement(importData, 'REQUESTDATA');

  const tm = SubElement(reqData, 'TALLYMESSAGE');
  setAttr(tm, 'xmlns:UDF', 'TallyUDF');
  const v = SubElement(tm, 'VOUCHER');
  setAttr(v, 'REMOTEID', '');
  setAttr(v, 'VCHTYPE', 'Purchase');
  setAttr(v, 'ACTION', 'Create');
  setAttr(v, 'OBJVIEW', 'Invoice Voucher View');

  SubElement(v, 'DATE').text = tallyDt;
  SubElement(v, 'REFERENCEDATE').text = tallyDt;
  SubElement(v, 'GSTREGISTRATIONTYPE').text = 'Regular';
  SubElement(v, 'VATDEALERTYPE').text = 'Regular';
  if (partyState) SubElement(v, 'STATENAME').text = partyState;
  SubElement(v, 'COUNTRYOFRESIDENCE').text = 'India';
  if (gstNo) SubElement(v, 'PARTYGSTIN').text = gstNo;
  if (partyState) SubElement(v, 'PLACEOFSUPPLY').text = partyState;
  SubElement(v, 'VOUCHERTYPENAME').text = 'Purchase';
  SubElement(v, 'PARTYNAME').text = partyName;
  if (cmpGstin) {
    const gstReg = SubElement(v, 'GSTREGISTRATION');
    setAttr(gstReg, 'TAXTYPE', 'GST');
    setAttr(gstReg, 'TAXREGISTRATION', cmpGstin);
    gstReg.text = regName;
    SubElement(v, 'CMPGSTIN').text = cmpGstin;
    SubElement(v, 'CMPGSTREGISTRATIONTYPE').text = 'Regular';
    if (cmpState) SubElement(v, 'CMPGSTSTATE').text = cmpState;
  }
  SubElement(v, 'PARTYLEDGERNAME').text = partyName;
  SubElement(v, 'REFERENCE').text = invNo;
  if (!useExistingSeries) SubElement(v, 'VOUCHERNUMBER').text = invNo;
  SubElement(v, 'PARTYMAILINGNAME').text = partyName;
  SubElement(v, 'BASICBASEPARTYNAME').text = partyName;
  SubElement(v, 'PERSISTEDVIEW').text = 'Invoice Voucher View';
  SubElement(v, 'VCHENTRYMODE').text = isItemMode ? 'Item Invoice' : 'Accounting Invoice';
  SubElement(v, 'ISINVOICE').text = 'Yes';
  SubElement(v, 'ISGSTOVERRIDDEN').text = 'No';
  SubElement(v, 'GSTTRANSACTIONTYPE').text = gstNo ? 'Tax Invoice' : 'Unregistered';
  SubElement(v, 'ISELIGIBLEFORITC').text = 'Yes';
  SubElement(v, 'EFFECTIVEDATE').text = tallyDt;
  SubElement(v, 'NARRATION').text = narration || '';

  const le = (name, deemed, amount, isParty = 'No') => {
    const node = SubElement(v, 'LEDGERENTRIES.LIST');
    SubElement(node, 'LEDGERNAME').text = name;
    SubElement(node, 'GSTCLASS').text = 'Not Applicable';
    SubElement(node, 'ISDEEMEDPOSITIVE').text = deemed;
    SubElement(node, 'LEDGERFROMITEM').text = 'No';
    SubElement(node, 'REMOVEZEROENTRIES').text = 'No';
    SubElement(node, 'ISPARTYLEDGER').text = isParty;
    SubElement(node, 'GSTOVERRIDDEN').text = 'No';
    SubElement(node, 'ISGSTASSESSABLEVALUEOVERRIDDEN').text = 'No';
    SubElement(node, 'AMOUNT').text = fmt2(amount);
    return node;
  };

  const pe = le(partyName, 'No', partyAmount, 'Yes');
  const ba = SubElement(pe, 'BILLALLOCATIONS.LIST');
  SubElement(ba, 'NAME').text = invNo;
  SubElement(ba, 'BILLTYPE').text = 'New Ref';
  SubElement(ba, 'AMOUNT').text = fmt2(partyAmount);

  if (isItemMode) {
    for (const t of taxEntries) {
      const amt = Math.round(parseFloat(t.amount || 0) * 100) / 100;
      if (amt) le(t.ledger, 'Yes', -amt);
    }

    for (const itm of itemEntries) {
      const amt = Math.round(parseFloat(itm.amount || 0) * 100) / 100;
      const qty = parseFloat(itm.qty || 0);
      const rate = parseFloat(itm.rate || 0);
      const unit = itm.unit || 'Nos';
      const qtyStr = formatG(qty);
      const invEntry = SubElement(v, 'ALLINVENTORYENTRIES.LIST');
      SubElement(invEntry, 'STOCKITEMNAME').text = itm.name;
      SubElement(invEntry, 'ISDEEMEDPOSITIVE').text = 'Yes';
      SubElement(invEntry, 'RATE').text = `${fmt2(rate)}/${unit}`;
      SubElement(invEntry, 'AMOUNT').text = fmt2(-amt);
      SubElement(invEntry, 'ACTUALQTY').text = `${qtyStr} ${unit}`;
      SubElement(invEntry, 'BILLEDQTY').text = `${qtyStr} ${unit}`;
      if (itm.hsn) SubElement(invEntry, 'HSNCODE').text = itm.hsn;
      const baInv = SubElement(invEntry, 'BATCHALLOCATIONS.LIST');
      SubElement(baInv, 'GODOWNNAME').text = 'Main Location';
      SubElement(baInv, 'BATCHNAME').text = 'Primary Batch';
      SubElement(baInv, 'AMOUNT').text = fmt2(-amt);
      SubElement(baInv, 'ACTUALQTY').text = `${qtyStr} ${unit}`;
      SubElement(baInv, 'BILLEDQTY').text = `${qtyStr} ${unit}`;
      const aa = SubElement(invEntry, 'ACCOUNTINGALLOCATIONS.LIST');
      SubElement(aa, 'LEDGERNAME').text = itm.ledger;
      SubElement(aa, 'ISDEEMEDPOSITIVE').text = 'Yes';
      SubElement(aa, 'AMOUNT').text = fmt2(-amt);
    }
  } else {
    for (const e of ledgerEntries) {
      const amt = Math.round(Math.abs(parseFloat(e.amount || 0)) * 100) / 100;
      if (amt) le(e.ledger, 'Yes', -amt);
    }
    for (const t of taxEntries) {
      const amt = Math.round(Math.abs(parseFloat(t.amount || 0)) * 100) / 100;
      if (amt) le(t.ledger, 'Yes', -amt);
    }
  }

  const rough = toCompactXml(envelope);
  return `<?xml version="1.0" encoding="UTF-8"?>${rough}`;
}

/** Mirrors Python's f"{qty:g}" (%g, 6 significant digits): fixed notation
 * trimmed of trailing zeros, switching to exponential (with 2-digit signed
 * exponent) once the magnitude exponent is <-4 or >=6. */
function formatG(n, precision = 6) {
  if (n === 0) return '0';
  const exp = Math.floor(Math.log10(Math.abs(n)));
  if (exp < -4 || exp >= precision) {
    let s = n.toExponential(precision - 1);
    let [mantissa, exponent] = s.split('e');
    if (mantissa.includes('.')) mantissa = mantissa.replace(/0+$/, '').replace(/\.$/, '');
    const expNum = parseInt(exponent, 10);
    const expStr = (expNum >= 0 ? '+' : '-') + String(Math.abs(expNum)).padStart(2, '0');
    return `${mantissa}e${expStr}`;
  }
  const decimals = Math.max(0, precision - 1 - exp);
  let s = n.toFixed(decimals);
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s;
}

// ─── ITC template (download a pre-filled mapping sheet, upload it back) ────

const ITC_YELLOW_HEADERS = [
  'GSTIN', 'Trade/Legal name', 'Invoice Number', 'Invoice Date', 'Invoice Type',
  'Invoice Value', 'Taxable Value',
  'CGST', 'CGST Rate', 'SGST', 'SGST Rate', 'IGST', 'IGST Rate', 'CESS',
  'ITC Availability', 'Supply Attract Reverse Charge',
];
const ITC_GREEN_HEADERS = [
  'ITC to be claimed or not', 'Whether Contains stock Item',
  'Mapping Ledger', 'CGST Ledger', 'SGST Ledger', 'IGST Ledger',
  'TDS Ledger', 'TDS Rate',
];
const ITC_COL_WIDTHS = [20, 32, 22, 14, 20, 16, 14, 10, 10, 10, 10, 10, 10, 10, 24, 30, 26, 28, 26, 26, 26, 26, 14, 10];

function colLetter(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function buildItcB2bSheet(destWs, srcWb, allSrcUpper) {
  const srcWs = allSrcUpper.B2B;
  if (!srcWs) return 0;

  const detected = detectColumns(srcWs);
  const columnMap = detected.columnMap;
  const dataStart = detected.dataStartRow;
  const hasRateCol = 'rate' in columnMap;

  const yellowHeaders = [...ITC_YELLOW_HEADERS, ...(hasRateCol ? ['Rate (%)'] : [])];
  const greenHeaders = ITC_GREEN_HEADERS;
  const allHeaders = [...yellowHeaders, ...greenHeaders];
  const nYellow = yellowHeaders.length;
  const dvColS = nYellow + 1;
  const dvColE = nYellow + 2;
  const colWidths = [...ITC_COL_WIDTHS.slice(0, 16), ...(hasRateCol ? [10] : []), ...ITC_COL_WIDTHS.slice(16)];

  const yellowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  const greenFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
  const headerFont = { name: 'Calibri', size: 11, bold: true };
  const headerAlign = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const dataFont = { name: 'Calibri', size: 10 };
  const numAlign = { horizontal: 'right' };

  const nCols = allHeaders.length;
  const msgCell = destWs.getCell(1, 1);
  msgCell.value = '*Kindly Map the GST Ledgers in case your tally has multiple rates of Input GST Ledger';
  msgCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFF0000' } };
  msgCell.alignment = { horizontal: 'center', vertical: 'middle' };
  destWs.mergeCells(1, 1, 1, nCols);
  destWs.getRow(1).height = 22;

  allHeaders.forEach((header, idx) => {
    const colIdx = idx + 1;
    const cell = destWs.getCell(2, colIdx);
    cell.value = header;
    cell.font = headerFont;
    cell.alignment = headerAlign;
    cell.fill = colIdx <= nYellow ? yellowFill : greenFill;
  });
  destWs.getRow(2).height = 32;

  let outRow = 3;
  let lastGstin = '', lastTradeName = '';
  const tradeCol = columnMap.trade_name !== undefined ? columnMap.trade_name : 1;
  const gstinCol = columnMap.gstin !== undefined ? columnMap.gstin : 0;

  for (let rowIdx = dataStart; rowIdx <= srcWs.rowCount; rowIdx++) {
    const wsRow = srcWs.getRow(rowIdx);
    let row = [];
    for (let c = 1; c <= srcWs.columnCount; c++) row.push(wsRow.getCell(c).value);
    const curGstin = gstinCol < row.length ? String(row[gstinCol] || '').trim() : '';
    if (curGstin && curGstin.length >= 15) {
      lastGstin = curGstin;
      lastTradeName = tradeCol < row.length ? String(row[tradeCol] || '').trim() : '';
    } else if (!curGstin && lastGstin) {
      row = row.slice();
      while (row.length <= Math.max(gstinCol, tradeCol)) row.push(null);
      row[gstinCol] = lastGstin;
      if (!row[tradeCol]) row[tradeCol] = lastTradeName;
      const invCol = columnMap.invoice_no;
      const hasInv = invCol !== undefined && invCol < row.length && String(row[invCol] || '').trim();
      if (!hasInv) {
        const amtCols = ['taxable_value', 'igst', 'cgst', 'sgst', 'cess'].map((f) => columnMap[f]);
        const hasAmounts = amtCols.some((c) => c !== undefined && c < row.length && row[c] !== null && row[c] !== 0 && row[c] !== '');
        if (!hasAmounts) continue;
      }
    } else {
      continue;
    }

    let rec;
    try {
      rec = parseB2bRow(row, rowIdx, columnMap);
    } catch {
      continue;
    }
    if (!rec) continue;

    const tv = Math.abs(rec.taxable_value || 0);
    const taxRate = (amt) => {
      const v = Math.abs(amt || 0);
      return (tv > 0 && v > 0) ? Math.round((v / tv * 100) * 100) / 100 : 0;
    };

    const rowValues = [
      rec.gstin || '', rec.trade_name || '', rec.invoice_no || '',
      rec.invoice_date || '', rec.invoice_type || 'Regular',
      rec.invoice_value || 0, rec.taxable_value || 0,
      rec.cgst || 0, taxRate(rec.cgst),
      rec.sgst || 0, taxRate(rec.sgst),
      rec.igst || 0, taxRate(rec.igst),
      rec.cess || 0, rec.itc_avail || '', rec.reverse_charge || 'No',
    ];
    if (hasRateCol) {
      const rateColIdx = columnMap.rate;
      const rawRate = (rateColIdx !== undefined && rateColIdx < row.length) ? row[rateColIdx] : '';
      rowValues.push(rawRate);
    }
    rowValues.forEach((value, idx) => {
      const cell = destWs.getCell(outRow, idx + 1);
      cell.value = value;
      cell.font = dataFont;
      if (typeof value === 'number') cell.alignment = numAlign;
    });
    outRow += 1;
  }

  const lastRow = Math.max(outRow - 1, 2);
  destWs.dataValidations.add(`${colLetter(dvColS)}3:${colLetter(dvColE)}${lastRow}`, {
    type: 'list', allowBlank: true, formulae: ['"Yes,No"'],
    showErrorMessage: true, errorTitle: 'Invalid Value',
    error: "Only 'Yes' or 'No' is allowed. Please select from the dropdown.", errorStyle: 'stop',
  });

  colWidths.forEach((width, idx) => { destWs.getColumn(idx + 1).width = width; });
  destWs.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
  return lastRow - 2;
}

function copySheetFormatted(srcWs, destWs) {
  srcWs.columns?.forEach((col, idx) => {
    if (col.width) destWs.getColumn(idx + 1).width = col.width;
    if (col.hidden) destWs.getColumn(idx + 1).hidden = col.hidden;
  });
  srcWs.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const destRow = destWs.getRow(rowNumber);
    if (row.height) destRow.height = row.height;
    if (row.hidden) destRow.hidden = row.hidden;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const destCell = destWs.getCell(rowNumber, colNumber);
      destCell.value = cell.value;
      if (cell.style && Object.keys(cell.style).length) {
        destCell.font = cell.font;
        destCell.fill = cell.fill;
        destCell.border = cell.border;
        destCell.alignment = cell.alignment;
        destCell.numFmt = cell.numFmt;
      }
    });
  });
  for (const merge of srcWs.model.merges || []) destWs.mergeCells(merge);
  destWs.views = srcWs.views;
}

/** Rebuilds the ITC template workbook (ITC B2B + B2BA + B2B CDNR tabs)
 * from the originally-uploaded GSTR-2B Excel file's raw bytes. */
async function generateItcTemplateWorkbook(sourceContent) {
  const srcWb = new ExcelJS.Workbook();
  await srcWb.xlsx.load(sourceContent);
  const allSrcUpper = {};
  for (const ws of srcWb.worksheets) allSrcUpper[ws.name.toUpperCase()] = ws;

  const wb = new ExcelJS.Workbook();
  const wsB2b = wb.addWorksheet('ITC B2B');
  buildItcB2bSheet(wsB2b, srcWb, allSrcUpper);

  const wsB2ba = wb.addWorksheet('B2BA');
  if (allSrcUpper.B2BA) copySheetFormatted(allSrcUpper.B2BA, wsB2ba);

  const wsCdnr = wb.addWorksheet('B2B CDNR');
  if (allSrcUpper['B2B-CDNR']) copySheetFormatted(allSrcUpper['B2B-CDNR'], wsCdnr);

  return wb.xlsx.writeBuffer();
}

function itcCell(row, idx, defaultValue = '') {
  return (idx !== undefined && idx !== null && idx >= 0 && idx < row.length && row[idx] !== null && row[idx] !== undefined)
    ? row[idx] : defaultValue;
}

function itcStr(row, idx, defaultValue = '') {
  const v = itcCell(row, idx, defaultValue);
  return String(v || '').trim();
}

function itcFloat(row, idx, defaultValue = 0.0) {
  const v = itcCell(row, idx, null);
  if (v === null || v === '') return defaultValue;
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : defaultValue;
}

function itcNormDate(raw) {
  if (raw instanceof Date) {
    return `${pad2(raw.getDate())}/${pad2(raw.getMonth() + 1)}/${raw.getFullYear()}`;
  }
  return raw ? normalizeDateStr(raw) : '';
}

/** Re-parses an uploaded ITC template (ITC B2B + optional B2BA / B2B CDNR
 * tabs). Returns {success, records, party_ledger_map, party_tds_ledger_map,
 * party_tds_rate_map, itc_map, errors, warnings, b2b_count, b2ba_count,
 * cdnr_count}. This REPLACES the working record set, same as the desktop
 * tool — it does not merge with whatever was parsed from the original file. */
async function parseItcTemplateWorkbook(content) {
  let wb;
  try {
    wb = new ExcelJS.Workbook();
    await wb.xlsx.load(content);
  } catch (exc) {
    return { success: false, records: [], errors: [`Failed to open file: ${exc.message}`], warnings: [] };
  }

  const ws = wb.getWorksheet('ITC B2B') || wb.worksheets[0];

  const row1 = ws.getRow(1);
  const row1Vals = [];
  for (let c = 1; c <= ws.columnCount; c++) row1Vals.push(row1.getCell(c).value);
  const row1A = row1Vals.length ? String(row1Vals[0] || '').trim().toLowerCase() : '';
  const hdrRow = row1A === 'gstin' ? 1 : 2;
  const dataStart = hdrRow + 1;

  const hdrRowObj = ws.getRow(hdrRow);
  const hdrRowVals = [];
  for (let c = 1; c <= ws.columnCount; c++) hdrRowVals.push(hdrRowObj.getCell(c).value);
  const hdr = hdrRowVals.map((c) => String(c || '').trim().toLowerCase());
  const hdrMap = {};
  hdr.forEach((h, i) => { if (!(h in hdrMap)) hdrMap[h] = i; });

  const hcol = (name, fallback = -1) => (name in hdrMap ? hdrMap[name] : fallback);

  let colRatePct = -1;
  for (let i = 0; i < hdr.length; i++) {
    const h = hdr[i];
    if (['rate (%)', 'rate(%)', 'rate'].includes(h) && !h.includes('tds') && !h.includes('cgst') && !h.includes('sgst') && !h.includes('igst')) {
      colRatePct = i;
      break;
    }
  }
  const hasRateCol = colRatePct >= 0;
  const gOff = hasRateCol ? 1 : 0;

  const colTaxable = hcol('taxable value', 5);
  const colCgst = hcol('cgst', 6);
  const colCgstRate = hcol('cgst rate', -1);
  const colSgst = hcol('sgst', 7);
  const colSgstRate = hcol('sgst rate', -1);
  const colIgst = hcol('igst', 8);
  const colIgstRate = hcol('igst rate', -1);
  const colCess = hcol('cess', 9);
  const colItcAvail = hcol('itc availability', 10);
  const colRevCharge = hcol('supply attract reverse charge', 11);

  let colItcClaimed = hcol('itc to be claimed or not');
  let colHasStock = hcol('whether contains stock item');
  let colMapping = hcol('mapping ledger');
  const colCgstLdr = hcol('cgst ledger');
  const colSgstLdr = hcol('sgst ledger');
  const colIgstLdr = hcol('igst ledger');
  let colTdsLdr = hcol('tds ledger');
  let colTdsRateC = hcol('tds rate');
  if (colItcClaimed < 0) colItcClaimed = 12 + gOff;
  if (colHasStock < 0) colHasStock = 13 + gOff;
  if (colMapping < 0) colMapping = 14 + gOff;
  if (colTdsLdr < 0) colTdsLdr = 15 + gOff;
  if (colTdsRateC < 0) colTdsRateC = 16 + gOff;
  const hasGstLedgerCols = colCgstLdr >= 0;

  const records = [];
  const partyLedger = {}, partyTdsLdr = {}, partyTdsRt = {}, itcMap = {};

  for (let rowIdx = dataStart; rowIdx <= ws.rowCount; rowIdx++) {
    const wsRow = ws.getRow(rowIdx);
    const row = [];
    for (let c = 1; c <= ws.columnCount; c++) row.push(wsRow.getCell(c).value);
    if (!row.length) continue;
    const gstin = itcStr(row, 0);
    const tradeName = itcStr(row, 1);
    if ((!gstin || gstin.length < 15) && !tradeName) continue;
    const invoiceNo = itcStr(row, 2);
    const invoiceDate = itcNormDate(itcCell(row, 3, ''));
    const invoiceType = itcStr(row, 4) || 'Regular';
    const taxable = itcFloat(row, colTaxable);
    let cgst = itcFloat(row, colCgst);
    let sgst = itcFloat(row, colSgst);
    let igst = itcFloat(row, colIgst);
    const cess = itcFloat(row, colCess);

    if (!cgst && taxable && colCgstRate >= 0) {
      const cr = itcFloat(row, colCgstRate);
      if (cr > 0) cgst = Math.round(Math.abs(taxable) * cr / 100 * 100) / 100;
    }
    if (!sgst && taxable && colSgstRate >= 0) {
      const sr = itcFloat(row, colSgstRate);
      if (sr > 0) sgst = Math.round(Math.abs(taxable) * sr / 100 * 100) / 100;
    }
    if (!igst && taxable && colIgstRate >= 0) {
      const ir = itcFloat(row, colIgstRate);
      if (ir > 0) igst = Math.round(Math.abs(taxable) * ir / 100 * 100) / 100;
    }

    const itcAvailSrc = itcStr(row, colItcAvail, 'Yes');
    const revCharge = itcStr(row, colRevCharge, 'No');
    const itcClaimed = itcStr(row, colItcClaimed);
    const hasStock = itcStr(row, colHasStock);
    const mappingLedger = itcStr(row, colMapping);
    let cgstLedger = '', sgstLedger = '', igstLedger = '';
    if (hasGstLedgerCols) {
      cgstLedger = colCgstLdr >= 0 ? itcStr(row, colCgstLdr) : '';
      sgstLedger = colSgstLdr >= 0 ? itcStr(row, colSgstLdr) : '';
      igstLedger = colIgstLdr >= 0 ? itcStr(row, colIgstLdr) : '';
    }
    const tdsLedger = colTdsLdr >= 0 ? itcStr(row, colTdsLdr) : '';
    const tdsRateRaw = colTdsRateC >= 0 ? itcCell(row, colTdsRateC, null) : null;
    let tdsRate;
    try { tdsRate = (tdsRateRaw !== null && tdsRateRaw !== '') ? parseFloat(tdsRateRaw) : 0.0; } catch { tdsRate = 0.0; }
    if (!Number.isFinite(tdsRate)) tdsRate = 0.0;

    const itcUpper = itcClaimed.toUpperCase();
    let effItc;
    if (['YES', 'Y'].includes(itcUpper)) effItc = 'Yes';
    else if (['NO', 'N'].includes(itcUpper)) effItc = 'Ineligible';
    else effItc = itcAvailSrc || 'Yes';

    let rate;
    if (hasRateCol) {
      const rateRaw = itcCell(row, colRatePct, null);
      try { rate = (rateRaw !== null && rateRaw !== '') ? parseFloat(rateRaw) : 0.0; } catch { rate = 0.0; }
      if (!Number.isFinite(rate)) rate = 0.0;
    } else {
      const totalTax = Math.abs(igst) + Math.abs(cgst) + Math.abs(sgst);
      const absTaxable = Math.abs(taxable);
      rate = (absTaxable > 0 && totalTax > 0) ? Math.round((totalTax / absTaxable) * 100) : 0.0;
    }

    const rec = {
      gstin, trade_name: tradeName, party_name: tradeName,
      invoice_no: invoiceNo, invoice_date: invoiceDate, invoice_type: invoiceType,
      taxable_value: taxable, cgst, sgst, igst, cess,
      rate, itc_avail: effItc, reverse_charge: revCharge,
      has_stock_item: ['YES', 'Y'].includes(hasStock.toUpperCase()),
      cgst_ledger: cgstLedger, sgst_ledger: sgstLedger, igst_ledger: igstLedger,
      tds_ledger: tdsLedger, tds_rate: tdsRate, tds_amount: '',
      purchase_ledger: mappingLedger,
      sheet_type: 'B2B', is_amendment: false, orig_invoice_no: '',
      filing_period: '', place_of_supply: '', party_state: '',
      party_mailing_name: tradeName, party_address1: '', party_address2: '',
      party_pincode: '', voucher_date: '', voucher_no: '',
      supplier_invoice_no: invoiceNo, supplier_invoice_date: invoiceDate,
      narration: '', row_idx: rowIdx,
    };
    records.push(rec);

    const key = tradeName.toUpperCase();
    if (tradeName && mappingLedger) partyLedger[key] = mappingLedger;
    if (tradeName && tdsLedger) partyTdsLdr[key] = tdsLedger;
    if (tradeName && tdsRate) partyTdsRt[key] = tdsRate;
    if (invoiceNo) itcMap[invoiceNo.toUpperCase()] = { itc_claimed: itcClaimed, has_stock: hasStock };
  }

  const b2bCountRaw = records.length;

  const b2baRecs = [];
  const b2baOrigNos = new Set();
  const b2baWs = wb.worksheets.find((s) => s.name.toUpperCase() === 'B2BA');
  if (b2baWs) {
    const detected = detectColumns(b2baWs);
    const b2baColumnMap = detected.columnMap;
    if ('gstin' in b2baColumnMap) {
      for (let rowIdx = detected.dataStartRow; rowIdx <= b2baWs.rowCount; rowIdx++) {
        const wsRow = b2baWs.getRow(rowIdx);
        const row = [];
        for (let c = 1; c <= b2baWs.columnCount; c++) row.push(wsRow.getCell(c).value);
        const gstinCol = b2baColumnMap.gstin !== undefined ? b2baColumnMap.gstin : 0;
        if (gstinCol >= row.length || !row[gstinCol]) continue;
        let rec;
        try { rec = parseB2bRow(row, rowIdx, b2baColumnMap); } catch { continue; }
        if (!rec) continue;
        rec.is_amendment = true;
        rec.sheet_type = 'B2BA';
        const key = (rec.trade_name || '').toUpperCase();
        rec.purchase_ledger = partyLedger[key] || '';
        rec.tds_ledger = partyTdsLdr[key] || '';
        rec.tds_rate = partyTdsRt[key] || 0.0;
        rec.tds_amount = '';
        if (rec.party_name === undefined) rec.party_name = rec.trade_name || '';
        if (rec.party_mailing_name === undefined) rec.party_mailing_name = rec.trade_name || '';
        for (const f of ['party_address1', 'party_address2', 'party_pincode', 'voucher_date', 'voucher_no', 'narration']) {
          if (rec[f] === undefined) rec[f] = '';
        }
        const supplierInv = rec.supplier_invoice_no || rec.invoice_no || '';
        if (rec.supplier_invoice_no === undefined) rec.supplier_invoice_no = supplierInv;
        if (rec.supplier_invoice_date === undefined) rec.supplier_invoice_date = rec.invoice_date || '';
        b2baRecs.push(rec);
        const origNo = (rec.orig_invoice_no || '').trim().toUpperCase();
        const invNo = (rec.invoice_no || '').trim().toUpperCase();
        b2baOrigNos.add(origNo || invNo);
      }
    }
  }

  let skippedB2b = 0;
  let finalRecords = records;
  if (b2baOrigNos.size) {
    const filtered = [];
    for (const r of records) {
      if (r.sheet_type === 'B2B' && b2baOrigNos.has((r.invoice_no || '').trim().toUpperCase())) skippedB2b += 1;
      else filtered.push(r);
    }
    finalRecords = filtered;
  }
  finalRecords.push(...b2baRecs);
  const b2baCount = b2baRecs.length;
  const b2bCount = b2bCountRaw - skippedB2b;

  let cdnrCount = 0;
  const cdnrWs = wb.worksheets.find((s) => ['B2B CDNR', 'B2B-CDNR', 'CDNR'].includes(s.name.toUpperCase()));
  if (cdnrWs) {
    const detected = detectColumns(cdnrWs);
    const cdnrColumnMap = detected.columnMap;
    if ('gstin' in cdnrColumnMap) {
      for (let rowIdx = detected.dataStartRow; rowIdx <= cdnrWs.rowCount; rowIdx++) {
        const wsRow = cdnrWs.getRow(rowIdx);
        const row = [];
        for (let c = 1; c <= cdnrWs.columnCount; c++) row.push(wsRow.getCell(c).value);
        const gstinCol = cdnrColumnMap.gstin !== undefined ? cdnrColumnMap.gstin : 0;
        if (gstinCol >= row.length || !row[gstinCol]) continue;
        let rec;
        try { rec = parseB2bRow(row, rowIdx, cdnrColumnMap); } catch { continue; }
        if (!rec) continue;
        rec.is_amendment = false;
        rec.sheet_type = 'CDNR';
        const key = (rec.trade_name || '').toUpperCase();
        rec.purchase_ledger = partyLedger[key] || '';
        rec.tds_ledger = partyTdsLdr[key] || '';
        rec.tds_rate = partyTdsRt[key] || 0.0;
        rec.tds_amount = '';
        if (rec.party_name === undefined) rec.party_name = rec.trade_name || '';
        if (rec.party_mailing_name === undefined) rec.party_mailing_name = rec.trade_name || '';
        for (const f of ['party_address1', 'party_address2', 'party_pincode', 'voucher_date', 'voucher_no', 'narration']) {
          if (rec[f] === undefined) rec[f] = '';
        }
        const supplierInv = rec.supplier_invoice_no || rec.invoice_no || '';
        if (rec.supplier_invoice_no === undefined) rec.supplier_invoice_no = supplierInv;
        if (rec.supplier_invoice_date === undefined) rec.supplier_invoice_date = rec.invoice_date || '';
        finalRecords.push(rec);
        cdnrCount += 1;
      }
    }
  }

  if (!finalRecords.length) {
    return {
      success: false, records: [],
      errors: ['No valid records found in the ITC B2B sheet. Make sure the sheet has GSTIN values in column A.'],
      warnings: [],
    };
  }

  const warnings = [];
  if (skippedB2b) {
    warnings.push(`B2BA: ${b2baCount} amendment(s) found — ${skippedB2b} B2B invoice(s) superseded and skipped.`);
  }

  return {
    success: true,
    records: finalRecords,
    party_ledger_map: partyLedger,
    party_tds_ledger_map: partyTdsLdr,
    party_tds_rate_map: partyTdsRt,
    itc_map: itcMap,
    errors: [],
    warnings,
    b2b_count: b2bCount,
    b2ba_count: b2baCount,
    cdnr_count: cdnrCount,
  };
}

module.exports = {
  ALLOWED_TAX_RATES, TAX_RATE_TOLERANCE, HEADER_PATTERNS, GSTIN_STATE_MAP,
  stateFromGstin, normalizeStateName, normalizeDateStr, tallyDate,
  detectColumns, parseB2bRow, parseReadme, safeGet, safeFloat, safeStr,
  parseGstr2bExcel, nearestAllowedTaxRate, validateTaxConfiguration,
  getGstLedger, getPurchaseLedger, resolvePurchaseLedger, asFloat, coalesceText,
  isIneligible, normalizeInvoiceNo, consolidateInvoiceRecords,
  addCommonLedgerFlags, companyRegistrationBlock, buildPurchaseVoucherXml,
  buildJournalVoucherXml, buildRcmJournalVoucherXml, buildVoucherXml, generateGstr2bXml,
  buildManualPurchaseVoucherXml, generateItcTemplateWorkbook, parseItcTemplateWorkbook,
  panFromGstin,
};
