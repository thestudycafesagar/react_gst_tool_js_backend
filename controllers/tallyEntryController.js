/** Tally Entry (generic Sales/Purchase/Note/Journal/Payment/Receipt/Export
 * voucher entry) endpoints (/tally-entry/*). 1:1 port of the corresponding
 * routes in Backend_Tally/main.py. */
const tallyBridge = require('../services/tallyBridge');
const templates = require('../services/tallyEntry/templates');
const common = require('../services/tallyEntry/common');
const sales = require('../services/tallyEntry/sales');
const purchase = require('../services/tallyEntry/purchase');
const notes = require('../services/tallyEntry/notes');
const journal = require('../services/tallyEntry/journal');
const paymentReceipt = require('../services/tallyEntry/paymentReceipt');
const salesExport = require('../services/tallyEntry/salesExport');
const ledgerCollection = require('../services/tallyEntry/ledgerCollection');
const stockItem = require('../services/tallyEntry/stockItem');
const { isExcelFilename } = require('./bankController');
const { pushErrorHint } = require('./tallyController');
const { pushResponse } = require('../utils/responses');

const TALLY_ENTRY_MODES = new Set([
  'accounting', 'item', 'purchase_accounting', 'purchase_item',
  'credit_note_accounting', 'credit_note_item', 'debit_note_accounting', 'debit_note_item',
  'sales_journal', 'purchase_journal', 'payment', 'receipt', 'sales_export',
]);
const TALLY_ENTRY_DEFAULT_VCH_TYPE = {
  accounting: 'Sales', item: 'Sales', purchase_accounting: 'Purchase', purchase_item: 'Purchase',
  credit_note_accounting: 'Credit Note', credit_note_item: 'Credit Note',
  debit_note_accounting: 'Debit Note', debit_note_item: 'Debit Note',
  sales_journal: 'Journal', purchase_journal: 'Journal',
  payment: 'Payment', receipt: 'Receipt', sales_export: 'Sales Export',
};
// Modes whose builder returns [xml, voucherCount] instead of a bare xml string.
const TALLY_ENTRY_COUNTED_MODES = new Set([
  'credit_note_accounting', 'credit_note_item', 'debit_note_accounting', 'debit_note_item',
  'sales_journal', 'purchase_journal', 'payment', 'receipt', 'sales_export',
]);

class HttpError extends Error {
  constructor(status, detail) {
    super(detail);
    this.status = status;
    this.detail = detail;
  }
}

/** Creates a single stock item in Tally — the standalone Create Stock
 * Item tab's manual-entry form. */
async function createStockItem(req, res) {
  const {
    host = 'localhost', port = '9000', company = '', name, parent = 'Primary', unit = 'Nos',
    hsnCode = '', gstRate = '', description = '', typeOfSupply = 'Goods',
  } = req.body;
  const trimmedName = String(name || '').trim();
  if (!trimmedName) return res.json({ success: false, message: 'Stock item name is required.' });

  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json({ success: false, message: exc.message });
  }

  const itemDef = {
    Name: trimmedName, Parent: String(parent).trim() || 'Primary', Unit: String(unit).trim() || 'Nos',
    HSNCode: String(hsnCode).trim(), GSTRate: String(gstRate).trim(), Description: String(description).trim(),
    TypeOfSupply: String(typeOfSupply).trim() || 'Goods',
  };

  let responseText;
  try {
    const xml = stockItem.generateStockitemXml([itemDef], company);
    responseText = await tallyBridge.postTallyXml(url, xml, 60000);
  } catch (exc) {
    return res.json({ success: false, message: `Could not reach Tally: ${exc.message}` });
  }

  const parsed = tallyBridge.parseTallyResponseDetails(responseText);
  if (parsed.success || parsed.created > 0 || parsed.altered > 0) {
    return res.json({ success: true, message: `'${trimmedName}' created successfully.` });
  }
  const detail = (parsed.line_errors && parsed.line_errors.join('; ')) || parsed.error || 'Tally returned an error.';
  res.json({ success: false, message: detail });
}

function tallyEntryStats(rows) {
  const sumAbs = (key) => rows.reduce((s, r) => s + Math.abs(common.rowFloat(r, key, 0.0)), 0);
  const totalTaxable = sumAbs('TaxableValue');
  const totalCgst = sumAbs('CGST Amount');
  const totalSgst = sumAbs('SGST Amount');
  const totalIgst = sumAbs('IGST Amount');
  return {
    totalInvoices: rows.length,
    totalTaxable: totalTaxable / 100000,
    totalIGST: totalIgst / 100000,
    totalCGST: totalCgst / 100000,
    totalSGST: totalSgst / 100000,
  };
}

async function processExcel(req, res) {
  const file = req.file;
  if (!file || !isExcelFilename(file.originalname)) {
    return res.status(400).json({ detail: 'Only Excel files (.xlsx, .xlsm, .xls) are supported.' });
  }
  const content = file.buffer;
  if (!content || !content.length) return res.status(400).json({ detail: 'Uploaded file is empty.' });

  let rows;
  try {
    rows = await templates.readExcelRows(content);
  } catch (exc) {
    return res.status(422).json({ detail: `Could not read this Excel file: ${exc.message}` });
  }
  if (!rows.length) return res.json({ success: false, rows: [], errors: ['No data rows found below the header row.'], stats: {} });
  res.json({ success: true, rows, errors: [], stats: tallyEntryStats(rows) });
}

async function downloadTemplate(req, res) {
  const mode = req.query.mode;
  if (!TALLY_ENTRY_MODES.has(mode)) return res.status(400).json({ detail: `Unknown mode '${mode}'.` });
  let workbookBytes;
  try {
    workbookBytes = await templates.generateTemplateWorkbook(mode);
  } catch (exc) {
    return res.status(422).json({ detail: `Failed to generate template: ${exc.message}` });
  }
  res.set({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': `attachment; filename=TallyEntry_${mode}_Template.xlsx`,
  });
  res.send(Buffer.from(workbookBytes));
}

function buildXml(body, companyGstRegistrations) {
  const {
    company = '', mode, rows = [], voucherType = '', dateMode = 'excel', customDate = '',
    roundOffLedger = '', useExistingSeries = true, companyGstin = '', startVoucherNumber = '',
    entryMode = 'accounting', bankLedger = '', contraLedgerNames = [],
  } = body;
  if (!TALLY_ENTRY_MODES.has(mode)) throw new HttpError(400, `Unknown mode '${mode}'.`);
  const resolvedVoucherType = String(voucherType).trim() || TALLY_ENTRY_DEFAULT_VCH_TYPE[mode];
  const startVno = String(startVoucherNumber).trim() || null;
  const commonOpts = {
    dateMode, customTallyDate: customDate, startVoucherNumber: startVno,
    voucherType: resolvedVoucherType, roundOffLedger,
  };

  if (mode === 'accounting') {
    return sales.generateAccountingXml(rows, company, { ...commonOpts, companyGstRegistrations });
  }
  if (mode === 'item') {
    return sales.generateItemXml(rows, company, { ...commonOpts, companyGstRegistrations });
  }
  if (mode === 'purchase_accounting') {
    return purchase.generatePurchaseAccountingXml(rows, company, { ...commonOpts, companyGstRegistrations, useExistingSeries });
  }
  if (mode === 'purchase_item') {
    return purchase.generatePurchaseItemXml(rows, company, {
      ...commonOpts, companyGstRegistrations, companyGstin, useExistingSeries,
    });
  }
  if (['credit_note_accounting', 'credit_note_item', 'debit_note_accounting', 'debit_note_item'].includes(mode)) {
    const noteVoucherType = mode.startsWith('debit_note') ? 'Debit Note' : 'Credit Note';
    const noteEntryMode = mode.endsWith('_item') ? 'item' : 'accounting';
    return notes.generateNoteXml(rows, company, {
      dateMode, customTallyDate: customDate, voucherType: noteVoucherType,
      companyGstRegistrations, entryMode: noteEntryMode, roundOffLedger,
    });
  }
  if (mode === 'sales_journal' || mode === 'purchase_journal') {
    const journalType = mode === 'sales_journal' ? 'sale' : 'purchase';
    return journal.generateJournalXml(rows, company, {
      dateMode, customTallyDate: customDate, journalType, companyGstRegistrations, roundOffLedger,
    });
  }
  if (mode === 'payment' || mode === 'receipt') {
    if (!String(bankLedger).trim()) throw new HttpError(422, 'Bank/Cash ledger is required for Payment/Receipt vouchers.');
    return paymentReceipt.generateSingleVoucherXml(rows, company, bankLedger, resolvedVoucherType, {
      dateMode, customTallyDate: customDate, startVno,
      contraLedgerNames: contraLedgerNames || [],
    });
  }
  return salesExport.generateSalesExportAccountingXml(rows, company, {
    dateMode, customTallyDate: customDate, startVoucherNumber: startVno,
    voucherType: resolvedVoucherType, roundOffLedger,
  });
}

/** buildXml returns a bare XML string for the original 4 modes, but
 * [xml, voucherCount] for every Stage-2 mode — unwrap here so callers
 * always get a plain XML string. */
function extractXml(body, companyGstRegistrations) {
  const result = buildXml(body, companyGstRegistrations);
  return TALLY_ENTRY_COUNTED_MODES.has(body.mode) ? result[0] : result;
}

async function generateXml(req, res) {
  const { host = 'localhost', port = '9000', company = '', rows = [], fetchCompanyGstRegistrations = true } = req.body;
  if (!rows.length) return res.status(422).json({ detail: 'No rows to generate.' });

  let companyGstRegistrations = [];
  if (fetchCompanyGstRegistrations) {
    try {
      const url = tallyBridge.buildTallyUrl(host, port);
      const regResult = await tallyBridge.fetchCompanyGstRegistrations(url, company);
      if (regResult.success) companyGstRegistrations = regResult.registrations || [];
    } catch {
      // best-effort; builders fall back to a synthetic registration name
    }
  }

  try {
    const xml = extractXml(req.body, companyGstRegistrations);
    res.json({ xml });
  } catch (exc) {
    if (exc instanceof HttpError) return res.status(exc.status).json({ detail: exc.detail });
    res.status(422).json({ detail: exc.message });
  }
}

async function pushVouchers(req, res) {
  const { host = 'localhost', port = '9000', company = '', rows = [], mode, fetchCompanyGstRegistrations = true } = req.body;
  if (!rows.length) return res.json(pushResponse({ error: 'No rows to push.' }));

  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json(pushResponse({ error: exc.message }));
  }

  let companyGstRegistrations = [];
  if (fetchCompanyGstRegistrations) {
    const regResult = await tallyBridge.fetchCompanyGstRegistrations(url, company);
    if (regResult.success) companyGstRegistrations = regResult.registrations || [];
  }

  const existingLedgerNames = await tallyBridge.fetchAllLedgerNames(url, company);
  const suspenseLedger = tallyBridge.pickSuspenseLedger(existingLedgerNames);
  let neededNames;
  try {
    neededNames = ledgerCollection.collectReferencedLedgers(rows, mode, suspenseLedger);
  } catch {
    neededNames = new Set();
  }
  const newLedgerNames = [...neededNames].filter((n) => !existingLedgerNames.has(n.toUpperCase()));
  if (newLedgerNames.length) {
    try {
      const ledgerXml = tallyBridge.generateLedgerXml(newLedgerNames.map((n) => ({ Name: n, Parent: suspenseLedger })), company);
      await tallyBridge.postTallyXml(url, ledgerXml, 60000);
    } catch {
      // best-effort; Tally will reject affected vouchers below with a clear line error
    }
  }

  let xml;
  try {
    xml = extractXml(req.body, companyGstRegistrations);
  } catch (exc) {
    return res.json(pushResponse({ error: exc instanceof HttpError ? exc.detail : exc.message }));
  }

  let responseText;
  try {
    responseText = await tallyBridge.postTallyXml(url, xml, 300000);
  } catch (exc) {
    const detail = `Could not reach Tally: ${exc.message}`;
    return res.json(pushResponse({ error: detail + pushErrorHint(detail) }));
  }

  const details = tallyBridge.parseTallyResponseDetails(responseText);
  if (!details.success) {
    const detail = details.line_errors.length ? details.line_errors[0] : (details.error || 'Tally rejected the vouchers.');
    return res.json(pushResponse({
      errors: details.errors, line_errors: details.line_errors,
      error: detail + pushErrorHint(detail),
    }));
  }
  res.json(pushResponse({ success: true, created: details.created, altered: details.altered, ignored: details.ignored }));
}

/** Pushes an already-generated XML file straight to Tally — the 'Browse
 * XML (Existing) -> Preview & Push directly' flow, for vouchers built
 * outside this tool. */
async function pushXml(req, res) {
  const { host = 'localhost', port = '9000', xml = '' } = req.body;
  if (!String(xml).trim()) return res.json(pushResponse({ error: 'XML content is empty.' }));
  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json(pushResponse({ error: exc.message }));
  }
  let responseText;
  try {
    responseText = await tallyBridge.postTallyXml(url, xml, 300000);
  } catch (exc) {
    const detail = `Could not reach Tally: ${exc.message}`;
    return res.json(pushResponse({ error: detail + pushErrorHint(detail) }));
  }
  const details = tallyBridge.parseTallyResponseDetails(responseText);
  if (!details.success) {
    const detail = details.line_errors.length ? details.line_errors[0] : (details.error || 'Tally rejected the vouchers.');
    return res.json(pushResponse({
      errors: details.errors, line_errors: details.line_errors,
      error: detail + pushErrorHint(detail),
    }));
  }
  res.json(pushResponse({ success: true, created: details.created, altered: details.altered, ignored: details.ignored }));
}

module.exports = {
  createStockItem, processExcel, downloadTemplate, generateXml, pushVouchers, pushXml,
  TALLY_ENTRY_MODES, TALLY_ENTRY_DEFAULT_VCH_TYPE, TALLY_ENTRY_COUNTED_MODES,
};
