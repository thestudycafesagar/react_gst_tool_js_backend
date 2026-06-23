/** GSTR-2B bridge endpoints (/gstr2b/*). 1:1 port of the corresponding
 * routes in Backend_Tally/main.py. */
const tallyBridge = require('../services/tallyBridge');
const gstr2bBridge = require('../services/gstr2bBridge');
const { isExcelFilename } = require('./bankController');
const { pushErrorHint } = require('./tallyController');
const { pushResponse } = require('../utils/responses');

async function processExcel(req, res) {
  const file = req.file;
  if (!file || !isExcelFilename(file.originalname)) {
    return res.status(400).json({ detail: 'Only Excel files (.xlsx, .xlsm, .xls) are supported.' });
  }
  const content = file.buffer;
  if (!content || !content.length) return res.status(400).json({ detail: 'Uploaded file is empty.' });

  const result = await gstr2bBridge.parseGstr2bExcel(content);
  res.json({
    success: result.success, records: result.records || [], errors: result.errors || [],
    warnings: result.warnings || [], company_gstin: result.company_gstin || '',
    company_name: result.company_name || '', trade_name: result.trade_name || '',
    financial_year: result.financial_year || '', tax_period: result.tax_period || '',
  });
}

/** Rebuilds the ITC template (ITC B2B + B2BA + B2B CDNR tabs, pre-filled
 * from the originally-uploaded GSTR-2B export) for the user to fill in
 * party/GST/TDS ledger mappings and ITC eligibility, then upload back. */
async function downloadTemplate(req, res) {
  const file = req.file;
  const content = file ? file.buffer : null;
  if (!content || !content.length) return res.status(400).json({ detail: 'Uploaded file is empty.' });

  let workbookBytes;
  try {
    workbookBytes = await gstr2bBridge.generateItcTemplateWorkbook(content);
  } catch (exc) {
    return res.status(422).json({ detail: `Failed to generate ITC template: ${exc.message}` });
  }
  res.set({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': 'attachment; filename=ITC_Template.xlsx',
  });
  res.send(Buffer.from(workbookBytes));
}

/** Re-parses a filled-in ITC template and REPLACES the working record set
 * (same full-replace behaviour as the desktop tool). */
async function uploadTemplate(req, res) {
  const file = req.file;
  const content = file ? file.buffer : null;
  if (!content || !content.length) return res.status(400).json({ detail: 'Uploaded file is empty.' });

  const result = await gstr2bBridge.parseItcTemplateWorkbook(content);
  res.json({
    success: result.success || false,
    records: result.records || [],
    partyLedgerMap: result.party_ledger_map || {},
    partyTdsLedgerMap: result.party_tds_ledger_map || {},
    partyTdsRateMap: result.party_tds_rate_map || {},
    itcMap: result.itc_map || {},
    errors: result.errors || [],
    warnings: result.warnings || [],
    b2bCount: result.b2b_count || 0,
    b2baCount: result.b2ba_count || 0,
    cdnrCount: result.cdnr_count || 0,
  });
}

function validateTax(req, res) {
  const { records = [] } = req.body;
  const [validRecords, invalidIssues] = gstr2bBridge.validateTaxConfiguration(records);
  res.json({ validRecords, invalidIssues });
}

function buildGstr2bXml(body) {
  const {
    company = '', records = [], companyGstin = '', companyRegistrationName = '',
    companyRegistrationState = '', purchaseLedger = 'Purchase Account',
    narrationTemplate = 'Being purchase from {party} vide Inv {inv} dt {date}',
    roundOffLedger = '', rcmLedgerMap = {}, gstLedgerRateMap = {}, partyLedgerMap = {},
    useInvoiceNoAsVoucherNo = false,
  } = body;
  const partyLedgerMapUpper = {};
  for (const [k, v] of Object.entries(partyLedgerMap)) partyLedgerMapUpper[k.toUpperCase()] = v;
  return gstr2bBridge.generateGstr2bXml(records, {
    companyName: company, companyGstin, companyRegistrationName, companyRegistrationState,
    purchaseLedger, narrationTemplate, roundOffLedger, rcmLedgerMap, gstLedgerRateMap,
    partyLedgerMap: partyLedgerMapUpper, useInvoiceNoAsVoucherNo,
  });
}

function generateXml(req, res) {
  const { records = [] } = req.body;
  if (!records.length) return res.status(422).json({ detail: 'No GSTR-2B records to generate.' });
  let xml;
  try {
    xml = buildGstr2bXml(req.body);
  } catch (exc) {
    return res.status(422).json({ detail: `XML generation failed: ${exc.message}` });
  }
  res.json({ xml });
}

async function pushVouchers(req, res) {
  const {
    host = 'localhost', port = '9000', company = '', records = [], purchaseLedger = '',
    roundOffLedger = '', rcmLedgerMap = {}, gstLedgerRateMap = {}, partyLedgerMap = {},
  } = req.body;
  if (!records.length) return res.json(pushResponse({ error: 'No GSTR-2B records to push.' }));

  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json(pushResponse({ error: exc.message }));
  }

  const existingLedgerNames = await tallyBridge.fetchAllLedgerNames(url, company);
  const suspenseLedger = tallyBridge.pickSuspenseLedger(existingLedgerNames);

  const partyLedgerMapUpper = {};
  for (const [k, v] of Object.entries(partyLedgerMap)) partyLedgerMapUpper[k.toUpperCase()] = v;

  const neededNames = new Set();
  for (const rec of records) {
    for (const field of ['trade_name', 'purchase_ledger', 'tds_ledger', 'cgst_ledger', 'sgst_ledger', 'igst_ledger']) {
      const val = String(rec[field] || '').trim();
      if (val) neededNames.add(val);
    }
  }
  for (const v of Object.values(partyLedgerMapUpper)) if (v) neededNames.add(v);
  for (const mapping of Object.values(gstLedgerRateMap || {})) {
    for (const v of Object.values(mapping || {})) if (v) neededNames.add(v);
  }
  if (purchaseLedger) neededNames.add(purchaseLedger);
  if (roundOffLedger) neededNames.add(roundOffLedger);
  for (const v of Object.values(rcmLedgerMap || {})) if (v) neededNames.add(v);

  const newLedgerNames = [...neededNames].filter((n) => !existingLedgerNames.has(n.toUpperCase()));
  if (newLedgerNames.length) {
    try {
      const ledgerDefs = newLedgerNames.map((n) => ({ Name: n, Parent: suspenseLedger }));
      const ledgerXml = tallyBridge.generateLedgerXml(ledgerDefs, company);
      await tallyBridge.postTallyXml(url, ledgerXml, 60000);
    } catch {
      // best-effort; Tally will reject affected vouchers below with a clear line error
    }
  }

  let xml;
  try {
    xml = buildGstr2bXml(req.body);
  } catch (exc) {
    return res.json(pushResponse({ error: `XML generation failed: ${exc.message}` }));
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

/** Pushes a single manually-entered voucher directly to Tally — the web
 * equivalent of the desktop tool's per-record manual voucher dialog, used
 * outside the bulk consolidation/generation pipeline for Stock-Item
 * records (item mode) and tax-slab-mismatch records the user chooses to
 * keep instead of dropping (accounting mode). */
async function pushManualVoucher(req, res) {
  const {
    host = 'localhost', port = '9000', company = '', companyGstin = '', companyRegistrationName = '',
    companyRegistrationState = '', useExistingSeries = true, invoiceNo, voucherDate, partyName,
    gstin = '', narration = '', purchaseLedger = '', taxEntries = [], itemEntries = [], ledgerEntries = [],
  } = req.body;
  const isItemMode = itemEntries.length > 0;
  if (!isItemMode && !ledgerEntries.length) return res.json(pushResponse({ error: 'Add at least one item or ledger row.' }));
  if (isItemMode && !String(purchaseLedger).trim()) return res.json(pushResponse({ error: 'Purchase ledger is required.' }));

  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json(pushResponse({ error: exc.message }));
  }

  const resolvedItemEntries = isItemMode ? itemEntries.map((it) => ({ ...it, ledger: purchaseLedger })) : [];
  const resolvedLedgerEntries = isItemMode ? [] : ledgerEntries;

  const existingLedgerNames = await tallyBridge.fetchAllLedgerNames(url, company);
  const suspenseLedger = tallyBridge.pickSuspenseLedger(existingLedgerNames);
  const neededNames = new Set([partyName, purchaseLedger]);
  for (const t of taxEntries) if (t.ledger) neededNames.add(t.ledger);
  for (const e of resolvedLedgerEntries) if (e.ledger) neededNames.add(e.ledger);
  neededNames.delete('');
  neededNames.delete(undefined);
  const newLedgerNames = [...neededNames].filter((n) => n && !existingLedgerNames.has(String(n).toUpperCase()));
  if (newLedgerNames.length) {
    try {
      const ledgerXml = tallyBridge.generateLedgerXml(newLedgerNames.map((n) => ({ Name: n, Parent: suspenseLedger })), company);
      await tallyBridge.postTallyXml(url, ledgerXml, 60000);
    } catch {
      // best-effort
    }
  }

  let xml;
  try {
    xml = gstr2bBridge.buildManualPurchaseVoucherXml(invoiceNo, voucherDate, partyName, gstin, taxEntries, {
      itemEntries: resolvedItemEntries, ledgerEntries: resolvedLedgerEntries, narration,
      companyName: company, companyGstin, companyRegistrationName, companyRegistrationState, useExistingSeries,
    });
  } catch (exc) {
    return res.json(pushResponse({ error: `XML generation failed: ${exc.message}` }));
  }

  let responseText;
  try {
    responseText = await tallyBridge.postTallyXml(url, xml, 60000);
  } catch (exc) {
    const detail = `Could not reach Tally: ${exc.message}`;
    return res.json(pushResponse({ error: detail + pushErrorHint(detail) }));
  }

  const details = tallyBridge.parseTallyResponseDetails(responseText);
  if (!details.success) {
    const detail = details.line_errors.length ? details.line_errors[0] : (details.error || 'Tally rejected the voucher.');
    return res.json(pushResponse({
      errors: details.errors, line_errors: details.line_errors,
      error: detail + pushErrorHint(detail),
    }));
  }

  res.json(pushResponse({ success: true, created: details.created, altered: details.altered, ignored: details.ignored }));
}

module.exports = {
  processExcel, downloadTemplate, uploadTemplate, validateTax, generateXml, pushVouchers, pushManualVoucher,
};
