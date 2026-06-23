/** Tally bridge endpoints (/tally/*). 1:1 port of the corresponding routes
 * in Backend_Tally/main.py. Tally's XML port sends no CORS headers, so the
 * browser can't call it directly — the React app calls these endpoints
 * instead, and this service (server-to-server, no CORS involved) talks to
 * Tally on its behalf. */
const tallyBridge = require('../services/tallyBridge');
const { pushResponse } = require('../utils/responses');

const PUSH_BATCH_SIZE = 40;

function pushErrorHint(detail) {
  const lowered = String(detail).toLowerCase();
  if (lowered.includes('out of range')) {
    return ' Tip: this usually means the target company\'s Books Beginning date is '
      + 'after the statement date. In Tally, open company settings and set '
      + 'Books Beginning From on/before the statement start date.';
  }
  if (lowered.includes('timed out') || lowered.includes('timeout')) {
    return ' Tip: Tally may still be processing a large push. Wait a minute, check '
      + 'for the created vouchers in Tally, then retry only the remaining rows.';
  }
  return '';
}

async function testConnection(req, res) {
  const { host = 'localhost', port = '9000' } = req.body;
  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json({ connected: false, error: exc.message });
  }
  const result = await tallyBridge.checkTallyConnection(url);
  res.json({ connected: result.connected, error: result.error ?? null });
}

async function companies(req, res) {
  const { host = 'localhost', port = '9000' } = req.body;
  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json({ success: false, companies: [], error: exc.message });
  }
  const result = await tallyBridge.fetchTallyCompanies(url);
  res.json({ success: result.success, companies: result.companies || [], error: result.error ?? null });
}

async function bankLedgers(req, res) {
  const { host = 'localhost', port = '9000', company = '' } = req.body;
  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json({ success: false, ledgers: [], error: exc.message });
  }
  const result = await tallyBridge.fetchBankLedgers(url, company);
  res.json({ success: result.success, ledgers: result.ledgers || [], error: result.error ?? null });
}

async function allLedgers(req, res) {
  const { host = 'localhost', port = '9000', company = '' } = req.body;
  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json({ success: false, ledgers: [], error: exc.message });
  }
  const names = await tallyBridge.fetchAllLedgers(url, company);
  if (names.length) return res.json({ success: true, ledgers: names, error: null });
  res.json({ success: false, error: 'No ledgers found — is Tally running?', ledgers: [] });
}

/** Every stock item name in Tally — powers the Item Name autocomplete in
 * the manual voucher entry dialog. */
async function stockItems(req, res) {
  const { host = 'localhost', port = '9000', company = '' } = req.body;
  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json({ success: false, items: [], error: exc.message });
  }
  const items = await tallyBridge.fetchAllStockItemNames(url, company);
  if (items.length) return res.json({ success: true, items, error: null });
  res.json({ success: false, error: 'No stock items found — is Tally running?', items: [] });
}

/** Creates a single ledger in Tally from the missing-ledger mapping dialog's
 * 'Create Ledger' form. */
async function createLedger(req, res) {
  const {
    host = 'localhost', port = '9000', company = '', name, parent = 'Suspense A/c',
    gstApplicable = '', gstRegistrationType = '', gstin = '', pan = '', state = '',
    address1 = '', address2 = '', pincode = '', billwise = '',
  } = req.body;
  const trimmedName = String(name || '').trim();
  if (!trimmedName) return res.json({ success: false, message: 'Ledger name is required.' });

  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json({ success: false, message: exc.message });
  }

  const gstinUpper = String(gstin).trim().toUpperCase();
  const ledgerDef = {
    Name: trimmedName,
    Parent: String(parent).trim() || 'Suspense A/c',
    MailingName: trimmedName,
    GSTIN: gstinUpper,
    PAN: String(pan).trim().toUpperCase() || tallyBridge.panFromGstin(gstinUpper),
    GSTApplicable: String(gstApplicable).trim(),
    GSTRegistrationType: String(gstRegistrationType).trim(),
    StateOfSupply: String(state).trim(),
    Address1: String(address1).trim(),
    Address2: String(address2).trim(),
    Pincode: String(pincode).trim(),
    Country: 'India',
    Billwise: String(billwise).trim(),
  };

  let responseText;
  try {
    const xml = tallyBridge.generateLedgerXml([ledgerDef], company);
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

async function generateXml(req, res) {
  const { host = 'localhost', port = '9000', company = '', bankLedger, dateMode = 'excel', customDate = '', rows } = req.body;
  let contraLedgerNames = [];
  try {
    const url = tallyBridge.buildTallyUrl(host, port);
    const bankResult = await tallyBridge.fetchBankLedgers(url, company);
    if (bankResult.success) contraLedgerNames = (bankResult.ledgers || []).map((l) => l.name);
  } catch {
    // best-effort; without it, bank-to-bank transfers just post as Payment/Receipt
  }

  try {
    const xml = tallyBridge.generateBankVoucherXml(rows, company, bankLedger, {
      dateMode, customTallyDate: customDate, contraLedgerNames,
    });
    res.json({ xml });
  } catch (exc) {
    res.status(422).json({ detail: exc.message });
  }
}

async function pushVouchers(req, res) {
  const {
    host = 'localhost', port = '9000', company = '', bankLedger = '', dateMode = 'excel',
    customDate = '', rows = [],
  } = req.body;
  const trimmedBankLedger = String(bankLedger).trim();
  if (!trimmedBankLedger) return res.json(pushResponse({ error: 'Please select a Bank Ledger before pushing.' }));
  if (!rows.length) return res.json(pushResponse({ error: 'No bank statement rows to push.' }));

  let url;
  try {
    url = tallyBridge.buildTallyUrl(host, port);
  } catch (exc) {
    return res.json(pushResponse({ error: exc.message }));
  }

  const bankResult = await tallyBridge.fetchBankLedgers(url, company);
  const contraLedgerNames = bankResult.success ? (bankResult.ledgers || []).map((l) => l.name) : [];

  const existingLedgerNames = await tallyBridge.fetchAllLedgerNames(url, company);
  const suspenseLedger = tallyBridge.pickSuspenseLedger(existingLedgerNames);

  const newLedgerNames = new Set();
  for (const row of rows) {
    const ledgerTrim = String(row.ledger || '').trim();
    if (ledgerTrim && !existingLedgerNames.has(ledgerTrim.toUpperCase()) && ledgerTrim.toUpperCase() !== trimmedBankLedger.toUpperCase()) {
      newLedgerNames.add(ledgerTrim);
    }
  }
  if (newLedgerNames.size) {
    try {
      const ledgerDefs = [...newLedgerNames].map((n) => ({ Name: n, Parent: suspenseLedger }));
      const ledgerXml = tallyBridge.generateLedgerXml(ledgerDefs, company);
      await tallyBridge.postTallyXml(url, ledgerXml, 60000);
    } catch {
      // best-effort; if it still doesn't exist, Tally rejects affected rows below with a clear line error
    }
  }

  let paymentCursor = null, receiptCursor = null, contraCursor = null;
  const pmtResult = await tallyBridge.fetchNextVoucherNumber(url, company, 'Payment');
  if (pmtResult.success) paymentCursor = pmtResult.next_number;
  const rctResult = await tallyBridge.fetchNextVoucherNumber(url, company, 'Receipt');
  if (rctResult.success) receiptCursor = rctResult.next_number;
  const ctrResult = await tallyBridge.fetchNextVoucherNumber(url, company, 'Contra');
  if (ctrResult.success) contraCursor = ctrResult.next_number;

  const batches = [];
  for (let i = 0; i < rows.length; i += PUSH_BATCH_SIZE) batches.push(rows.slice(i, i + PUSH_BATCH_SIZE));
  let createdTotal = 0, alteredTotal = 0, ignoredTotal = 0;

  for (const batchRows of batches) {
    let xml;
    try {
      xml = tallyBridge.generateBankVoucherXml(batchRows, company, trimmedBankLedger, {
        dateMode, customTallyDate: customDate, suspenseLedger,
        paymentStartVno: paymentCursor, receiptStartVno: receiptCursor,
        contraLedgerNames, contraStartVno: contraCursor,
      });
    } catch (exc) {
      return res.status(422).json({ detail: exc.message });
    }

    let responseText;
    try {
      responseText = await tallyBridge.postTallyXml(url, xml, 300000);
    } catch (exc) {
      const detail = `Could not reach Tally: ${exc.message}`;
      return res.json(pushResponse({
        created: createdTotal, altered: alteredTotal, ignored: ignoredTotal,
        error: detail + pushErrorHint(detail),
      }));
    }

    const details = tallyBridge.parseTallyResponseDetails(responseText);
    if (!details.success) {
      const detail = details.line_errors.length ? details.line_errors[0] : (details.error || 'Tally rejected the vouchers.');
      return res.json(pushResponse({
        created: createdTotal, altered: alteredTotal, ignored: ignoredTotal,
        errors: details.errors, line_errors: details.line_errors,
        error: detail + pushErrorHint(detail),
      }));
    }

    createdTotal += details.created;
    alteredTotal += details.altered;
    ignoredTotal += details.ignored;

    const { payment, receipt, contra } = tallyBridge.countVoucherEntries(batchRows, contraLedgerNames);
    if (paymentCursor !== null) paymentCursor += payment;
    if (receiptCursor !== null) receiptCursor += receipt;
    if (contraCursor !== null) contraCursor += contra;
  }

  res.json(pushResponse({ success: true, created: createdTotal, altered: alteredTotal, ignored: ignoredTotal }));
}

module.exports = {
  testConnection, companies, bankLedgers, allLedgers, stockItems, createLedger,
  generateXml, pushVouchers, pushErrorHint, PUSH_BATCH_SIZE,
};
