/** Scans raw rows for every ledger name a mode's voucher builder will
 * reference, so the caller can auto-create missing ones under Suspense
 * before pushing — one function shared by every mode.
 * 1:1 port of Backend_Tally/tally_entry/ledger_collection.py. */
const { SUSPENSE_LEDGER, ledgerOrSuspense, resolvePartyLedger, rowText, rowFloat, pickTaxLedgerName } = require('./common');

const TAX_LEDGER_KEYS = {
  cgst: ['CGSTLedger', 'CGST Ledger', 'CentralTaxLedger', 'Central Tax Ledger', 'Central Tax'],
  sgst: ['SGSTLedger', 'SGST Ledger', 'StateTaxLedger', 'State Tax Ledger', 'State Tax', 'UTGSTLedger', 'UTGST Ledger'],
  igst: ['IGSTLedger', 'IGST Ledger', 'IntegratedTaxLedger', 'Integrated Tax Ledger', 'Integrated Tax'],
};

const NOTE_MODES = new Set(['credit_note_accounting', 'credit_note_item', 'debit_note_accounting', 'debit_note_item']);
const JOURNAL_MODES = new Set(['sales_journal', 'purchase_journal']);
const SINGLE_VOUCHER_MODES = new Set(['payment', 'receipt']);

/** Scans raw (pre-consolidation) rows for every ledger name the voucher
 * builder for `mode` will reference — party, sales/purchase, tax, TDS, and
 * (for item modes) per-item sales/purchase ledgers — so the caller can
 * auto-create whichever ones don't exist in Tally yet before pushing. */
function collectReferencedLedgers(rows, mode, fallbackLedger = SUSPENSE_LEDGER) {
  const names = new Set();

  if (SINGLE_VOUCHER_MODES.has(mode)) {
    for (const r of rows || []) {
      const contraRaw = (rowText(r, 'LEDGER') || rowText(r, 'Ledger')).trim() || 'Suspense A/c';
      names.add(contraRaw);
    }
    names.delete('');
    return names;
  }

  if (mode === 'sales_export') {
    for (const r of rows || []) {
      names.add(ledgerOrSuspense(resolvePartyLedger(r, false), fallbackLedger));
      names.add(ledgerOrSuspense(rowText(r, 'SalesLedger'), fallbackLedger));
    }
    names.delete('');
    return names;
  }

  if (NOTE_MODES.has(mode)) {
    const isItem = mode.endsWith('_item');
    for (const r of rows || []) {
      names.add(ledgerOrSuspense(rowText(r, 'PartyLedger') || rowText(r, 'Party Ledger'), fallbackLedger));
      if (isItem) {
        const purchaseLedgerRaw = rowText(r, 'PurchaseLedger') || rowText(r, 'Purchase Ledger');
        if (purchaseLedgerRaw) names.add(ledgerOrSuspense(purchaseLedgerRaw, fallbackLedger));
      } else {
        const particularRaw = rowText(r, 'Particular') || rowText(r, 'Particulars')
          || rowText(r, 'SalesLedger') || rowText(r, 'Sales Ledger')
          || rowText(r, 'PurchaseLedger') || rowText(r, 'Purchase Ledger')
          || rowText(r, 'ExpenseLedger');
        if (particularRaw) names.add(ledgerOrSuspense(particularRaw, fallbackLedger));
      }
      for (const [taxType, keys] of Object.entries(TAX_LEDGER_KEYS)) {
        const rate = rowFloat(r, `${taxType.toUpperCase()}Rate`, 0.0);
        const amt = rowFloat(r, `${taxType.toUpperCase()} Amount`, 0.0);
        const ledName = pickTaxLedgerName(r, keys, rate, taxType.toUpperCase(), amt);
        if (ledName && (ledName !== fallbackLedger || rate > 0 || amt > 0)) names.add(ledName);
      }
    }
    names.delete('');
    return names;
  }

  if (JOURNAL_MODES.has(mode)) {
    const isPurchase = mode === 'purchase_journal';
    for (const r of rows || []) {
      names.add(ledgerOrSuspense(rowText(r, 'PartyLedger') || rowText(r, 'Party Ledger'), fallbackLedger));
      const particularRaw = rowText(r, 'Particular') || rowText(r, 'Particulars')
        || rowText(r, 'SalesLedger') || rowText(r, 'Sales Ledger')
        || rowText(r, 'PurchaseLedger') || rowText(r, 'Purchase Ledger')
        || rowText(r, 'ExpenseLedger');
      if (particularRaw) names.add(ledgerOrSuspense(particularRaw, fallbackLedger));
      for (const [taxType, keys] of Object.entries(TAX_LEDGER_KEYS)) {
        const rate = rowFloat(r, `${taxType.toUpperCase()}Rate`, 0.0);
        const amt = rowFloat(r, `${taxType.toUpperCase()} Amount`, 0.0);
        const ledName = pickTaxLedgerName(r, keys, rate, taxType.toUpperCase(), amt);
        if (ledName && (ledName !== fallbackLedger || rate > 0 || amt > 0)) names.add(ledName);
      }
      if (isPurchase) {
        const tdsLedger = rowText(r, 'TDSLedger') || rowText(r, 'TDS Ledger') || rowText(r, 'Tds Ledger');
        if (tdsLedger) names.add(tdsLedger);
      }
    }
    names.delete('');
    return names;
  }

  const isPurchase = mode === 'purchase_accounting' || mode === 'purchase_item';
  const isItem = mode === 'item' || mode === 'purchase_item';

  for (const r of rows || []) {
    names.add(ledgerOrSuspense(resolvePartyLedger(r, isPurchase), fallbackLedger));

    if (isItem) {
      const items = r.items || [r];
      for (const it of items) {
        const ledKey = isPurchase ? 'PurchaseLedger' : 'SalesLedger';
        let ledVal = (it && typeof it === 'object' && ledKey in it) ? it[ledKey] : null;
        if (ledVal == null) ledVal = isPurchase ? rowText(r, 'PurchaseLedger') : rowText(r, 'SalesLedger');
        names.add(ledgerOrSuspense(String(ledVal || ''), fallbackLedger));
      }
    } else {
      let ledRaw;
      if (isPurchase) {
        ledRaw = rowText(r, 'PurchaseLedger') || rowText(r, 'PurchaseAccount')
          || rowText(r, 'Purchase Ledger') || rowText(r, 'ExpenseLedger') || rowText(r, 'SalesLedger');
      } else {
        ledRaw = rowText(r, 'SalesLedger');
      }
      names.add(ledgerOrSuspense(ledRaw, fallbackLedger));
    }

    for (const [taxType, keys] of Object.entries(TAX_LEDGER_KEYS)) {
      const rate = rowFloat(r, `${taxType.toUpperCase()}Rate`, 0.0);
      const amt = rowFloat(r, `${taxType.toUpperCase()} Amount`, 0.0);
      const ledName = pickTaxLedgerName(r, keys, rate, taxType.toUpperCase(), amt);
      if (ledName && ledName !== fallbackLedger) names.add(ledName);
      else if (rate > 0 || amt > 0) names.add(ledName);
    }

    const tdsLedger = rowText(r, 'TDSLedger') || rowText(r, 'TDS Ledger') || rowText(r, 'Tds Ledger');
    if (tdsLedger) names.add(tdsLedger);
  }

  names.delete('');
  return names;
}

module.exports = { collectReferencedLedgers, TAX_LEDGER_KEYS, NOTE_MODES, JOURNAL_MODES, SINGLE_VOUCHER_MODES };
