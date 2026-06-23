/**
 * Shared row/date/ledger/GST-context helpers used by every voucher builder
 * in this package — 1:1 port of Backend_Tally/tally_entry/common.py.
 */
const tallyBridge = require('../tallyBridge');

const { xmlEscape, stateNameFromGstin: _stateNameFromGstin, normalizeGstRegistrationType: _normalizeGstRegistrationType } = tallyBridge;

const SUSPENSE_LEDGER = 'Suspense A/c';

function fmtAmt(num) {
  return Number(num).toFixed(2);
}

function nameKey(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

const MONTH_ABBR = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

function pad2(n) { return String(n).padStart(2, '0'); }
function ymd(y, mo, d) { return `${y}${pad2(mo)}${pad2(d)}`; }

/** Resolves any plausible date representation (Date object, Excel serial
 * number, ISO-8601 string, or one of a dozen text formats) to Tally's
 * YYYYMMDD — 1:1 port of tally_date(), including the ISO-8601 "T" handling
 * needed because rows round-trip through the frontend as JSON between the
 * process-excel and generate-xml requests. */
function tallyDate(dt) {
  const today = ymd(new Date().getFullYear(), new Date().getMonth() + 1, new Date().getDate());
  if (dt == null || dt === '') return today;
  if (dt instanceof Date) return ymd(dt.getFullYear(), dt.getMonth() + 1, dt.getDate());
  if (typeof dt === 'number' && Number.isFinite(dt)) {
    if (dt > 1000) {
      const base = new Date(1899, 11, 30);
      base.setDate(base.getDate() + Math.floor(dt));
      return ymd(base.getFullYear(), base.getMonth() + 1, base.getDate());
    }
  }

  let text = String(dt).trim();
  if (!text) return today;
  if (text.endsWith('.0') && /^\d+$/.test(text.slice(0, -2))) text = text.slice(0, -2);
  if (/^\d{8}$/.test(text)) {
    const year = parseInt(text.slice(0, 4), 10);
    if (year >= 1900 && year <= 2100) return text;
    return `${text.slice(4, 8)}${text.slice(2, 4)}${text.slice(0, 2)}`;
  }

  if (text.includes('T')) {
    const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})T/);
    if (isoMatch) return `${isoMatch[1]}${isoMatch[2]}${isoMatch[3]}`;
  }

  const candidates = [text];
  if (text.includes(' ')) candidates.push(text.split(' ')[0]);

  for (const candidate of candidates) {
    let m;
    if ((m = candidate.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/))) return ymd(+m[3], +m[2], +m[1]);
    if ((m = candidate.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2})$/))) return ymd(2000 + +m[3], +m[2], +m[1]);
    if ((m = candidate.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) return ymd(+m[1], +m[2], +m[3]);
    if ((m = candidate.match(/^(\d{1,2})-([A-Za-z]+)-(\d{2,4})$/))) {
      const monIdx = MONTH_ABBR.findIndex((mo) => m[2].toLowerCase().startsWith(mo));
      if (monIdx !== -1) {
        let y = +m[3]; if (y < 100) y += 2000;
        return ymd(y, monIdx + 1, +m[1]);
      }
    }
  }
  return today;
}

/** Validates + reformats a manually-typed custom date to YYYYMMDD,
 * throwing the same message as the Python version on invalid input. */
function normalizeManualDateToTally(dateText) {
  const text = String(dateText || '').trim();
  if (!text) throw new Error('Custom date is empty.');
  const compact = text.replace(/\s+/g, '');
  if (/^\d{8}$/.test(compact)) {
    // Try YYYYMMDD then DDMMYYYY, matching strptime's try-in-order semantics.
    const y1 = parseInt(compact.slice(0, 4), 10);
    if (y1 >= 1 && parseInt(compact.slice(4, 6), 10) <= 12 && parseInt(compact.slice(6, 8), 10) <= 31) return compact;
    const d2 = parseInt(compact.slice(0, 2), 10), mo2 = parseInt(compact.slice(2, 4), 10), y2 = parseInt(compact.slice(4, 8), 10);
    if (mo2 >= 1 && mo2 <= 12 && d2 >= 1 && d2 <= 31) return ymd(y2, mo2, d2);
  }
  let m;
  if ((m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) return ymd(+m[3], +m[2], +m[1]);
  if ((m = text.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/))) return ymd(+m[3], +m[2], +m[1]);
  if ((m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) return ymd(+m[1], +m[2], +m[3]);
  if ((m = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/))) return ymd(+m[1], +m[2], +m[3]);
  if ((m = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/))) return ymd(+m[3], +m[2], +m[1]);
  throw new Error('Invalid custom date format. Use DD/MM/YYYY, DD-MM-YYYY, or YYYY-MM-DD.');
}

/** round_off = rounded_total - actual_total (signed). Sales: positive ->
 * extra collected -> credit. Purchase: positive -> extra paid -> debit. */
function appendRoundOffEntry(a, roundOffLedger, roundOff, isPurchaseMode) {
  if (!roundOffLedger || Math.abs(roundOff) < 0.005) return;
  const esc = xmlEscape(roundOffLedger);
  const isCredit = (roundOff > 0) !== isPurchaseMode;
  a('     <LEDGERENTRIES.LIST>');
  a(`      <LEDGERNAME>${esc}</LEDGERNAME>`);
  if (isCredit) {
    a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
    a(`      <AMOUNT>${fmtAmt(Math.abs(roundOff))}</AMOUNT>`);
  } else {
    a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
    a(`      <AMOUNT>-${fmtAmt(Math.abs(roundOff))}</AMOUNT>`);
  }
  a('     </LEDGERENTRIES.LIST>');
}

// ─── Row helpers (tolerant of header spacing/case differences) ────────────

function rowGet(row, key, defaultValue = null) {
  if (Object.prototype.hasOwnProperty.call(row, key)) {
    const value = row[key];
    return value == null ? defaultValue : value;
  }
  const target = String(key || '').replace(/\s+/g, '').toLowerCase();
  for (const [rawKey, rawValue] of Object.entries(row)) {
    if (String(rawKey || '').replace(/\s+/g, '').toLowerCase() === target) {
      return rawValue == null ? defaultValue : rawValue;
    }
  }
  return defaultValue;
}

/** Tries every plausible voucher-date column name across all Tally Entry
 * templates ('Date' for Sales/Purchase Item, 'Voucher Date' for Purchase
 * Accounting), so a row keeps its Excel date regardless of column header. */
function resolveExcelDate(row) {
  for (const key of ['Date', 'Voucher Date', 'VoucherDate', 'InvoiceDate', 'Invoice Date']) {
    const value = rowGet(row, key, '');
    if (value !== null && value !== '') return value;
  }
  return '';
}

function rowText(row, key, defaultValue = '') {
  const value = rowGet(row, key, defaultValue);
  return value == null ? defaultValue : String(value).trim();
}

function rowTextAny(row, keys, defaultValue = '') {
  for (const key of keys || []) {
    const value = rowText(row, key, '');
    if (value) return value;
  }
  return defaultValue;
}

function rowFloat(row, key, defaultValue = 0.0) {
  const value = rowGet(row, key, defaultValue);
  if (value === null || value === '') return Number(defaultValue);
  const n = parseFloat(value);
  return Number.isFinite(n) ? n : Number(defaultValue);
}

function rowVoucherNumber(row, defaultValue = '') {
  return rowText(row, 'InvoiceNo') || rowText(row, 'VoucherNo') || rowText(row, 'BillNo') || defaultValue;
}

function rowInvoiceReference(row, defaultValue = '') {
  return rowText(row, 'ReferenceNo') || rowText(row, 'VoucherNo')
    || rowText(row, 'SupplierInvoiceNo') || rowText(row, 'BillNo') || defaultValue;
}

function ledgerNameKey(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function isEffectivelyBlankLedger(value) {
  const text = String(value || '').trim();
  if (!text) return true;
  const key = text.toLowerCase();
  if (['na', 'n/a', 'none', 'null', 'nil', 'not applicable', '* not applicable', '-', '--'].includes(key)) return true;
  const compact = text.replace(/,/g, '');
  if (/^[+-]?\d+(?:\.\d+)?$/.test(compact)) {
    const f = parseFloat(compact);
    return Number.isFinite(f) ? f === 0.0 : false;
  }
  return false;
}

function ledgerOrSuspense(value, fallback = SUSPENSE_LEDGER) {
  let text = String(value || '').trim();
  if (isEffectivelyBlankLedger(text)) text = '';
  return text || fallback;
}

function resolvePartyLedger(row, isPurchaseMode = false) {
  const partyLedger = rowText(row, 'PartyLedger');
  if (partyLedger) return partyLedger;
  const fallbackKeys = isPurchaseMode
    ? ['PartyName', 'Party Name', 'SupplierName', 'Supplier', 'VendorName', 'Vendor', 'BillToName', 'Bill To Name', 'Party']
    : ['PartyName', 'Party Name', 'BuyerName', 'CustomerName', 'Customer', 'BillToName', 'Bill To Name', 'Party'];
  return rowTextAny(row, fallbackKeys, '');
}

function normalizeStateForLedger(value) {
  const text = String(value || '').trim();
  if (['not applicable', '* not applicable', 'na', 'n/a'].includes(text.toLowerCase())) return '';
  return text;
}

function companyStaticBlock(company) {
  const selected = String(company || '').trim();
  if (!selected) return '';
  return `   <STATICVARIABLES><SVCURRENTCOMPANY>${xmlEscape(selected)}</SVCURRENTCOMPANY></STATICVARIABLES>`;
}

function normalizeStockUnitName(value) {
  const text = String(value || '').trim();
  if (!text) return 'Nos';
  const aliases = {
    no: 'Nos', 'no.': 'Nos', nos: 'Nos', 'nos.': 'Nos',
    number: 'Nos', numbers: 'Nos', piece: 'pcs', pieces: 'pcs',
  };
  return aliases[text.toLowerCase()] || text;
}

// ─── Voucher numbering ─────────────────────────────────────────────────

function normalizeVoucherNumberText(value) {
  let text = String(value || '').trim();
  if (!text) return '';
  if (text.endsWith('.0') && /^\d+$/.test(text.slice(0, -2))) text = text.slice(0, -2);
  return text;
}

function incrementVoucherNumberText(value) {
  const text = normalizeVoucherNumberText(value);
  if (!text) return null;
  if (/^\d+$/.test(text)) return String(parseInt(text, 10) + 1);
  const match = text.match(/(\d+)(?!.*\d)/);
  if (!match) return null;
  const digits = match[1];
  const start = match.index;
  const end = start + digits.length;
  return `${text.slice(0, start)}${String(parseInt(digits, 10) + 1).padStart(digits.length, '0')}${text.slice(end)}`;
}

function voucherNumberWithOffset(startValue, offset) {
  const text = normalizeVoucherNumberText(startValue);
  if (!text) return '';
  if (offset <= 0) return text;
  if (/^\d+$/.test(text)) return String(parseInt(text, 10) + offset);
  let current = text;
  for (let idx = 0; idx < offset; idx++) {
    const nextValue = incrementVoucherNumberText(current);
    if (!nextValue) return `${text}-${idx + 2}`;
    current = nextValue;
  }
  return current;
}

// ─── Party / company GST context ──────────────────────────────────────

function collectPartyContext(row, partyLedger, allowPlaceOfSupplyColumn = true) {
  const partyLedgerRaw = ledgerOrSuspense(partyLedger);
  const partyNameRaw = rowTextAny(row, ['PartyName', 'BuyerName', 'SupplierName', 'BillToName', 'Party'], partyLedgerRaw);
  const mailingNameRaw = rowTextAny(row, ['PartyMailingName', 'MailingName', 'BillingName', 'Supplier', 'Bill To Name'], partyNameRaw);
  const gstinRaw = rowTextAny(row, ['PartyGSTIN', 'GSTIN', 'GSTIN/UIN', 'GSTIN UIN', 'Party GSTIN', 'SupplierGSTIN', 'Supplier GSTIN', 'GST No', 'GST Number']).toUpperCase();

  let stateRaw = normalizeStateForLedger(rowTextAny(row, ['PartyState', 'State', 'StateName', 'State Name']));
  if (!stateRaw && gstinRaw) stateRaw = _stateNameFromGstin(gstinRaw);

  let placeRaw = allowPlaceOfSupplyColumn
    ? rowTextAny(row, ['PlaceOfSupply', 'Place Of Supply', 'Place of Supply', 'POS', 'StateOfSupply'], stateRaw)
    : stateRaw;
  placeRaw = normalizeStateForLedger(placeRaw);
  if (!placeRaw && gstinRaw) placeRaw = _stateNameFromGstin(gstinRaw);
  if (!stateRaw && placeRaw && !gstinRaw) stateRaw = placeRaw;

  const countryRaw = rowTextAny(row, ['PartyCountry', 'Country', 'Country Name', 'CountryOfResidence'], 'India');
  const pincodeRaw = rowTextAny(row, ['PartyPincode', 'Pincode', 'PinCode', 'PIN', 'PIN Code', 'PostalCode', 'Postal Code']);
  const address1Raw = rowTextAny(row, ['PartyAddress1', 'PartyAddressLine1', 'Address1', 'Address Line 1', 'AddressLine1', 'BillToAddress', 'Address']);
  const address2Raw = rowTextAny(row, ['PartyAddress2', 'PartyAddressLine2', 'Address2', 'Address Line 2', 'AddressLine2']);

  const gstAppRaw = rowTextAny(row, ['GSTApplicable', 'GST Applicable', 'IsGSTApplicable', 'GST']);
  const regTypeRaw = rowTextAny(row, ['GSTRegistrationType', 'GST Registration Type', 'GST Reg Type', 'RegistrationType', 'Registration Type', 'RegType', 'Reg Type']);
  let regType = _normalizeGstRegistrationType(regTypeRaw, gstinRaw, gstAppRaw);
  if (regType.toLowerCase() === 'regular' && !gstinRaw) regType = '';

  return {
    party_ledger: partyLedgerRaw, party_name: partyNameRaw, mailing_name: mailingNameRaw,
    gstin: gstinRaw, state: stateRaw, place_of_supply: placeRaw, country: countryRaw || 'India',
    pincode: pincodeRaw, address1: address1Raw, address2: address2Raw, registration_type: regType,
  };
}

function appendInvoicePartyContextXml(addLine, partyContext, opts = {}) {
  const {
    includeBasicBuyer = false, includeState = true, includePlaceOfSupply = true, placeOfSupplyOverride = null,
  } = opts;
  const partyName = xmlEscape(partyContext.party_name || '');
  const mailingName = xmlEscape(partyContext.mailing_name || partyContext.party_name || '');
  const partyGstin = xmlEscape(partyContext.gstin || '');
  const partyState = includeState ? xmlEscape(partyContext.state || '') : '';
  const placeSource = placeOfSupplyOverride === null ? (partyContext.place_of_supply || '') : placeOfSupplyOverride;
  const placeOfSupply = includePlaceOfSupply ? xmlEscape(placeSource) : '';
  const country = xmlEscape(partyContext.country || 'India');
  const pincode = xmlEscape(partyContext.pincode || '');
  const address1 = xmlEscape(partyContext.address1 || '');
  const address2 = xmlEscape(partyContext.address2 || '');
  let regTypeRaw = String(partyContext.registration_type || '').trim();
  const partyGstinRaw = String(partyContext.gstin || '').trim();
  const countryRaw = String(partyContext.country || 'India').trim();

  if (!regTypeRaw) {
    if (partyGstinRaw) regTypeRaw = 'Regular';
    else if (countryRaw && countryRaw.toLowerCase() !== 'india') regTypeRaw = 'Overseas';
    else regTypeRaw = 'Unregistered/Consumer';
  }
  if (regTypeRaw.toLowerCase() === 'unregistered') regTypeRaw = 'Unregistered/Consumer';
  const regType = xmlEscape(regTypeRaw);

  if (address1 || address2) {
    addLine('     <ADDRESS.LIST TYPE="String">');
    if (address1) addLine(`      <ADDRESS>${address1}</ADDRESS>`);
    if (address2) addLine(`      <ADDRESS>${address2}</ADDRESS>`);
    addLine('     </ADDRESS.LIST>');
  }

  if (regType) addLine(`     <GSTREGISTRATIONTYPE>${regType}</GSTREGISTRATIONTYPE>`);
  addLine('     <VATDEALERTYPE>Regular</VATDEALERTYPE>');

  if (partyState) {
    addLine(`     <STATENAME>${partyState}</STATENAME>`);
    addLine(`     <PARTYSTATENAME>${partyState}</PARTYSTATENAME>`);
  }
  addLine(`     <COUNTRYOFRESIDENCE>${country}</COUNTRYOFRESIDENCE>`);
  if (partyGstin) addLine(`     <PARTYGSTIN>${partyGstin}</PARTYGSTIN>`);
  if (placeOfSupply) addLine(`     <PLACEOFSUPPLY>${placeOfSupply}</PLACEOFSUPPLY>`);
  if (partyName) {
    addLine(`     <PARTYNAME>${partyName}</PARTYNAME>`);
    addLine(`     <BASICBASEPARTYNAME>${partyName}</BASICBASEPARTYNAME>`);
    if (includeBasicBuyer) addLine(`     <BASICBUYERNAME>${partyName}</BASICBUYERNAME>`);
  }
  if (mailingName) addLine(`     <PARTYMAILINGNAME>${mailingName}</PARTYMAILINGNAME>`);
  const consigneeName = mailingName || partyName;
  if (consigneeName) addLine(`     <CONSIGNEEMAILINGNAME>${consigneeName}</CONSIGNEEMAILINGNAME>`);
  const consigneeState = partyState || placeOfSupply;
  if (consigneeState) addLine(`     <CONSIGNEESTATENAME>${consigneeState}</CONSIGNEESTATENAME>`);
  addLine(`     <CONSIGNEECOUNTRYNAME>${country}</CONSIGNEECOUNTRYNAME>`);
  if (pincode) addLine(`     <PARTYPINCODE>${pincode}</PARTYPINCODE>`);
}

function stateKey(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function isGstinLike(value) {
  return /^\d{2}[A-Z0-9]{13}$/.test(String(value || '').trim().toUpperCase());
}

function pickCompanyGstRegistration(registrations, preferredState = '') {
  const rows = registrations || [];
  if (!rows.length) return {};
  const preferredKey = stateKey(preferredState);
  const score = (entry) => {
    const sk = stateKey(entry.state || '');
    const name = String(entry.name || '').trim();
    const gstin = String(entry.gstin || '').trim().toUpperCase();
    return [
      preferredKey && preferredKey === sk ? 1 : 0,
      name && !isGstinLike(name) ? 1 : 0,
      gstin ? 1 : 0,
    ];
  };
  let best = rows[0], bestScore = score(rows[0]);
  for (const entry of rows.slice(1)) {
    const s = score(entry);
    if (s[0] > bestScore[0] || (s[0] === bestScore[0] && (s[1] > bestScore[1] || (s[1] === bestScore[1] && s[2] > bestScore[2])))) {
      best = entry; bestScore = s;
    }
  }
  return best;
}

function resolveCompanyGstState(registrations, preferredState = '') {
  const selected = pickCompanyGstRegistration(registrations, preferredState);
  return selected ? String(selected.state || '').trim() : '';
}

function appendCompanyGstContextXml(addLine, partyContext, companyGstRegistrations = null, preferPartyState = true) {
  const registrations = companyGstRegistrations || [];
  if (!registrations.length) return;
  let preferredState = '';
  if (preferPartyState) {
    preferredState = String(partyContext.place_of_supply || '').trim() || String(partyContext.state || '').trim();
  }
  const selected = pickCompanyGstRegistration(registrations, preferredState);
  if (!selected || !Object.keys(selected).length) return;

  let regNameRaw = String(selected.name || '').trim();
  const regGstinRaw = String(selected.gstin || '').trim().toUpperCase();
  const regStateRaw = String(selected.state || '').trim();
  if (!regNameRaw && regStateRaw) regNameRaw = `${regStateRaw} Registration`;
  if (!regNameRaw && regGstinRaw) regNameRaw = regGstinRaw;

  const regName = xmlEscape(regNameRaw);
  const regGstin = xmlEscape(regGstinRaw);
  const regState = xmlEscape(regStateRaw);
  if (regName && regGstin) {
    addLine(`     <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="${regGstin}">${regName}</GSTREGISTRATION>`);
    addLine(`     <CMPGSTIN>${regGstin}</CMPGSTIN>`);
    addLine('     <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>');
  }
  if (regState) addLine(`     <CMPGSTSTATE>${regState}</CMPGSTSTATE>`);
}

function appendTaxObjectAllocationXml(addLine, taxClassification) {
  const taxClass = xmlEscape(String(taxClassification || '').trim());
  if (!taxClass) return;
  if (['igst', 'cgst', 'sgst', 'utgst', 'cess'].includes(taxClass.toLowerCase())) return;
  addLine('      <TAXOBJECTALLOCATIONS.LIST>');
  addLine('       <TAXOBJECTALLOCATIONS>');
  addLine('        <TAXTYPE>GST</TAXTYPE>');
  addLine('        <TAXABILITY>Taxable</TAXABILITY>');
  addLine(`        <TAXCLASSIFICATIONNAME>${taxClass}</TAXCLASSIFICATIONNAME>`);
  addLine('       </TAXOBJECTALLOCATIONS>');
  addLine('      </TAXOBJECTALLOCATIONS.LIST>');
}

function gstTransactionType(regType, gstin = '') {
  const key = String(regType || '').trim().toLowerCase();
  if (['regular', 'registered'].includes(key) && gstin) return 'Tax Invoice';
  if (['composition', 'consumer', 'unregistered', ''].includes(key)) return 'Unregistered';
  if (key === 'overseas') return 'Overseas';
  if (['sez', 'sez unit', 'sez developer'].includes(key)) return 'SEZ exports with payment';
  return gstin ? 'Tax Invoice' : 'Unregistered';
}

function pickTaxLedgerName(row, ledgerKeys, rateValue, defaultName, amountValue = 0.0) {
  let taxLedgerRaw = rowTextAny(row, ledgerKeys, '');
  if (isEffectivelyBlankLedger(taxLedgerRaw)) taxLedgerRaw = '';
  if ((rateValue > 0 || amountValue > 0) && !taxLedgerRaw) taxLedgerRaw = defaultName;
  return ledgerOrSuspense(taxLedgerRaw);
}

// ─── Row consolidation (multi-line invoices -> one voucher) ──────────────

function consolidateAccountingRows(rows, resolvedMode, resolvedCustomDate) {
  const consolidatedRows = [];
  const groupMap = new Map();
  for (const r of rows) {
    let sourceDate;
    if (resolvedMode === 'current') sourceDate = new Date();
    else if (resolvedMode === 'custom') sourceDate = resolvedCustomDate;
    else sourceDate = resolveExcelDate(r);
    const dt = tallyDate(sourceDate);

    const vnoRaw = rowVoucherNumber(r);
    const refRaw = rowInvoiceReference(r, vnoRaw);
    const invKey = String(refRaw || vnoRaw).trim().toLowerCase();
    const gstinKey = (rowText(r, 'PartyGSTIN') || rowText(r, 'GSTIN')).trim().toLowerCase();
    const partyRaw = ledgerOrSuspense(rowText(r, 'PartyLedger'));
    const key = `${dt}|${invKey}|${gstinKey}`;

    if (groupMap.has(key)) {
      const existingR = groupMap.get(key);
      const existingParty = ledgerOrSuspense(resolvePartyLedger(existingR, true));
      if (partyRaw.trim().toLowerCase() !== existingParty.trim().toLowerCase()) {
        throw new Error(`Conflicting Party Ledgers for Invoice '${refRaw || vnoRaw}': '${existingParty}' vs '${partyRaw}'`);
      }
      if (!existingR.acct_legs) existingR.acct_legs = [];
      existingR.acct_legs.push({ ...r });
    } else {
      const newR = { ...r };
      groupMap.set(key, newR);
      consolidatedRows.push(newR);
    }
  }
  return consolidatedRows;
}

function effectiveTaxAmount(row, amountKey, rateKey) {
  const amt = rowFloat(row, amountKey, 0.0);
  if (amt > 0) return amt;
  const rate = rowFloat(row, rateKey, 0.0);
  const taxable = Math.abs(rowFloat(row, 'TaxableValue', 0.0));
  return rate > 0 && taxable > 0 ? Math.round((taxable * rate / 100) * 100) / 100 : 0.0;
}

function consolidateItemRows(rows, resolvedMode, resolvedCustomDate) {
  const consolidatedRows = [];
  const groupMap = new Map();
  for (const r of rows) {
    let sourceDate;
    if (resolvedMode === 'current') sourceDate = new Date();
    else if (resolvedMode === 'custom') sourceDate = resolvedCustomDate;
    else sourceDate = resolveExcelDate(r);
    const dt = tallyDate(sourceDate);

    const vnoRaw = rowVoucherNumber(r);
    const supplierInvoiceRaw = rowInvoiceReference(r, vnoRaw);
    const invKey = String(supplierInvoiceRaw || vnoRaw).trim().toLowerCase();
    const gstinKey = (rowText(r, 'PartyGSTIN') || rowText(r, 'GSTIN')).trim().toLowerCase();
    const partyRaw = ledgerOrSuspense(resolvePartyLedger(r, true));
    const key = `${dt}|${invKey}|${gstinKey}`;

    if (groupMap.has(key)) {
      const existingR = groupMap.get(key);
      const existingParty = ledgerOrSuspense(resolvePartyLedger(existingR, true));
      if (partyRaw.trim().toLowerCase() !== existingParty.trim().toLowerCase()) {
        throw new Error(`Conflicting Party Ledgers for Invoice '${supplierInvoiceRaw || vnoRaw}': '${existingParty}' vs '${partyRaw}'`);
      }

      const exCgst = effectiveTaxAmount(existingR, 'CGST Amount', 'CGSTRate');
      const exSgst = effectiveTaxAmount(existingR, 'SGST Amount', 'SGSTRate');
      const exIgst = effectiveTaxAmount(existingR, 'IGST Amount', 'IGSTRate');
      const newCgst = effectiveTaxAmount(r, 'CGST Amount', 'CGSTRate');
      const newSgst = effectiveTaxAmount(r, 'SGST Amount', 'SGSTRate');
      const newIgst = effectiveTaxAmount(r, 'IGST Amount', 'IGSTRate');

      existingR['TaxableValue'] = rowFloat(existingR, 'TaxableValue', 0.0) + rowFloat(r, 'TaxableValue', 0.0);
      existingR['CGST Amount'] = exCgst + newCgst;
      existingR['SGST Amount'] = exSgst + newSgst;
      existingR['IGST Amount'] = exIgst + newIgst;
      existingR['Cess Amount'] = rowFloat(existingR, 'Cess Amount', 0.0) + rowFloat(r, 'Cess Amount', 0.0);
      existingR['Invoice Value'] = rowFloat(existingR, 'Invoice Value', 0.0) + rowFloat(r, 'Invoice Value', 0.0);
      existingR['CGSTRate'] = 0;
      existingR['SGSTRate'] = 0;
      existingR['IGSTRate'] = 0;

      if (!existingR.items) {
        existingR.items = [];
        const origIt = {
          ItemName: rowText(existingR, 'ItemName') || rowText(existingR, 'Item') || rowText(existingR, 'StockItem') || rowText(existingR, 'ProductName'),
          Quantity: rowFloat(existingR, 'Quantity', 0.0) || rowFloat(existingR, 'Qty', 0.0) || rowFloat(existingR, 'Unit', 0.0),
          Rate: rowFloat(existingR, 'Rate', 0.0),
          Per: rowText(existingR, 'Per') || rowText(existingR, 'UOM') || rowText(existingR, 'Unit'),
          GodownName: rowText(existingR, 'GodownName'),
          PurchaseLedger: rowText(existingR, 'PurchaseLedger'),
          SalesLedger: rowTextAny(existingR, ['SalesAccount', 'Sales Ledger', 'IncomeLedger', 'SalesLedger']),
        };
        if (origIt.ItemName) existingR.items.push(origIt);
      }

      const newRawItems = r.items || [];
      if (newRawItems.length) {
        existingR.items.push(...newRawItems);
      } else {
        const newIt = {
          ItemName: rowText(r, 'ItemName') || rowText(r, 'Item') || rowText(r, 'StockItem') || rowText(r, 'ProductName'),
          Quantity: rowFloat(r, 'Quantity', 0.0) || rowFloat(r, 'Qty', 0.0) || rowFloat(r, 'Unit', 0.0),
          Rate: rowFloat(r, 'Rate', 0.0),
          Per: rowText(r, 'Per') || rowText(r, 'UOM') || rowText(r, 'Unit'),
          GodownName: rowText(r, 'GodownName'),
          PurchaseLedger: rowText(r, 'PurchaseLedger'),
          SalesLedger: rowTextAny(r, ['SalesAccount', 'Sales Ledger', 'IncomeLedger', 'SalesLedger']),
        };
        if (newIt.ItemName) existingR.items.push(newIt);
      }
    } else {
      const newR = { ...r };
      groupMap.set(key, newR);
      consolidatedRows.push(newR);
    }
  }
  return consolidatedRows;
}

function cleanTaxLedger(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (['0', '0.0', 'none', 'na', 'n/a', '-'].includes(text.toLowerCase())) return '';
  return text;
}

function ledgerOrDefault(value, fallback = SUSPENSE_LEDGER) {
  const text = String(value || '').trim();
  return text || fallback;
}

function rowReferenceNumber(row, defaultValue = '') {
  return rowText(row, 'Reference') || rowText(row, 'RefNo') || rowText(row, 'Ref No')
    || rowText(row, 'BillRef') || rowText(row, 'BillNo') || rowText(row, 'Bill No')
    || rowText(row, 'InvoiceNo') || rowText(row, 'Invoice No')
    || rowText(row, 'SupplierInvoiceNo') || rowText(row, 'Supplier Invoice No') || defaultValue;
}

function appendCommonLedgerFlags(a, isParty, isDebit = null) {
  if (isDebit === null) isDebit = !isParty;
  a('      <GSTCLASS>Not Applicable</GSTCLASS>');
  a(`      <ISDEEMEDPOSITIVE>${isDebit ? 'Yes' : 'No'}</ISDEEMEDPOSITIVE>`);
  a('      <LEDGERFROMITEM>No</LEDGERFROMITEM>');
  a('      <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>');
  a(`      <ISPARTYLEDGER>${isParty ? 'Yes' : 'No'}</ISPARTYLEDGER>`);
  a('      <GSTOVERRIDDEN>No</GSTOVERRIDDEN>');
  a('      <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>');
}

function normalizeStockGroupName(value) {
  const text = String(value || '').trim();
  if (!text) return 'Primary';
  const ledgerLikeGroups = new Set([
    'indirect income', 'direct income', 'indirect expenses', 'direct expenses',
    'sales accounts', 'purchase accounts', 'sundry debtors', 'sundry creditors',
    'duties & taxes', 'duties and taxes', 'bank accounts', 'cash-in-hand', 'cash in hand',
  ]);
  if (ledgerLikeGroups.has(text.toLowerCase())) return 'Primary';
  return text;
}

module.exports = {
  SUSPENSE_LEDGER, fmtAmt, nameKey, tallyDate, normalizeManualDateToTally, appendRoundOffEntry,
  rowGet, resolveExcelDate, rowText, rowTextAny, rowFloat, rowVoucherNumber, rowInvoiceReference,
  ledgerNameKey, isEffectivelyBlankLedger, ledgerOrSuspense, resolvePartyLedger, normalizeStateForLedger,
  companyStaticBlock, normalizeStockUnitName, normalizeVoucherNumberText, incrementVoucherNumberText,
  voucherNumberWithOffset, collectPartyContext, appendInvoicePartyContextXml, stateKey, isGstinLike,
  pickCompanyGstRegistration, resolveCompanyGstState, appendCompanyGstContextXml, appendTaxObjectAllocationXml,
  gstTransactionType, pickTaxLedgerName, consolidateAccountingRows, effectiveTaxAmount, consolidateItemRows,
  cleanTaxLedger, ledgerOrDefault, rowReferenceNumber, appendCommonLedgerFlags, normalizeStockGroupName,
  xmlEscape, stateNameFromGstin: _stateNameFromGstin,
  normalizeGstApplicable: tallyBridge.normalizeGstApplicable,
  normalizeGstRegistrationType: _normalizeGstRegistrationType,
  fetchCompanyGstRegistrations: tallyBridge.fetchCompanyGstRegistrations,
};
