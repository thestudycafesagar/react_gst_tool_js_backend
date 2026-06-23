/**
 * Direct HTTP/XML bridge to a local TallyPrime/Tally ERP 9 instance.
 *
 * 1:1 port of Backend_Tally/tally_bridge.py. Output-equivalent to the
 * Python version: XML parsing here is regex/tree-walk based rather than
 * ElementTree, but every extraction strategy (including the "parse once,
 * don't fall through to a looser match" semantics some functions rely on)
 * is preserved so results match for the same Tally response.
 */
const axios = require('axios');
const { parseXml, iterAll, findAll, findText } = require('../utils/xmlTree');
const {
  xmlEscape, fmtAmt, sanitizeTallyXml, normalizeCompanyName, isValidCompanyName,
  isGstinLike, stateNameFromGstin, normalizeStateForLedger, normalizeGstApplicable,
  normalizeGstRegistrationType, currentFyStart, parseStatementDatetime,
  normalizeManualDateToTally,
} = require('../utils/common');

function companyStaticBlock(company, fromDate = '', toDate = '', currentDate = '') {
  const parts = [];
  const selected = String(company || '').trim();
  if (selected) parts.push(`<SVCURRENTCOMPANY>${xmlEscape(selected)}</SVCURRENTCOMPANY>`);
  if (fromDate) parts.push(`<SVFROMDATE TYPE="Date">${fromDate}</SVFROMDATE>`);
  if (toDate) parts.push(`<SVTODATE TYPE="Date">${toDate}</SVTODATE>`);
  if (currentDate) parts.push(`<SVCURRENTDATE TYPE="Date">${currentDate}</SVCURRENTDATE>`);
  if (!parts.length) return '';
  return `   <STATICVARIABLES>${parts.join('')}</STATICVARIABLES>`;
}

function buildTallyUrl(host, port) {
  let hostText = String(host || 'localhost').trim();
  const portText = String(port || '9000').trim();
  if (hostText.startsWith('http://')) hostText = hostText.slice(7);
  else if (hostText.startsWith('https://')) hostText = hostText.slice(8);
  hostText = hostText.replace(/\/+$/, '').replace(/^\/+/, '') || 'localhost';
  if (hostText.includes('/')) hostText = hostText.split('/', 1)[0];
  if (!/^\d+$/.test(portText)) throw new Error('Port must be numeric.');
  return `http://${hostText}:${portText}`;
}

async function postTallyXml(tallyUrl, xmlPayload, timeoutMs = 15000) {
  const response = await axios.post(tallyUrl, xmlPayload, {
    headers: { 'Content-Type': 'application/xml' },
    timeout: timeoutMs,
    responseType: 'text',
    transformResponse: [(d) => d],
  });
  return response.data;
}

async function checkTallyConnection(tallyUrl, timeoutMs = 5000) {
  const probeXml = (
    '<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>'
    + '<BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Companies</REPORTNAME>'
    + '</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>'
  );
  try {
    await postTallyXml(tallyUrl, probeXml, timeoutMs);
    return { connected: true };
  } catch (exc) {
    if (exc.response) return { connected: false, error: `HTTP ${exc.response.status}` };
    if (exc.code === 'ECONNREFUSED' || exc.code === 'ENOTFOUND' || exc.code === 'ECONNABORTED') {
      return { connected: false, error: 'Could not reach Tally. Is it running with the ODBC/HTTP server enabled?' };
    }
    return { connected: false, error: exc.message };
  }
}

function isPartyParent(parent) {
  const key = String(parent || '').replace(/\s+/g, ' ').trim().toLowerCase();
  return key === 'sundry debtors' || key === 'sundry creditors';
}

function isDutiesParent(parent) {
  const key = String(parent || '').replace(/\s+/g, ' ').trim().toLowerCase();
  return ['duties & taxes', 'duties and taxes', 'duty'].includes(key);
}

function extractCompanyNames(responseText) {
  const names = new Set();
  try {
    const root = parseXml(responseText);
    for (const node of iterAll(root)) {
      const tag = node.tag.toUpperCase();
      const txt = normalizeCompanyName(node.text);
      const attrName = normalizeCompanyName(node.attrs.NAME || '');
      if (['COMPANYNAME', 'SVCURRENTCOMPANY', 'CURRENTCOMPANY'].includes(tag) && isValidCompanyName(txt)) names.add(txt);
      if (tag.includes('COMPANY') && isValidCompanyName(attrName)) names.add(attrName);
      if (tag === 'COMPANY' && isValidCompanyName(txt)) names.add(txt);
    }
  } catch {
    // fall through to regex-only extraction below
  }
  const patterns = [
    /COMPANY[^>]*NAME="([^"]+)"/gi,
    /<COMPANYNAME>([\s\S]*?)<\/COMPANYNAME>/gi,
    /<SVCURRENTCOMPANY>([\s\S]*?)<\/SVCURRENTCOMPANY>/gi,
  ];
  for (const pattern of patterns) {
    for (const match of responseText.matchAll(pattern)) {
      const value = normalizeCompanyName(match[1]);
      if (isValidCompanyName(value)) names.add(value);
    }
  }
  return names;
}

async function fetchTallyCompanies(tallyUrl, timeoutMs = 15000) {
  const requestsXml = [
    ['report-list-companies',
      '<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>'
      + '<BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Companies</REPORTNAME>'
      + '<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>'
      + '</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>'],
    ['collection-company',
      '<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
      + '<TYPE>Collection</TYPE><ID>Company Collection</ID></HEADER><BODY><DESC>'
      + '<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>'
      + "<TDL><TDLMESSAGE><COLLECTION NAME='Company Collection'>"
      + '<TYPE>Company</TYPE><FETCH>Name</FETCH><NATIVEMETHOD>Name</NATIVEMETHOD>'
      + '</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>'],
  ];
  const companies = new Set();
  const errors = [];
  for (const [label, xmlPayload] of requestsXml) {
    try {
      const responseText = await postTallyXml(tallyUrl, xmlPayload, timeoutMs);
      for (const n of extractCompanyNames(responseText)) companies.add(n);
    } catch (exc) {
      errors.push(`${label}: ${exc.message}`);
    }
  }
  const sortedCompanies = [...companies].sort((a, b) => a.toUpperCase().localeCompare(b.toUpperCase()));
  if (sortedCompanies.length) return { success: true, companies: sortedCompanies };
  return { success: false, error: errors.length ? errors.join('; ') : 'No companies returned by Tally.', companies: [] };
}

function extractCompanyGstRegistrations(responseText) {
  const registrations = [];
  const seen = new Set();
  try {
    const root = parseXml(responseText);
    for (const taxUnit of findAll(root, 'TAXUNIT')) {
      const taxType = String(taxUnit.attrs.TAXTYPE || findText(taxUnit, 'TAXTYPE') || '').trim().toUpperCase();
      if (taxType && taxType !== 'GST') continue;
      const nameRaw = String(taxUnit.attrs.NAME || findText(taxUnit, 'NAME') || '').trim();
      let gstinRaw = String(
        taxUnit.attrs.TAXREGISTRATION || findText(taxUnit, 'GSTREGNUMBER') || findText(taxUnit, 'GSTIN') || ''
      ).trim().toUpperCase();
      let stateRaw = String(findText(taxUnit, 'STATENAME') || '').trim();
      if (!gstinRaw && isGstinLike(nameRaw)) gstinRaw = nameRaw.toUpperCase();
      if (!stateRaw && gstinRaw) stateRaw = stateNameFromGstin(gstinRaw);
      if (!gstinRaw) continue;
      const name = normalizeCompanyName(nameRaw || gstinRaw);
      if (!name) continue;
      const key = `${name.toLowerCase()}|${gstinRaw}`;
      if (seen.has(key)) continue;
      seen.add(key);
      registrations.push({ name, gstin: gstinRaw, state: stateRaw });
    }
  } catch {
    // malformed XML -> no registrations, matches Python's except ET.ParseError: pass
  }
  return registrations;
}

async function fetchCompanyGstRegistrations(tallyUrl, company = '', timeoutMs = 15000) {
  let staticVars = '<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>';
  if (company) staticVars += `<SVCURRENTCOMPANY>${xmlEscape(company)}</SVCURRENTCOMPANY>`;
  staticVars += '</STATICVARIABLES>';
  const requestXml = (
    '<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
    + '<TYPE>Collection</TYPE><ID>Tax Unit Lookup</ID></HEADER>'
    + `<BODY><DESC>${staticVars}<TDL><TDLMESSAGE>`
    + "<COLLECTION NAME='Tax Unit Lookup'><TYPE>TaxUnit</TYPE>"
    + '<FETCH>Name,TaxType,TaxRegistration,GSTRegNumber,StateName,UseFor</FETCH>'
    + '<NATIVEMETHOD>Name</NATIVEMETHOD></COLLECTION>'
    + '</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>'
  );
  let responseText;
  try {
    responseText = await postTallyXml(tallyUrl, requestXml, timeoutMs);
  } catch (exc) {
    return { success: false, error: exc.message, registrations: [] };
  }
  const registrations = extractCompanyGstRegistrations(responseText);
  if (registrations.length) return { success: true, registrations };
  return { success: false, error: 'No GST registrations returned for selected company.', registrations: [] };
}

function safeInt(text, defaultValue = 0) {
  const n = parseFloat(String(text).trim());
  return Number.isFinite(n) ? Math.trunc(n) : defaultValue;
}

function parseTallyResponseDetails(responseText) {
  const details = {
    success: false, created: 0, altered: 0, deleted: 0, ignored: 0,
    errors: 0, exceptions: 0, line_errors: [], error: '',
  };
  let root;
  try {
    root = parseXml(responseText);
  } catch (exc) {
    details.error = `Could not parse Tally response: ${exc.message}`;
    return details;
  }
  const count = (tagName) => {
    const node = findAll(root, tagName)[0];
    return safeInt(node ? node.text : 0);
  };
  details.created = count('CREATED');
  details.altered = count('ALTERED');
  details.deleted = count('DELETED');
  details.ignored = count('IGNORED');
  details.errors = count('ERRORS');
  details.exceptions = count('EXCEPTIONS');

  const lineErrors = [];
  const seen = new Set();
  for (const node of findAll(root, 'LINEERROR')) {
    const text = (node.text || '').trim();
    if (text && !seen.has(text)) {
      seen.add(text);
      lineErrors.push(text);
    }
  }
  details.line_errors = lineErrors;
  details.success = details.errors === 0 && details.exceptions === 0 && lineErrors.length === 0;
  if (!details.success && lineErrors.length) details.error = lineErrors[0];
  return details;
}

const BANK_GROUPS = new Set(['bank accounts', 'bank account', 'bank', 'bank ods', 'bank od accounts']);
const CASH_GROUPS = new Set(['cash-in-hand', 'cash in hand']);
const BANK_OR_CASH_GROUPS = new Set([...BANK_GROUPS, ...CASH_GROUPS]);

function tagName(tag) {
  const raw = String(tag || '');
  return (raw.includes('}') ? raw.split('}')[1] : raw).toUpperCase();
}

function extractLedgers(respText) {
  const ledgers = [];
  for (const candidate of [respText, sanitizeTallyXml(respText)]) {
    if (!candidate) continue;
    let root;
    try {
      root = parseXml(candidate);
    } catch {
      continue;
    }
    for (const ledgerNode of iterAll(root)) {
      if (tagName(ledgerNode.tag) !== 'LEDGER') continue;
      let name = normalizeCompanyName(ledgerNode.attrs.NAME || '');
      let parent = '';
      for (const child of ledgerNode.children) {
        const tag = tagName(child.tag);
        const text = normalizeCompanyName(child.text);
        if (['NAME', 'LEDGERNAME'].includes(tag) && text) name = text;
        else if (['PARENT', 'PARENTGROUP'].includes(tag) && text) parent = text;
      }
      if (name && BANK_OR_CASH_GROUPS.has(parent.toLowerCase())) ledgers.push({ name, parent });
    }
    break; // parsed successfully (even if zero matches) — don't retry sanitized variant
  }
  return ledgers;
}

async function fetchBankLedgers(tallyUrl, company = '', timeoutMs = 15000) {
  let staticVars = '<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>';
  if (company) staticVars += `<SVCURRENTCOMPANY>${xmlEscape(company)}</SVCURRENTCOMPANY>`;
  staticVars += '</STATICVARIABLES>';

  const filterXml = (
    '<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
    + '<TYPE>Collection</TYPE><ID>BankLedgerCollection</ID></HEADER><BODY><DESC>'
    + `${staticVars}<TDL><TDLMESSAGE>`
    + "<COLLECTION NAME='BankLedgerCollection'>"
    + '<TYPE>Ledger</TYPE><NATIVEMETHOD>Name</NATIVEMETHOD><NATIVEMETHOD>Parent</NATIVEMETHOD>'
    + '<FILTER>IsBankOrCash</FILTER></COLLECTION>'
    + "<SYSTEM TYPE='Formulae' NAME='IsBankOrCash'>"
    + '$$InList:$Parent:"Bank Accounts":"Bank ODs":"Bank OD Accounts":"Cash-in-Hand"'
    + '</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>'
  );
  const allXml = (
    '<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
    + '<TYPE>Collection</TYPE><ID>AllLedgersWithParent</ID></HEADER><BODY><DESC>'
    + `${staticVars}<TDL><TDLMESSAGE>`
    + "<COLLECTION NAME='AllLedgersWithParent'>"
    + '<TYPE>Ledger</TYPE><NATIVEMETHOD>Name</NATIVEMETHOD><NATIVEMETHOD>Parent</NATIVEMETHOD>'
    + '</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>'
  );

  const dedup = (ledgers) => {
    const seen = new Set();
    const unique = [];
    for (const led of ledgers) {
      const key = led.name.toUpperCase();
      if (!seen.has(key)) { seen.add(key); unique.push(led); }
    }
    return unique.sort((a, b) => a.name.toUpperCase().localeCompare(b.name.toUpperCase()));
  };

  try {
    const resp1 = await postTallyXml(tallyUrl, filterXml, timeoutMs);
    const ledgers = extractLedgers(resp1);
    if (ledgers.length) return { success: true, ledgers: dedup(ledgers) };
  } catch {
    // fall through to attempt 2
  }

  try {
    const resp2 = await postTallyXml(tallyUrl, allXml, timeoutMs);
    const ledgers = extractLedgers(resp2);
    if (ledgers.length) return { success: true, ledgers: dedup(ledgers) };
    return { success: false, error: 'No bank/cash ledgers found. Ensure Bank Accounts / Cash-in-Hand group ledgers exist in Tally.', ledgers: [] };
  } catch (exc) {
    return { success: false, error: exc.message, ledgers: [] };
  }
}

async function fetchAllLedgerNames(tallyUrl, company = '', timeoutMs = 15000) {
  let staticVars = '<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>';
  if (company) staticVars += `<SVCURRENTCOMPANY>${xmlEscape(company)}</SVCURRENTCOMPANY>`;
  staticVars += '</STATICVARIABLES>';
  const xmlPayload = (
    '<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
    + '<TYPE>Collection</TYPE><ID>AllLedgerNames</ID></HEADER><BODY><DESC>'
    + `${staticVars}<TDL><TDLMESSAGE>`
    + "<COLLECTION NAME='AllLedgerNames'>"
    + '<TYPE>Ledger</TYPE><NATIVEMETHOD>Name</NATIVEMETHOD>'
    + '</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>'
  );
  let resp;
  try {
    resp = await postTallyXml(tallyUrl, xmlPayload, timeoutMs);
  } catch {
    return new Set();
  }
  const names = new Set();
  for (const candidate of [resp, sanitizeTallyXml(resp)]) {
    let root;
    try {
      root = parseXml(candidate);
    } catch {
      continue;
    }
    for (const node of iterAll(root)) {
      if (tagName(node.tag) !== 'LEDGER') continue;
      let name = normalizeCompanyName(node.attrs.NAME || '');
      if (!name) {
        for (const child of node.children) {
          if (['NAME', 'LEDGERNAME'].includes(tagName(child.tag)) && child.text) { name = normalizeCompanyName(child.text); break; }
        }
      }
      if (name) names.add(name.toUpperCase());
    }
    break;
  }
  if (!names.size) {
    for (const match of resp.matchAll(/<LEDGER[^>]+NAME="([^"]*)"/gi)) {
      const name = normalizeCompanyName(match[1]);
      if (name) names.add(name.toUpperCase());
    }
  }
  return names;
}

async function fetchAllLedgers(tallyUrl, company = '', timeoutMs = 15000) {
  let staticVars = '<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>';
  if (company) staticVars += `<SVCURRENTCOMPANY>${xmlEscape(company)}</SVCURRENTCOMPANY>`;
  staticVars += '</STATICVARIABLES>';
  const xmlPayload = (
    '<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
    + '<TYPE>Collection</TYPE><ID>AllLedgerNames</ID></HEADER><BODY><DESC>'
    + `${staticVars}<TDL><TDLMESSAGE>`
    + "<COLLECTION NAME='AllLedgerNames'>"
    + '<TYPE>Ledger</TYPE><NATIVEMETHOD>Name</NATIVEMETHOD>'
    + '</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>'
  );
  let resp;
  try {
    resp = await postTallyXml(tallyUrl, xmlPayload, timeoutMs);
  } catch {
    return [];
  }
  const byUpper = new Map();
  for (const candidate of [resp, sanitizeTallyXml(resp)]) {
    let root;
    try {
      root = parseXml(candidate);
    } catch {
      continue;
    }
    for (const node of iterAll(root)) {
      if (tagName(node.tag) !== 'LEDGER') continue;
      let name = normalizeCompanyName(node.attrs.NAME || '');
      if (!name) {
        for (const child of node.children) {
          if (['NAME', 'LEDGERNAME'].includes(tagName(child.tag)) && child.text) { name = normalizeCompanyName(child.text); break; }
        }
      }
      if (name && !byUpper.has(name.toUpperCase())) byUpper.set(name.toUpperCase(), name);
    }
    break;
  }
  if (!byUpper.size) {
    for (const match of resp.matchAll(/<LEDGER[^>]+NAME="([^"]*)"/gi)) {
      const name = normalizeCompanyName(match[1]);
      if (name && !byUpper.has(name.toUpperCase())) byUpper.set(name.toUpperCase(), name);
    }
  }
  return [...byUpper.values()].sort((a, b) => a.toUpperCase().localeCompare(b.toUpperCase()));
}

async function fetchAllStockItemNames(tallyUrl, company = '', timeoutMs = 15000) {
  let staticVars = '<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>';
  if (company) staticVars += `<SVCURRENTCOMPANY>${xmlEscape(company)}</SVCURRENTCOMPANY>`;
  staticVars += '</STATICVARIABLES>';
  const xmlPayload = (
    '<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>'
    + '<TYPE>Collection</TYPE><ID>Stock Item Collection</ID></HEADER>'
    + `<BODY><DESC>${staticVars}<TDL><TDLMESSAGE>`
    + "<COLLECTION NAME='Stock Item Collection'>"
    + '<TYPE>Stock Item</TYPE><FETCH>Name</FETCH><NATIVEMETHOD>Name</NATIVEMETHOD>'
    + '</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>'
  );
  let resp;
  try {
    resp = await postTallyXml(tallyUrl, xmlPayload, timeoutMs);
  } catch {
    return [];
  }
  const byUpper = new Map();
  for (const candidate of [resp, sanitizeTallyXml(resp)]) {
    let root;
    try {
      root = parseXml(candidate);
    } catch {
      continue;
    }
    for (const node of iterAll(root)) {
      const tag = tagName(node.tag);
      if (tag === 'STOCKITEM') {
        const name = normalizeCompanyName(node.attrs.NAME || '') || (node.text ? normalizeCompanyName(node.text) : '');
        if (name && !byUpper.has(name.toUpperCase())) byUpper.set(name.toUpperCase(), name);
      } else if (['STOCKITEMNAME', 'DSPSTOCKITEMNAME', 'NAME'].includes(tag) && node.text) {
        const name = normalizeCompanyName(node.text);
        if (name && !byUpper.has(name.toUpperCase())) byUpper.set(name.toUpperCase(), name);
      }
    }
    break;
  }
  if (!byUpper.size) {
    for (const match of resp.matchAll(/STOCKITEM[^>]*NAME="([^"]+)"/gi)) {
      const name = normalizeCompanyName(match[1]);
      if (name && !byUpper.has(name.toUpperCase())) byUpper.set(name.toUpperCase(), name);
    }
  }
  return [...byUpper.values()].sort((a, b) => a.toUpperCase().localeCompare(b.toUpperCase()));
}

function countVoucherEntries(rows, contraLedgerNames = null) {
  let payment = 0, receipt = 0, contra = 0;
  const contraSet = new Set((contraLedgerNames || []).map((n) => n.toUpperCase()));
  for (const r of rows) {
    const ledger = String(r.ledger || '').trim();
    const isContra = contraSet.size > 0 && contraSet.has(ledger.toUpperCase());
    const debitAmt = parseFloat(r.debit || 0);
    const creditAmt = parseFloat(r.credit || 0);
    if (debitAmt > 0) { if (isContra) contra += 1; else payment += 1; }
    if (creditAmt > 0) { if (isContra) contra += 1; else receipt += 1; }
  }
  return { payment, receipt, contra };
}

function pickSuspenseLedger(existingNames) {
  const upper = new Set([...(existingNames || [])].map((n) => String(n).trim().toUpperCase()));
  if (upper.has('SUSPENSE A/C')) return 'Suspense A/c';
  if (upper.has('SUSPENSE')) return 'Suspense';
  return 'Suspense A/c';
}

function generateLedgerXml(ledgers, company) {
  const lines = [];
  const a = (s) => lines.push(s);
  const companyStatic = companyStaticBlock(company);
  a('<?xml version="1.0" encoding="UTF-8"?>');
  a('<ENVELOPE>');
  a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>');
  a(' <BODY><IMPORTDATA>');
  a('  <REQUESTDESC><REPORTNAME>All Masters</REPORTNAME>');
  if (companyStatic) a(companyStatic);
  a('  </REQUESTDESC>');
  a('  <REQUESTDATA>');

  for (const led of ledgers) {
    const nameRaw = String(led.Name || '').trim();
    if (!nameRaw) continue;

    const parentRaw = String(led.Parent || 'Suspense A/c').trim() || 'Suspense A/c';
    const gstinRaw = String(led.GSTIN || '').trim().toUpperCase();
    let stateRaw = normalizeStateForLedger(String(led.StateOfSupply || '').trim());
    const address1Raw = String(led.Address1 || '').trim();
    const address2Raw = String(led.Address2 || '').trim();
    const pincodeRaw = String(led.Pincode || '').trim();
    const countryRaw = String(led.Country || '').trim() || 'India';
    const mailingNameRaw = String(led.MailingName || '').trim() || nameRaw;
    if (!stateRaw && gstinRaw) stateRaw = stateNameFromGstin(gstinRaw);

    const isParty = isPartyParent(parentRaw);
    const billwiseRaw = String(led.Billwise || '').trim();
    const billwiseOn = billwiseRaw ? ['yes', 'y', 'true', '1', 'on'].includes(billwiseRaw.toLowerCase()) : Boolean(isParty);

    const gstAppValue = normalizeGstApplicable(String(led.GSTApplicable || '').trim(), gstinRaw);
    const regType = normalizeGstRegistrationType(String(led.GSTRegistrationType || '').trim(), gstinRaw, gstAppValue);

    const name = xmlEscape(nameRaw);
    const parent = xmlEscape(parentRaw);
    const gstApp = xmlEscape(gstAppValue);
    const gstin = xmlEscape(gstinRaw);
    const state = xmlEscape(stateRaw);
    const address1 = xmlEscape(address1Raw);
    const address2 = xmlEscape(address2Raw);
    const pincode = xmlEscape(pincodeRaw);
    const country = xmlEscape(countryRaw);
    const mailingName = xmlEscape(mailingNameRaw);

    let taxTypeRaw = String(led.TypeOfTaxation || '').trim();
    if (['not applicable', 'na', 'n/a'].includes(taxTypeRaw.toLowerCase())) taxTypeRaw = '';
    const taxType = xmlEscape(taxTypeRaw);
    const gstRate = String(led.GSTRate || '').trim();
    const applicableFrom = currentFyStart();

    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a(`    <LEDGER NAME="${name}" RESERVEDNAME="" ACTION="Create">`);
    a(`     <NAME>${name}</NAME>`);
    a(`     <PARENT>${parent}</PARENT>`);
    a(`     <ISBILLWISEON>${billwiseOn ? 'Yes' : 'No'}</ISBILLWISEON>`);
    a('     <ISCOSTCENTRESON>No</ISCOSTCENTRESON>');
    a('     <ISINTERESTON>No</ISINTERESTON>');
    a('     <ALLOWINMOBILE>No</ALLOWINMOBILE>');
    a('     <ISUPDATINGTARGETID>No</ISUPDATINGTARGETID>');
    a('     <ASORIGINAL>Yes</ASORIGINAL>');
    a('     <AFFECTSSTOCK>No</AFFECTSSTOCK>');
    a('     <CURRENCYNAME>INR</CURRENCYNAME>');
    a(`     <COUNTRYOFRESIDENCE>${country}</COUNTRYOFRESIDENCE>`);

    if (isParty && gstApp) a(`     <GSTAPPLICABLE>${gstApp}</GSTAPPLICABLE>`);
    if (isParty && regType) a(`     <GSTREGISTRATIONTYPE>${xmlEscape(regType)}</GSTREGISTRATIONTYPE>`);
    if (isParty && gstin) a(`     <PARTYGSTIN>${gstin}</PARTYGSTIN>`);
    if (state) {
      a(`     <PRIORSTATENAME>${state}</PRIORSTATENAME>`);
      if (isParty) a(`     <LEDSTATENAME>${state}</LEDSTATENAME>`);
    }

    a('     <LANGUAGENAME.LIST>');
    a('      <NAME.LIST TYPE="String">');
    a(`       <NAME>${name}</NAME>`);
    a('      </NAME.LIST>');
    a('      <LANGUAGEID>1033</LANGUAGEID>');
    a('     </LANGUAGENAME.LIST>');

    if (isParty && (gstin || regType)) {
      a('     <LEDGSTREGDETAILS.LIST>');
      a(`      <APPLICABLEFROM>${applicableFrom}</APPLICABLEFROM>`);
      if (regType) a(`      <GSTREGISTRATIONTYPE>${xmlEscape(regType)}</GSTREGISTRATIONTYPE>`);
      if (state) a(`      <PLACEOFSUPPLY>${state}</PLACEOFSUPPLY>`);
      if (gstin) a(`      <GSTIN>${gstin}</GSTIN>`);
      a('      <ISOTHTERRITORYASSESSEE>No</ISOTHTERRITORYASSESSEE>');
      a('      <CONSIDERPURCHASEFOREXPORT>No</CONSIDERPURCHASEFOREXPORT>');
      a('      <ISTRANSPORTER>No</ISTRANSPORTER>');
      a('      <ISCOMMONPARTY>No</ISCOMMONPARTY>');
      a('     </LEDGSTREGDETAILS.LIST>');
    }

    if (isParty && (state || gstin || address1 || address2 || pincode || country)) {
      a('     <LEDMAILINGDETAILS.LIST>');
      if (address1 || address2) {
        a('      <ADDRESS.LIST TYPE="String">');
        if (address1) a(`       <ADDRESS>${address1}</ADDRESS>`);
        if (address2) a(`       <ADDRESS>${address2}</ADDRESS>`);
        a('      </ADDRESS.LIST>');
      }
      a(`      <APPLICABLEFROM>${applicableFrom}</APPLICABLEFROM>`);
      if (pincode) a(`      <PINCODE>${pincode}</PINCODE>`);
      a(`      <MAILINGNAME>${mailingName}</MAILINGNAME>`);
      if (state) a(`      <STATE>${state}</STATE>`);
      a(`      <COUNTRY>${country}</COUNTRY>`);
      a('     </LEDMAILINGDETAILS.LIST>');
    }

    if (isDutiesParent(parentRaw)) {
      if (taxType) a(`     <TAXTYPE>${taxType}</TAXTYPE>`);
      if (gstRate) a(`     <GSTRATE>${gstRate}</GSTRATE>`);
    }

    a('    </LEDGER>');
    a('   </TALLYMESSAGE>');
  }
  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return lines.join('\n');
}

async function fetchNextVoucherNumber(tallyUrl, company, voucherType, timeoutMs = 15000) {
  let staticVars = '<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>';
  staticVars += `<SVVOUCHERTYPENAME>${xmlEscape(voucherType)}</SVVOUCHERTYPENAME>`;
  if (company) staticVars += `<SVCURRENTCOMPANY>${xmlEscape(company)}</SVCURRENTCOMPANY>`;
  staticVars += '</STATICVARIABLES>';
  const xmlPayload = (
    '<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>'
    + '<BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Vouchers</REPORTNAME>'
    + `${staticVars}</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`
  );
  let resp;
  try {
    resp = await postTallyXml(tallyUrl, xmlPayload, timeoutMs);
  } catch (exc) {
    return { success: false, next_number: null, error: exc.message };
  }
  const numbers = [];
  for (const match of resp.matchAll(/<VOUCHERNUMBER>([\s\S]*?)<\/VOUCHERNUMBER>/gi)) {
    let text = String(match[1] || '').trim();
    if (text.endsWith('.0') && /^\d+$/.test(text.slice(0, -2))) text = text.slice(0, -2);
    if (/^\d+$/.test(text)) numbers.push(parseInt(text, 10));
  }
  return { success: true, next_number: numbers.length ? Math.max(...numbers) + 1 : 1 };
}

function generateBankVoucherXml(rows, company, bankLedger, opts = {}) {
  const {
    dateMode = 'excel', customTallyDate = '', suspenseLedger = 'Suspense A/c',
    paymentStartVno = null, receiptStartVno = null, contraLedgerNames = null, contraStartVno = null,
  } = opts;
  let mode = String(dateMode || 'excel').trim().toLowerCase();
  if (!['current', 'excel', 'custom'].includes(mode)) mode = 'excel';
  const resolvedCustomDate = mode === 'custom' ? normalizeManualDateToTally(customTallyDate) : '';

  const { from: periodFrom, to: periodTo, current: periodCurrent } = deriveImportPeriod(rows, mode, resolvedCustomDate);

  const lines = [];
  const a = (s) => lines.push(s);
  const companyStatic = companyStaticBlock(company, periodFrom, periodTo, periodCurrent);
  a('<?xml version="1.0" encoding="UTF-8"?>');
  a('<ENVELOPE>');
  a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>');
  a(' <BODY><IMPORTDATA>');
  a('  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>');
  if (companyStatic) a(companyStatic);
  a('  </REQUESTDESC>');
  a('  <REQUESTDATA>');

  const bankEsc = xmlEscape(bankLedger);
  const contraSet = new Set((contraLedgerNames || []).map((n) => n.toUpperCase()));
  let paymentCounter = 0, receiptCounter = 0, contraCounter = 0;

  rows.forEach((r, idx) => {
    let dt;
    if (mode === 'current') {
      dt = require('../utils/common').toTallyDateString(new Date());
    } else if (mode === 'custom') {
      dt = resolvedCustomDate;
    } else {
      const parsedSourceDate = parseStatementDatetime(r.date);
      if (parsedSourceDate == null) throw new Error(`Invalid date in row ${idx + 1}: '${r.date}'.`);
      dt = require('../utils/common').toTallyDateString(parsedSourceDate);
    }

    const ledgerName = String(r.ledger || '').trim() || suspenseLedger;
    const ledgerEsc = xmlEscape(ledgerName);
    const isContraRow = contraSet.size > 0 && contraSet.has(ledgerName.toUpperCase());

    const narrationParts = [r.description, r.chequeNo ? `Chq: ${r.chequeNo}` : ''].filter(Boolean);
    const narration = xmlEscape(narrationParts.join(' | '));

    const debitAmt = parseFloat(r.debit || 0);
    const creditAmt = parseFloat(r.credit || 0);
    if (debitAmt <= 0 && creditAmt <= 0) return;

    if (debitAmt > 0) {
      let vchType, vno;
      if (isContraRow) {
        vchType = 'Contra'; contraCounter += 1;
        vno = contraStartVno != null ? contraStartVno + contraCounter - 1 : contraCounter;
      } else {
        vchType = 'Payment'; paymentCounter += 1;
        vno = paymentStartVno != null ? paymentStartVno + paymentCounter - 1 : paymentCounter;
      }
      a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
      a(`    <VOUCHER VCHTYPE="${vchType}" ACTION="Create" OBJVIEW="Accounting Voucher View">`);
      a(`     <DATE>${dt}</DATE>`);
      a(`     <VOUCHERTYPENAME>${vchType}</VOUCHERTYPENAME>`);
      a(`     <VOUCHERNUMBER>${vno}</VOUCHERNUMBER>`);
      a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
      a('     <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>');
      if (narration) a(`     <NARRATION>${narration}</NARRATION>`);
      a('     <ALLLEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${ledgerEsc}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(debitAmt)}</AMOUNT>`);
      a('     </ALLLEDGERENTRIES.LIST>');
      a('     <ALLLEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${bankEsc}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(debitAmt)}</AMOUNT>`);
      a('     </ALLLEDGERENTRIES.LIST>');
      a('    </VOUCHER>');
      a('   </TALLYMESSAGE>');
    }

    if (creditAmt > 0) {
      let vchType, vno;
      if (isContraRow) {
        vchType = 'Contra'; contraCounter += 1;
        vno = contraStartVno != null ? contraStartVno + contraCounter - 1 : contraCounter;
      } else {
        vchType = 'Receipt'; receiptCounter += 1;
        vno = receiptStartVno != null ? receiptStartVno + receiptCounter - 1 : receiptCounter;
      }
      a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
      a(`    <VOUCHER VCHTYPE="${vchType}" ACTION="Create" OBJVIEW="Accounting Voucher View">`);
      a(`     <DATE>${dt}</DATE>`);
      a(`     <VOUCHERTYPENAME>${vchType}</VOUCHERTYPENAME>`);
      a(`     <VOUCHERNUMBER>${vno}</VOUCHERNUMBER>`);
      a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
      a('     <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>');
      if (narration) a(`     <NARRATION>${narration}</NARRATION>`);
      a('     <ALLLEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${bankEsc}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(creditAmt)}</AMOUNT>`);
      a('     </ALLLEDGERENTRIES.LIST>');
      a('     <ALLLEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${ledgerEsc}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(creditAmt)}</AMOUNT>`);
      a('     </ALLLEDGERENTRIES.LIST>');
      a('    </VOUCHER>');
      a('   </TALLYMESSAGE>');
    }
  });

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return lines.join('\n');
}

function deriveImportPeriod(rows, mode, resolvedCustomDate) {
  const { toTallyDateString } = require('../utils/common');
  if (mode === 'current') {
    const today = toTallyDateString(new Date());
    return { from: today, to: today, current: today };
  }
  if (mode === 'custom') {
    return { from: resolvedCustomDate, to: resolvedCustomDate, current: resolvedCustomDate };
  }
  const parsedDates = rows.map((r) => parseStatementDatetime(r.date)).filter((d) => d != null);
  if (!parsedDates.length) {
    const today = toTallyDateString(new Date());
    return { from: today, to: today, current: today };
  }
  const min = new Date(Math.min(...parsedDates.map((d) => d.getTime())));
  const max = new Date(Math.max(...parsedDates.map((d) => d.getTime())));
  return { from: toTallyDateString(min), to: toTallyDateString(max), current: toTallyDateString(max) };
}

module.exports = {
  xmlEscape, fmtAmt, companyStaticBlock, buildTallyUrl, postTallyXml, checkTallyConnection,
  isPartyParent, isDutiesParent, fetchTallyCompanies, fetchCompanyGstRegistrations,
  parseTallyResponseDetails, fetchBankLedgers, fetchAllLedgerNames, fetchAllLedgers,
  fetchAllStockItemNames, countVoucherEntries, pickSuspenseLedger, generateLedgerXml,
  fetchNextVoucherNumber, generateBankVoucherXml, panFromGstin: require('../utils/common').panFromGstin,
  stateNameFromGstin, normalizeCompanyName, normalizeGstApplicable, normalizeGstRegistrationType,
};
