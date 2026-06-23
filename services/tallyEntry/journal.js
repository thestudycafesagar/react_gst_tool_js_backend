/** Sale/Purchase Journal voucher builder. 1:1 port of Backend_Tally/tally_entry/journal.py. */
const {
  fmtAmt, tallyDate, normalizeManualDateToTally, appendRoundOffEntry,
  rowText, rowFloat, rowVoucherNumber, rowInvoiceReference, ledgerOrDefault,
  cleanTaxLedger, rowReferenceNumber, appendCommonLedgerFlags, resolveExcelDate,
  stateNameFromGstin, companyStaticBlock, xmlEscape,
} = require('./common');

/** Purchase: Expense Dr / Tax Dr / Party Cr. Sale: Party Dr / Sales Cr / Tax Cr.
 * Returns [xml, voucherCount] — mirrors the Python tuple return. */
function generateJournalXml(rows, company, opts = {}) {
  const {
    useTodayDate = false, dateMode = '', customTallyDate = '', includeVoucherNumber = true,
    includeBillAllocations = true, journalType = 'purchase', companyGstRegistrations = null,
    roundOffLedger = '',
  } = opts;
  const isSale = String(journalType || 'purchase').trim().toLowerCase() === 'sale';
  const lines = [];
  const a = (s) => lines.push(s);
  const companyStatic = companyStaticBlock(company);

  a('<?xml version="1.0" encoding="UTF-8"?>');
  a('<ENVELOPE>');
  a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>');
  a(' <BODY><IMPORTDATA>');
  a('  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>');
  if (companyStatic) a(companyStatic);
  a('  </REQUESTDESC>');
  a('  <REQUESTDATA>');

  let resolvedMode = String(dateMode || (useTodayDate ? 'current' : 'excel')).trim().toLowerCase();
  if (!['current', 'excel', 'custom'].includes(resolvedMode)) resolvedMode = useTodayDate ? 'current' : 'excel';
  const resolvedCustomDate = resolvedMode === 'custom' ? normalizeManualDateToTally(customTallyDate) : '';

  const cmpRegs = companyGstRegistrations || [];
  let cmpGstin = '', cmpState = '', cmpName = '';
  if (cmpRegs.length) {
    const r0 = cmpRegs[0];
    cmpGstin = xmlEscape(String(r0.gstin || '').trim());
    cmpState = xmlEscape(String(r0.state || '').trim());
    cmpName = xmlEscape(String(r0.name || '').trim());
  }

  let voucherCount = 0;

  rows.forEach((r, idx) => {
    const taxable = rowFloat(r, 'TaxableValue', 0.0);
    if (taxable <= 0) return;

    let sourceDate;
    if (resolvedMode === 'current') sourceDate = new Date();
    else if (resolvedMode === 'custom') sourceDate = resolvedCustomDate;
    else sourceDate = resolveExcelDate(r);
    const dt = tallyDate(sourceDate);
    const excelVno = rowVoucherNumber(r, '');
    const vnoRaw = excelVno || String(idx + 1);
    const invoiceRefRaw = rowInvoiceReference(r, vnoRaw); // eslint-disable-line no-unused-vars

    const partyRaw = ledgerOrDefault(rowText(r, 'PartyLedger'));
    let particularRaw = rowText(r, 'Particular') || rowText(r, 'Particulars')
      || rowText(r, 'SalesLedger') || rowText(r, 'Sales Ledger')
      || rowText(r, 'ExpenseLedger') || rowText(r, 'PurchaseLedger') || 'Journal Adjustment';
    particularRaw = ledgerOrDefault(particularRaw, 'Journal Adjustment');

    let cgstLedgerRaw = cleanTaxLedger(rowText(r, 'CGSTLedger'));
    let sgstLedgerRaw = cleanTaxLedger(rowText(r, 'SGSTLedger'));
    let igstLedgerRaw = cleanTaxLedger(rowText(r, 'IGSTLedger'));

    const cgstRate = rowFloat(r, 'CGSTRate', 0.0);
    const sgstRate = rowFloat(r, 'SGSTRate', 0.0);
    const igstRate = rowFloat(r, 'IGSTRate', 0.0);

    const cgstAmtExplicit = rowFloat(r, 'CGST Amount', 0.0);
    const sgstAmtExplicit = rowFloat(r, 'SGST Amount', 0.0);
    const igstAmtExplicit = rowFloat(r, 'IGST Amount', 0.0);

    let cgstAmt = cgstRate > 0 ? Math.round(taxable * cgstRate / 100 * 100) / 100 : cgstAmtExplicit;
    let sgstAmt = sgstRate > 0 ? Math.round(taxable * sgstRate / 100 * 100) / 100 : sgstAmtExplicit;
    let igstAmt = igstRate > 0 ? Math.round(taxable * igstRate / 100 * 100) / 100 : igstAmtExplicit;

    if (!cgstLedgerRaw && cgstAmt > 0) cgstLedgerRaw = 'CGST';
    if (!sgstLedgerRaw && sgstAmt > 0) sgstLedgerRaw = 'SGST';
    if (!igstLedgerRaw && igstAmt > 0) igstLedgerRaw = 'IGST';

    cgstAmt = cgstLedgerRaw ? cgstAmt : 0.0;
    sgstAmt = sgstLedgerRaw ? sgstAmt : 0.0;
    igstAmt = igstLedgerRaw ? igstAmt : 0.0;

    const total = taxable + cgstAmt + sgstAmt + igstAmt;
    const hasGst = (cgstAmt + sgstAmt + igstAmt) > 0;
    const roAmtJnl = roundOffLedger ? Math.round((Math.round(total) - total) * 100) / 100 : 0.0;
    const roTotalJnl = Math.round((total + roAmtJnl) * 100) / 100;

    const jnlTdsLedgerRaw = !isSale ? (rowText(r, 'TDSLedger') || rowText(r, 'TDS Ledger') || rowText(r, 'Tds Ledger')) : '';
    const jnlTdsRate = !isSale ? (rowFloat(r, 'TDSRate', 0.0) || rowFloat(r, 'TDS Rate', 0.0)) : 0.0;
    const jnlTdsAmountRaw = !isSale ? (rowFloat(r, 'TDSAmount', 0.0) || rowFloat(r, 'TDS Amount', 0.0)) : 0.0;
    let jnlTdsAmount;
    if (jnlTdsLedgerRaw && jnlTdsAmountRaw <= 0 && jnlTdsRate > 0) jnlTdsAmount = Math.round(taxable * jnlTdsRate / 100 * 100) / 100;
    else jnlTdsAmount = Math.abs(jnlTdsAmountRaw);
    const jnlTdsLed = xmlEscape(jnlTdsLedgerRaw);
    const jnlPartyTotal = (jnlTdsLed && jnlTdsAmount > 0) ? roTotalJnl - jnlTdsAmount : roTotalJnl;

    const billReferenceRaw = rowReferenceNumber(r, '');
    const voucherReferenceRaw = billReferenceRaw || (includeVoucherNumber ? vnoRaw : '');
    const vno = xmlEscape(vnoRaw);
    const reference = xmlEscape(voucherReferenceRaw);
    const billReference = xmlEscape(billReferenceRaw);
    const party = xmlEscape(partyRaw);
    const particular = xmlEscape(particularRaw);
    const narration = xmlEscape(rowText(r, 'Narration'));

    const partyGstinRaw = (rowText(r, 'GSTIN/UIN') || rowText(r, 'PartyGSTIN') || rowText(r, 'GSTIN')).trim().toUpperCase();
    const posRaw = String(rowText(r, 'PlaceOfSupply') || '').trim();
    const partyStateRaw = posRaw || (partyGstinRaw ? stateNameFromGstin(partyGstinRaw) : '');
    const placeOfSupply = xmlEscape(isSale ? partyStateRaw : (posRaw || cmpState));
    const partyGstin = xmlEscape(partyGstinRaw);
    const partyState = xmlEscape(partyStateRaw);

    voucherCount += 1;
    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a('    <VOUCHER REMOTEID="" VCHTYPE="Journal" ACTION="Create" OBJVIEW="Accounting Voucher View">');
    a(`     <DATE>${dt}</DATE>`);
    a('     <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>');
    a('     <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>');
    a('     <VCHENTRYMODE>Accounting Voucher View</VCHENTRYMODE>');
    a('     <ISINVOICE>No</ISINVOICE>');
    a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
    a('     <ISELIGIBLEFORITC>No</ISELIGIBLEFORITC>');

    if (hasGst) {
      a('     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>');
      a('     <GSTTRANSACTIONTYPE>Tax Invoice</GSTTRANSACTIONTYPE>');
      a('     <GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>');
      if (partyGstin) a(`     <PARTYGSTIN>${partyGstin}</PARTYGSTIN>`);
      if (cmpGstin && cmpName) {
        a(`     <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="${cmpGstin}">${cmpName}</GSTREGISTRATION>`);
        a(`     <CMPGSTIN>${cmpGstin}</CMPGSTIN>`);
        a('     <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>');
      }
      if (cmpState) a(`     <CMPGSTSTATE>${cmpState}</CMPGSTSTATE>`);
      if (placeOfSupply) a(`     <PLACEOFSUPPLY>${placeOfSupply}</PLACEOFSUPPLY>`);
    }
    if (reference) a(`     <REFERENCE>${reference}</REFERENCE>`);
    if (includeVoucherNumber && vno) a(`     <VOUCHERNUMBER>${vno}</VOUCHERNUMBER>`);
    if (narration) a(`     <NARRATION>${narration}</NARRATION>`);

    if (isSale) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${party}</LEDGERNAME>`);
      appendCommonLedgerFlags(a, true, true);
      if (hasGst) {
        a('      <GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>');
        if (partyGstin) a(`      <GSTIN>${partyGstin}</GSTIN>`);
        if (partyState) a(`      <STATENAME>${partyState}</STATENAME>`);
        a('      <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>');
      }
      a(`      <AMOUNT>-${fmtAmt(roTotalJnl)}</AMOUNT>`);
      if (includeBillAllocations && billReference) {
        a('      <BILLALLOCATIONS.LIST>');
        a(`       <NAME>${billReference}</NAME>`);
        a('       <BILLTYPE>New Ref</BILLTYPE>');
        a(`       <AMOUNT>-${fmtAmt(roTotalJnl)}</AMOUNT>`);
        a('      </BILLALLOCATIONS.LIST>');
      }
      a('     </LEDGERENTRIES.LIST>');

      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${particular}</LEDGERNAME>`);
      appendCommonLedgerFlags(a, false, false);
      a(`      <AMOUNT>${fmtAmt(taxable)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');

      for (const [ln, la] of [[cgstLedgerRaw, cgstAmt], [sgstLedgerRaw, sgstAmt], [igstLedgerRaw, igstAmt]]) {
        if (la > 0 && ln) {
          a('     <LEDGERENTRIES.LIST>');
          a(`      <LEDGERNAME>${xmlEscape(ln)}</LEDGERNAME>`);
          appendCommonLedgerFlags(a, false, false);
          a(`      <AMOUNT>${fmtAmt(la)}</AMOUNT>`);
          a('     </LEDGERENTRIES.LIST>');
        }
      }

      appendRoundOffEntry(a, roundOffLedger, roAmtJnl, false);
    } else {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${particular}</LEDGERNAME>`);
      appendCommonLedgerFlags(a, false);
      a(`      <AMOUNT>-${fmtAmt(taxable)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');

      if (cgstAmt > 0 && cgstLedgerRaw) {
        a('     <LEDGERENTRIES.LIST>');
        a(`      <LEDGERNAME>${xmlEscape(cgstLedgerRaw)}</LEDGERNAME>`);
        appendCommonLedgerFlags(a, false);
        a(`      <AMOUNT>-${fmtAmt(cgstAmt)}</AMOUNT>`);
        a('     </LEDGERENTRIES.LIST>');
      }
      if (sgstAmt > 0 && sgstLedgerRaw) {
        a('     <LEDGERENTRIES.LIST>');
        a(`      <LEDGERNAME>${xmlEscape(sgstLedgerRaw)}</LEDGERNAME>`);
        appendCommonLedgerFlags(a, false);
        a(`      <AMOUNT>-${fmtAmt(sgstAmt)}</AMOUNT>`);
        a('     </LEDGERENTRIES.LIST>');
      }
      if (igstAmt > 0 && igstLedgerRaw) {
        a('     <LEDGERENTRIES.LIST>');
        a(`      <LEDGERNAME>${xmlEscape(igstLedgerRaw)}</LEDGERNAME>`);
        appendCommonLedgerFlags(a, false);
        a(`      <AMOUNT>-${fmtAmt(igstAmt)}</AMOUNT>`);
        a('     </LEDGERENTRIES.LIST>');
      }

      if (jnlTdsLed && jnlTdsAmount > 0) {
        a('     <LEDGERENTRIES.LIST>');
        a(`      <LEDGERNAME>${jnlTdsLed}</LEDGERNAME>`);
        appendCommonLedgerFlags(a, false, false);
        a(`      <AMOUNT>${fmtAmt(jnlTdsAmount)}</AMOUNT>`);
        a('     </LEDGERENTRIES.LIST>');
      }

      appendRoundOffEntry(a, roundOffLedger, roAmtJnl, true);

      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${party}</LEDGERNAME>`);
      appendCommonLedgerFlags(a, true);
      if (hasGst) {
        a('      <GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>');
        if (partyGstin) a(`      <GSTIN>${partyGstin}</GSTIN>`);
        if (partyState) a(`      <STATENAME>${partyState}</STATENAME>`);
        a('      <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>');
      }
      a(`      <AMOUNT>${fmtAmt(jnlPartyTotal)}</AMOUNT>`);
      if (includeBillAllocations && billReference) {
        a('      <BILLALLOCATIONS.LIST>');
        a(`       <NAME>${billReference}</NAME>`);
        a('       <BILLTYPE>New Ref</BILLTYPE>');
        a(`       <AMOUNT>${fmtAmt(jnlPartyTotal)}</AMOUNT>`);
        a('      </BILLALLOCATIONS.LIST>');
      }
      a('     </LEDGERENTRIES.LIST>');
    }

    a('    </VOUCHER>');
    a('   </TALLYMESSAGE>');
  });

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return [lines.join('\n'), voucherCount];
}

module.exports = { generateJournalXml };
