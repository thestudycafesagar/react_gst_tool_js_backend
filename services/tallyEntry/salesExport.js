/** Sales Export Accounting voucher builder (no GST — exempt, INR only).
 * 1:1 port of Backend_Tally/tally_entry/sales_export.py. */
const {
  fmtAmt, tallyDate, normalizeManualDateToTally, appendRoundOffEntry,
  rowText, rowTextAny, rowFloat, rowVoucherNumber, ledgerOrSuspense,
  resolvePartyLedger, voucherNumberWithOffset, resolveExcelDate, companyStaticBlock, xmlEscape,
} = require('./common');

/** Sales Export accounting voucher — no GST (exempt), INR only.
 * Excel columns: Date, InvoiceNo, PartyLedger, PartyName, Country,
 * SalesLedger, Amount, HSNCode, HSNDescription, Narration.
 * Returns [xml, voucherCount] — mirrors the Python tuple return. */
function generateSalesExportAccountingXml(rows, company, opts = {}) {
  const {
    useTodayDate = false, dateMode = '', customTallyDate = '', startVoucherNumber = null,
    voucherType = 'Sales Export', roundOffLedger = '',
  } = opts;
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

  const effectiveVchType = String(voucherType || 'Sales Export').trim() || 'Sales Export';
  const vchTypeEsc = xmlEscape(effectiveVchType);
  let voucherCount = 0;

  rows.forEach((r, idx) => {
    let sourceDate;
    if (resolvedMode === 'current') sourceDate = new Date();
    else if (resolvedMode === 'custom') sourceDate = resolvedCustomDate;
    else sourceDate = resolveExcelDate(r);
    const dt = tallyDate(sourceDate);

    const excelVno = rowVoucherNumber(r, '');
    let vnoRaw;
    if (excelVno) vnoRaw = excelVno;
    else if (startVoucherNumber != null) vnoRaw = voucherNumberWithOffset(startVoucherNumber, idx) || String(startVoucherNumber);
    else vnoRaw = '';
    const vno = xmlEscape(vnoRaw);

    const partyRaw = ledgerOrSuspense(resolvePartyLedger(r, false));
    const party = xmlEscape(partyRaw);
    const partyNameRaw = rowTextAny(r, ['PartyName', 'BuyerName', 'PartyMailingName'], partyRaw);
    const partyName = xmlEscape(partyNameRaw || partyRaw);

    const salesLedgerRaw = ledgerOrSuspense(rowText(r, 'SalesLedger'));
    const salesLedger = xmlEscape(salesLedgerRaw);

    const countryRaw = rowTextAny(r, ['Country', 'CountryOfResidence', 'PartyCountry'], 'Outside India');
    const country = xmlEscape(countryRaw || 'Outside India');

    const hsnCode = xmlEscape(rowTextAny(r, ['HSNCode', 'HSN', 'SACCode', 'SAC'], '9983'));
    const hsnDesc = xmlEscape(rowTextAny(r, ['HSNDescription', 'HSNDesc', 'Description', 'ServiceDescription'], 'Advertisement'));
    const narr = xmlEscape(rowText(r, 'Narration'));

    const inrTotal = Math.round(Math.abs(rowFloat(r, 'Amount', 0.0)) * 100) / 100;
    if (inrTotal <= 0) return;
    const roAmt = roundOffLedger ? Math.round((Math.round(inrTotal) - inrTotal) * 100) / 100 : 0.0;
    const roTotal = Math.round((inrTotal + roAmt) * 100) / 100;

    const partyAmtStr = `-${fmtAmt(roTotal)}`;
    const salesAmtStr = fmtAmt(inrTotal);

    voucherCount += 1;
    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a(`    <VOUCHER VCHTYPE="${vchTypeEsc}" ACTION="Create" OBJVIEW="Invoice Voucher View">`);
    a(`     <DATE>${dt}</DATE>`);
    a('     <GSTREGISTRATIONTYPE>Unknown</GSTREGISTRATIONTYPE>');
    a('     <VATDEALERTYPE>Regular</VATDEALERTYPE>');
    a(`     <COUNTRYOFRESIDENCE>${country}</COUNTRYOFRESIDENCE>`);
    a(`     <VOUCHERTYPENAME>${vchTypeEsc}</VOUCHERTYPENAME>`);
    a(`     <PARTYNAME>${party}</PARTYNAME>`);
    a(`     <PARTYLEDGERNAME>${party}</PARTYLEDGERNAME>`);
    a(`     <VOUCHERNUMBER>${vno}</VOUCHERNUMBER>`);
    a(`     <BASICBUYERNAME>${partyName}</BASICBUYERNAME>`);
    a(`     <PARTYMAILINGNAME>${partyName}</PARTYMAILINGNAME>`);
    a(`     <CONSIGNEEMAILINGNAME>${partyName}</CONSIGNEEMAILINGNAME>`);
    a(`     <CONSIGNEECOUNTRYNAME>${country}</CONSIGNEECOUNTRYNAME>`);
    a(`     <BASICBASEPARTYNAME>${partyName}</BASICBASEPARTYNAME>`);
    a('     <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>');
    a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
    a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>');
    a('     <VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>');
    a('     <ISINVOICE>Yes</ISINVOICE>');
    a('     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>');
    if (narr) a(`     <NARRATION>${narr}</NARRATION>`);

    a('     <LEDGERENTRIES.LIST>');
    a(`      <LEDGERNAME>${party}</LEDGERNAME>`);
    a('      <GSTCLASS>Not Applicable</GSTCLASS>');
    a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
    a('      <LEDGERFROMITEM>No</LEDGERFROMITEM>');
    a('      <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>');
    a('      <ISPARTYLEDGER>Yes</ISPARTYLEDGER>');
    a('      <GSTOVERRIDDEN>No</GSTOVERRIDDEN>');
    a(`      <AMOUNT>${partyAmtStr}</AMOUNT>`);
    a('     </LEDGERENTRIES.LIST>');

    a('     <LEDGERENTRIES.LIST>');
    a(`      <LEDGERNAME>${salesLedger}</LEDGERNAME>`);
    a('      <GSTCLASS>Not Applicable</GSTCLASS>');
    a('      <GSTOVRDNTAXABILITY>Exempt</GSTOVRDNTAXABILITY>');
    a('      <GSTSOURCETYPE>Ledger</GSTSOURCETYPE>');
    a(`      <GSTLEDGERSOURCE>${salesLedger}</GSTLEDGERSOURCE>`);
    a('      <HSNSOURCETYPE>Ledger</HSNSOURCETYPE>');
    a(`      <HSNLEDGERSOURCE>${salesLedger}</HSNLEDGERSOURCE>`);
    a('      <GSTOVRDNTYPEOFSUPPLY>Services</GSTOVRDNTYPEOFSUPPLY>');
    a(`      <GSTHSNNAME>${hsnCode}</GSTHSNNAME>`);
    a(`      <GSTHSNDESCRIPTION>${hsnDesc}</GSTHSNDESCRIPTION>`);
    a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
    a('      <LEDGERFROMITEM>No</LEDGERFROMITEM>');
    a('      <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>');
    a('      <ISPARTYLEDGER>No</ISPARTYLEDGER>');
    a('      <GSTOVERRIDDEN>No</GSTOVERRIDDEN>');
    a(`      <AMOUNT>${salesAmtStr}</AMOUNT>`);
    a(`      <VATEXPAMOUNT>${salesAmtStr}</VATEXPAMOUNT>`);
    a('     </LEDGERENTRIES.LIST>');

    appendRoundOffEntry(a, roundOffLedger, roAmt, false);

    a('    </VOUCHER>');
    a('   </TALLYMESSAGE>');
  });

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return [lines.join('\n'), voucherCount];
}

module.exports = { generateSalesExportAccountingXml };
