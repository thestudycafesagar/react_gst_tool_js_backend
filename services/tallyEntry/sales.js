/** Sales Accounting Invoice and Sales Item Invoice voucher builders.
 * 1:1 port of Backend_Tally/tally_entry/sales.py. */
const {
  SUSPENSE_LEDGER, fmtAmt, tallyDate, normalizeManualDateToTally, appendRoundOffEntry,
  rowGet, resolveExcelDate, rowText, rowFloat, rowVoucherNumber, rowInvoiceReference,
  ledgerOrSuspense, resolvePartyLedger, companyStaticBlock, normalizeStockUnitName,
  voucherNumberWithOffset, collectPartyContext, appendInvoicePartyContextXml,
  appendCompanyGstContextXml, appendTaxObjectAllocationXml, gstTransactionType,
  pickTaxLedgerName, consolidateAccountingRows, consolidateItemRows, nameKey, xmlEscape,
} = require('./common');

function generateAccountingXml(rows, company, opts = {}) {
  const {
    useTodayDate = false, dateMode = '', customTallyDate = '', startVoucherNumber = null,
    companyGstRegistrations = null, voucherType = 'Sales', roundOffLedger = '',
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

  rows = consolidateAccountingRows(rows, resolvedMode, resolvedCustomDate);

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
    const invoiceRefRaw = rowInvoiceReference(r, vnoRaw);
    const invoiceRef = xmlEscape(invoiceRefRaw);
    const partyRaw = ledgerOrSuspense(resolvePartyLedger(r, true));
    const partyContext = collectPartyContext(r, partyRaw);
    const party = xmlEscape(partyRaw);
    const narr = xmlEscape(rowText(r, 'Narration'));

    const extraLegs = r.acct_legs || [];
    const allLegs = [r, ...extraLegs];

    const salesLegs = [];
    const gstCgst = new Map(), gstSgst = new Map(), gstIgst = new Map();

    for (const legR of allLegs) {
      const legTaxable = rowFloat(legR, 'TaxableValue', 0.0);
      const legCgstR = rowFloat(legR, 'CGSTRate', 0.0);
      const legCgstExplicit = rowFloat(legR, 'CGST Amount', 0.0);
      const legSgstR = rowFloat(legR, 'SGSTRate', 0.0);
      const legSgstExplicit = rowFloat(legR, 'SGST Amount', 0.0);
      const legIgstR = rowFloat(legR, 'IGSTRate', 0.0);
      const legIgstExplicit = rowFloat(legR, 'IGST Amount', 0.0);

      const legCgst = legCgstR > 0 ? Math.round(legTaxable * legCgstR / 100 * 100) / 100 : legCgstExplicit;
      const legSgst = legSgstR > 0 ? Math.round(legTaxable * legSgstR / 100 * 100) / 100 : legSgstExplicit;
      const legIgst = legIgstR > 0 ? Math.round(legTaxable * legIgstR / 100 * 100) / 100 : legIgstExplicit;

      const legSalesRaw = ledgerOrSuspense(rowText(legR, 'SalesLedger'));
      const legCgstName = pickTaxLedgerName(legR, ['CGSTLedger', 'CGST Ledger', 'CentralTaxLedger', 'Central Tax Ledger', 'Central Tax'], legCgstR, 'CGST', legCgstExplicit);
      const legSgstName = pickTaxLedgerName(legR, ['SGSTLedger', 'SGST Ledger', 'StateTaxLedger', 'State Tax Ledger', 'State Tax', 'UTGSTLedger', 'UTGST Ledger'], legSgstR, 'SGST', legSgstExplicit);
      const legIgstName = pickTaxLedgerName(legR, ['IGSTLedger', 'IGST Ledger', 'IntegratedTaxLedger', 'Integrated Tax Ledger', 'Integrated Tax'], legIgstR, 'IGST', legIgstExplicit);

      salesLegs.push({ sales_raw: legSalesRaw, taxable: legTaxable });
      if (legCgst) gstCgst.set(legCgstName, (gstCgst.get(legCgstName) || 0.0) + legCgst);
      if (legSgst) gstSgst.set(legSgstName, (gstSgst.get(legSgstName) || 0.0) + legSgst);
      if (legIgst) gstIgst.set(legIgstName, (gstIgst.get(legIgstName) || 0.0) + legIgst);
    }

    const totalTaxable = salesLegs.reduce((s, l) => s + l.taxable, 0);
    const sumMap = (m) => [...m.values()].reduce((s, v) => s + v, 0);
    const totalGst = sumMap(gstCgst) + sumMap(gstSgst) + sumMap(gstIgst);
    const grandTotal = totalTaxable + totalGst;
    const roAmt = roundOffLedger ? Math.round((Math.round(grandTotal) - grandTotal) * 100) / 100 : 0.0;
    const roTotal = Math.round((grandTotal + roAmt) * 100) / 100;

    const rowVtype = String(rowGet(r, 'VoucherType', '') || '').trim();
    const effectiveVchType = rowVtype || String(voucherType || 'Sales').trim() || 'Sales';
    const vchTypeEsc = xmlEscape(effectiveVchType);
    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a(`    <VOUCHER VCHTYPE="${vchTypeEsc}" ACTION="Create" OBJVIEW="Invoice Voucher View">`);
    a(`     <DATE>${dt}</DATE>`);
    a(`     <VOUCHERTYPENAME>${vchTypeEsc}</VOUCHERTYPENAME>`);
    a(`     <VOUCHERNUMBER>${vno}</VOUCHERNUMBER>`);
    a(`     <PARTYLEDGERNAME>${party}</PARTYLEDGERNAME>`);
    appendInvoicePartyContextXml(a, partyContext, { includeBasicBuyer: true });
    appendCompanyGstContextXml(a, partyContext, companyGstRegistrations);
    a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
    a('     <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>');
    a('     <ISINVOICE>Yes</ISINVOICE>');
    a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>');
    a('     <VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>');
    a('     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>');
    a('     <VCHSTATUSISREACCEPHSNSIXONEDONE>Yes</VCHSTATUSISREACCEPHSNSIXONEDONE>');
    a('     <VCHGSTSTATUSISUNCERTAIN>No</VCHGSTSTATUSISUNCERTAIN>');
    a('     <VCHGSTSTATUSISINCLUDED>Yes</VCHGSTSTATUSISINCLUDED>');
    a('     <VCHGSTSTATUSISAPPLICABLE>Yes</VCHGSTSTATUSISAPPLICABLE>');
    const gstTxnType = gstTransactionType(partyContext.registration_type || '', partyContext.gstin || '');
    a(`     <GSTTRANSACTIONTYPE>${xmlEscape(gstTxnType)}</GSTTRANSACTIONTYPE>`);
    if (invoiceRef) a(`     <REFERENCE>${invoiceRef}</REFERENCE>`);
    if (narr) a(`     <NARRATION>${narr}</NARRATION>`);

    const partyAmt = `-${fmtAmt(roTotal)}`;
    a('     <LEDGERENTRIES.LIST>');
    a(`      <LEDGERNAME>${party}</LEDGERNAME>`);
    a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
    a(`      <AMOUNT>${partyAmt}</AMOUNT>`);
    a('      <BILLALLOCATIONS.LIST>');
    a(`       <NAME>${invoiceRef || vno}</NAME>`);
    a('       <BILLTYPE>New Ref</BILLTYPE>');
    a(`       <AMOUNT>${partyAmt}</AMOUNT>`);
    a('      </BILLALLOCATIONS.LIST>');
    a('     </LEDGERENTRIES.LIST>');

    for (const leg of salesLegs) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(leg.sales_raw)}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(leg.taxable)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }

    for (const [ledName, amt] of gstCgst) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(ledName)}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(amt)}</AMOUNT>`);
      appendTaxObjectAllocationXml(a, 'CGST');
      a('     </LEDGERENTRIES.LIST>');
    }
    for (const [ledName, amt] of gstSgst) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(ledName)}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(amt)}</AMOUNT>`);
      appendTaxObjectAllocationXml(a, 'SGST');
      a('     </LEDGERENTRIES.LIST>');
    }
    for (const [ledName, amt] of gstIgst) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(ledName)}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(amt)}</AMOUNT>`);
      appendTaxObjectAllocationXml(a, 'IGST');
      a('     </LEDGERENTRIES.LIST>');
    }

    appendRoundOffEntry(a, roundOffLedger, roAmt, false);
    a('    </VOUCHER>');
    a('   </TALLYMESSAGE>');
  });

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return lines.join('\n');
}

/** Each row needs ItemName, Quantity, Rate, Per (unit), GodownName (optional).
 * Uses ALLINVENTORYENTRIES.LIST for stock items + LEDGERENTRIES.LIST for
 * accounting legs (party, tax ledgers). */
function generateItemXml(rows, company, opts = {}) {
  const {
    useTodayDate = false, dateMode = '', customTallyDate = '', startVoucherNumber = null,
    fallbackSalesLedger = SUSPENSE_LEDGER, companyGstRegistrations = null,
    voucherType = 'Sales', roundOffLedger = '',
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

  rows = consolidateItemRows(rows, resolvedMode, resolvedCustomDate);

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
    const invoiceRefRaw = rowInvoiceReference(r, vnoRaw);
    const invoiceRef = xmlEscape(invoiceRefRaw);
    const partyRaw = ledgerOrSuspense(rowText(r, 'PartyLedger'));
    const partyContext = collectPartyContext(r, partyRaw);
    const party = xmlEscape(partyRaw);
    const taxable = rowFloat(r, 'TaxableValue', 0.0);

    let invItems, itemTotal;
    const rawItems = r.items || [];
    if (rawItems.length) {
      invItems = rawItems.map((it) => {
        const itNameRaw = String(it.ItemName || '');
        if (!itNameRaw) throw new Error(`Sales item row ${idx + 1}: item name missing in items list.`);
        const itQty = parseFloat(it.Quantity || 0);
        if (itQty <= 0) throw new Error(`Sales item row ${idx + 1}: quantity must be > 0.`);
        const itRate = parseFloat(it.Rate || 0);
        const itPer = xmlEscape(normalizeStockUnitName(String(it.Per || 'Nos')) || 'Nos');
        const itGodown = xmlEscape(String(it.GodownName || 'Main Location'));
        const sLedRaw = ledgerOrSuspense(String(it.SalesLedger || fallbackSalesLedger), fallbackSalesLedger);
        const itSled = xmlEscape(sLedRaw);
        const itAmt = itRate > 0 ? Math.round(itQty * itRate * 100) / 100 : 0.0;
        return {
          item_name: xmlEscape(itNameRaw), qty: itQty,
          rate: itRate > 0 ? itRate : (itQty > 0 ? itAmt / itQty : 0.0),
          per_unit: itPer, godown: itGodown, sled: itSled, amt: itAmt,
        };
      });
      itemTotal = invItems.reduce((s, it) => s + it.amt, 0);
    } else {
      const itemNameRaw = rowText(r, 'ItemName') || rowText(r, 'Item') || rowText(r, 'StockItem')
        || rowText(r, 'ProductName') || rowText(r, 'SalesLedger');
      if (!itemNameRaw) throw new Error(`Item row ${idx + 1}: item name is missing (ItemName/Item/SalesLedger).`);
      const itemName = xmlEscape(itemNameRaw);
      const itemNameKey = nameKey(itemNameRaw);

      const qty = rowFloat(r, 'Quantity', 0.0) || rowFloat(r, 'Qty', 0.0) || rowFloat(r, 'Unit', 0.0);
      if (qty <= 0) throw new Error(`Item row ${idx + 1}: quantity is missing/zero (Quantity/Qty/Unit).`);

      let rate = rowFloat(r, 'Rate', 0.0);
      if (rate <= 0 && taxable > 0 && qty > 0) rate = taxable / qty;
      const perUnitRaw = rowText(r, 'Per', '') || rowText(r, 'UOM', '') || rowText(r, 'Unit', '') || 'Nos';
      const perUnit = xmlEscape(normalizeStockUnitName(perUnitRaw) || 'Nos');
      const godown = xmlEscape(rowText(r, 'GodownName', 'Main Location') || 'Main Location');

      const explicitSalesLedger = rowText(r, 'SalesAccount') || rowText(r, 'Sales Ledger') || rowText(r, 'IncomeLedger');
      let defaultSalesLedger = ledgerOrSuspense(fallbackSalesLedger);
      if (nameKey(defaultSalesLedger) === itemNameKey) {
        for (const candidate of ['Sales Account', 'Sales', 'Sales A/c', 'Sales Ledger']) {
          if (nameKey(candidate) !== itemNameKey) { defaultSalesLedger = candidate; break; }
        }
      }

      let salesLedgerRaw = explicitSalesLedger || rowText(r, 'SalesLedger') || defaultSalesLedger;
      if (nameKey(salesLedgerRaw) === itemNameKey) salesLedgerRaw = defaultSalesLedger;
      salesLedgerRaw = ledgerOrSuspense(salesLedgerRaw, defaultSalesLedger);
      if (nameKey(salesLedgerRaw) === itemNameKey) {
        throw new Error(
          `Item row ${idx + 1}: sales ledger cannot be same as item '${itemNameRaw}'. `
          + 'Provide SalesAccount/IncomeLedger in Excel or use a valid fallback sales ledger.'
        );
      }
      const sales = xmlEscape(salesLedgerRaw);
      const itemAmt = (qty && rate) ? Math.round(qty * rate * 100) / 100 : taxable;
      invItems = [{
        item_name: itemName, qty, rate, per_unit: perUnit,
        godown, sled: sales, amt: itemAmt,
      }];
      itemTotal = itemAmt;
    }

    const cgstR = rowFloat(r, 'CGSTRate', 0.0);
    const cgstAmtExplicit = rowFloat(r, 'CGST Amount', 0.0);
    const cgstLed = xmlEscape(pickTaxLedgerName(r, ['CGSTLedger', 'CGST Ledger', 'CentralTaxLedger', 'Central Tax Ledger', 'Central Tax'], cgstR, 'CGST', cgstAmtExplicit));
    const sgstR = rowFloat(r, 'SGSTRate', 0.0);
    const sgstAmtExplicit = rowFloat(r, 'SGST Amount', 0.0);
    const sgstLed = xmlEscape(pickTaxLedgerName(r, ['SGSTLedger', 'SGST Ledger', 'StateTaxLedger', 'State Tax Ledger', 'State Tax', 'UTGSTLedger', 'UTGST Ledger'], sgstR, 'SGST', sgstAmtExplicit));
    const igstR = rowFloat(r, 'IGSTRate', 0.0);
    const igstAmtExplicit = rowFloat(r, 'IGST Amount', 0.0);
    const igstLed = xmlEscape(pickTaxLedgerName(r, ['IGSTLedger', 'IGST Ledger', 'IntegratedTaxLedger', 'Integrated Tax Ledger', 'Integrated Tax'], igstR, 'IGST', igstAmtExplicit));
    const narr = xmlEscape(rowText(r, 'Narration'));

    const taxableForTax = taxable > 0 ? taxable : itemTotal;
    const cgstAmt = cgstR > 0 ? Math.round(taxableForTax * cgstR / 100 * 100) / 100 : cgstAmtExplicit;
    const sgstAmt = sgstR > 0 ? Math.round(taxableForTax * sgstR / 100 * 100) / 100 : sgstAmtExplicit;
    const igstAmt = igstR > 0 ? Math.round(taxableForTax * igstR / 100 * 100) / 100 : igstAmtExplicit;
    const total = itemTotal + cgstAmt + sgstAmt + igstAmt;
    const roAmtI = roundOffLedger ? Math.round((Math.round(total) - total) * 100) / 100 : 0.0;
    const roTotalI = Math.round((total + roAmtI) * 100) / 100;

    const rowVtypeI = String(rowGet(r, 'VoucherType', '') || '').trim();
    const vchTypeEsc = xmlEscape(rowVtypeI || String(voucherType || 'Sales').trim() || 'Sales');
    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a(`    <VOUCHER VCHTYPE="${vchTypeEsc}" ACTION="Create" OBJVIEW="Invoice Voucher View">`);
    a(`     <DATE>${dt}</DATE>`);
    a(`     <VOUCHERTYPENAME>${vchTypeEsc}</VOUCHERTYPENAME>`);
    a(`     <VOUCHERNUMBER>${vno}</VOUCHERNUMBER>`);
    a(`     <PARTYLEDGERNAME>${party}</PARTYLEDGERNAME>`);
    appendInvoicePartyContextXml(a, partyContext, { includeBasicBuyer: true });
    appendCompanyGstContextXml(a, partyContext, companyGstRegistrations);
    a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
    a('     <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>');
    a('     <ISINVOICE>Yes</ISINVOICE>');
    a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>');
    a('     <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>');
    a('     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>');
    a('     <VCHSTATUSISREACCEPHSNSIXONEDONE>Yes</VCHSTATUSISREACCEPHSNSIXONEDONE>');
    a('     <VCHGSTSTATUSISUNCERTAIN>No</VCHGSTSTATUSISUNCERTAIN>');
    a('     <VCHGSTSTATUSISINCLUDED>Yes</VCHGSTSTATUSISINCLUDED>');
    a('     <VCHGSTSTATUSISAPPLICABLE>Yes</VCHGSTSTATUSISAPPLICABLE>');
    const gstTxnType = gstTransactionType(partyContext.registration_type || '', partyContext.gstin || '');
    a(`     <GSTTRANSACTIONTYPE>${xmlEscape(gstTxnType)}</GSTTRANSACTIONTYPE>`);
    if (invoiceRef) a(`     <REFERENCE>${invoiceRef}</REFERENCE>`);
    if (narr) a(`     <NARRATION>${narr}</NARRATION>`);

    for (const it of invItems) {
      a('     <ALLINVENTORYENTRIES.LIST>');
      a(`      <STOCKITEMNAME>${it.item_name}</STOCKITEMNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <RATE>${fmtAmt(it.rate)}/${it.per_unit}</RATE>`);
      a(`      <AMOUNT>${fmtAmt(it.amt)}</AMOUNT>`);
      a(`      <ACTUALQTY>${fmtAmt(it.qty)} ${it.per_unit}</ACTUALQTY>`);
      a(`      <BILLEDQTY>${fmtAmt(it.qty)} ${it.per_unit}</BILLEDQTY>`);
      a('      <BATCHALLOCATIONS.LIST>');
      a(`       <GODOWNNAME>${it.godown}</GODOWNNAME>`);
      a(`       <AMOUNT>${fmtAmt(it.amt)}</AMOUNT>`);
      a(`       <ACTUALQTY>${fmtAmt(it.qty)} ${it.per_unit}</ACTUALQTY>`);
      a(`       <BILLEDQTY>${fmtAmt(it.qty)} ${it.per_unit}</BILLEDQTY>`);
      a('      </BATCHALLOCATIONS.LIST>');
      a('      <ACCOUNTINGALLOCATIONS.LIST>');
      a(`       <LEDGERNAME>${it.sled}</LEDGERNAME>`);
      a('       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`       <AMOUNT>${fmtAmt(it.amt)}</AMOUNT>`);
      a('      </ACCOUNTINGALLOCATIONS.LIST>');
      a('     </ALLINVENTORYENTRIES.LIST>');
    }

    a('     <LEDGERENTRIES.LIST>');
    a(`      <LEDGERNAME>${party}</LEDGERNAME>`);
    a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
    a(`      <AMOUNT>-${fmtAmt(roTotalI)}</AMOUNT>`);
    a('      <BILLALLOCATIONS.LIST>');
    a(`       <NAME>${invoiceRef || vno}</NAME>`);
    a('       <BILLTYPE>New Ref</BILLTYPE>');
    a(`       <AMOUNT>-${fmtAmt(roTotalI)}</AMOUNT>`);
    a('      </BILLALLOCATIONS.LIST>');
    a('     </LEDGERENTRIES.LIST>');

    if (cgstAmt) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${cgstLed}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(cgstAmt)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }
    if (sgstAmt) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${sgstLed}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(sgstAmt)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }
    if (igstAmt) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${igstLed}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(igstAmt)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }

    appendRoundOffEntry(a, roundOffLedger, roAmtI, false);
    a('    </VOUCHER>');
    a('   </TALLYMESSAGE>');
  });

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return lines.join('\n');
}

module.exports = { generateAccountingXml, generateItemXml };
