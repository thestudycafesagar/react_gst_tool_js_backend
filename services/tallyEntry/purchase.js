/** Purchase Accounting Invoice and Purchase Item Invoice voucher builders.
 * 1:1 port of Backend_Tally/tally_entry/purchase.py. */
const {
  SUSPENSE_LEDGER, fmtAmt, tallyDate, normalizeManualDateToTally, appendRoundOffEntry,
  rowGet, resolveExcelDate, rowText, rowFloat, rowVoucherNumber, rowInvoiceReference,
  ledgerOrSuspense, resolvePartyLedger, companyStaticBlock, normalizeStockUnitName,
  voucherNumberWithOffset, collectPartyContext, appendInvoicePartyContextXml,
  appendCompanyGstContextXml, appendTaxObjectAllocationXml, gstTransactionType,
  pickTaxLedgerName, consolidateAccountingRows, consolidateItemRows, nameKey,
  resolveCompanyGstState, stateNameFromGstin, xmlEscape,
} = require('./common');

/** Purchase accounting invoice XML (mirror of sales with debit/credit reversed). */
function generatePurchaseAccountingXml(rows, company, opts = {}) {
  const {
    useTodayDate = false, dateMode = '', customTallyDate = '', startVoucherNumber = null,
    companyGstRegistrations = null, voucherType = 'Purchase', roundOffLedger = '', useExistingSeries = true,
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
  const purchaseCompanyState = resolveCompanyGstState(companyGstRegistrations || [], '');

  rows = consolidateAccountingRows(rows, resolvedMode, resolvedCustomDate);

  rows.forEach((r, idx) => {
    let sourceDate;
    if (resolvedMode === 'current') sourceDate = new Date();
    else if (resolvedMode === 'custom') sourceDate = resolvedCustomDate;
    else sourceDate = resolveExcelDate(r);
    const dt = tallyDate(sourceDate);

    const supplierInvNoRaw = rowText(r, 'SupplierInvoiceNo') || rowText(r, 'Supplier Invoice No')
      || rowText(r, 'ReferenceNo') || rowText(r, 'InvoiceNo') || rowText(r, 'BillNo');
    const supplierInvDateRaw = rowGet(r, 'SupplierInvoiceDate', '') || rowGet(r, 'Supplier Invoice Date', '') || rowGet(r, 'ReferenceDate', '');
    const supplierInvDate = supplierInvDateRaw ? tallyDate(supplierInvDateRaw) : dt;

    const excelVno = rowVoucherNumber(r, '');
    let vnoRaw;
    if (excelVno) vnoRaw = excelVno;
    else if (startVoucherNumber != null) vnoRaw = voucherNumberWithOffset(startVoucherNumber, idx) || String(startVoucherNumber);
    else vnoRaw = '';
    const supplierInvoiceRaw = supplierInvNoRaw || vnoRaw;
    if (!vnoRaw) vnoRaw = supplierInvoiceRaw;
    const vno = xmlEscape(vnoRaw);
    const supplierInvoice = xmlEscape(supplierInvoiceRaw);

    const partyRaw = ledgerOrSuspense(rowText(r, 'PartyLedger'));
    const partyContext = collectPartyContext(r, partyRaw, false);
    const party = xmlEscape(partyRaw);
    const narr = xmlEscape(rowText(r, 'Narration'));

    const extraLegs = r.acct_legs || [];
    const allLegs = [r, ...extraLegs];

    const purchaseLegs = [];
    const gstCgst = new Map(), gstSgst = new Map(), gstIgst = new Map(), tdsByLed = new Map();

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

      let legPurchaseRaw = rowText(legR, 'PurchaseLedger') || rowText(legR, 'PurchaseAccount')
        || rowText(legR, 'Purchase Ledger') || rowText(legR, 'ExpenseLedger') || rowText(legR, 'SalesLedger');
      legPurchaseRaw = ledgerOrSuspense(legPurchaseRaw);

      const legCgstName = pickTaxLedgerName(legR, ['CGSTLedger', 'CGST Ledger', 'CentralTaxLedger', 'Central Tax Ledger', 'Central Tax'], legCgstR, 'CGST', legCgstExplicit);
      const legSgstName = pickTaxLedgerName(legR, ['SGSTLedger', 'SGST Ledger', 'StateTaxLedger', 'State Tax Ledger', 'State Tax', 'UTGSTLedger', 'UTGST Ledger'], legSgstR, 'SGST', legSgstExplicit);
      const legIgstName = pickTaxLedgerName(legR, ['IGSTLedger', 'IGST Ledger', 'IntegratedTaxLedger', 'Integrated Tax Ledger', 'Integrated Tax'], legIgstR, 'IGST', legIgstExplicit);

      const legTdsLedgerRaw = rowText(legR, 'TDSLedger') || rowText(legR, 'TDS Ledger') || rowText(legR, 'Tds Ledger');
      const legTdsRate = rowFloat(legR, 'TDSRate', 0.0) || rowFloat(legR, 'TDS Rate', 0.0);
      const legTdsRawAmt = rowFloat(legR, 'TDSAmount', 0.0) || rowFloat(legR, 'TDS Amount', 0.0);
      let legTds;
      if (legTdsLedgerRaw && legTdsRawAmt <= 0 && legTdsRate > 0) legTds = Math.round(legTaxable * legTdsRate / 100 * 100) / 100;
      else legTds = Math.abs(legTdsRawAmt);

      purchaseLegs.push({ purchase_raw: legPurchaseRaw, taxable: legTaxable });
      if (legCgst) gstCgst.set(legCgstName, (gstCgst.get(legCgstName) || 0.0) + legCgst);
      if (legSgst) gstSgst.set(legSgstName, (gstSgst.get(legSgstName) || 0.0) + legSgst);
      if (legIgst) gstIgst.set(legIgstName, (gstIgst.get(legIgstName) || 0.0) + legIgst);
      if (legTdsLedgerRaw && legTds > 0) tdsByLed.set(legTdsLedgerRaw, (tdsByLed.get(legTdsLedgerRaw) || 0.0) + legTds);
    }

    const totalTaxable = purchaseLegs.reduce((s, l) => s + l.taxable, 0);
    const sumMap = (m) => [...m.values()].reduce((s, v) => s + v, 0);
    const totalGst = sumMap(gstCgst) + sumMap(gstSgst) + sumMap(gstIgst);
    const totalTds = sumMap(tdsByLed);
    const grandTotal = totalTaxable + totalGst;
    const roAmtPa = roundOffLedger ? Math.round((Math.round(grandTotal) - grandTotal) * 100) / 100 : 0.0;
    const roGrandPa = Math.round((grandTotal + roAmtPa) * 100) / 100;
    const partyTotal = totalTds > 0 ? roGrandPa - totalTds : roGrandPa;

    const vchTypeEsc = xmlEscape(String(voucherType || 'Purchase').trim() || 'Purchase');

    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a(`    <VOUCHER VCHTYPE="${vchTypeEsc}" ACTION="Create" OBJVIEW="Invoice Voucher View">`);
    a(`     <DATE>${dt}</DATE>`);
    a(`     <REFERENCEDATE>${supplierInvDate}</REFERENCEDATE>`);
    a(`     <VOUCHERTYPENAME>${vchTypeEsc}</VOUCHERTYPENAME>`);
    if (!useExistingSeries) a(`     <VOUCHERNUMBER>${vno}</VOUCHERNUMBER>`);
    a(`     <PARTYLEDGERNAME>${party}</PARTYLEDGERNAME>`);
    appendInvoicePartyContextXml(a, partyContext, {
      includeBasicBuyer: false,
      includePlaceOfSupply: Boolean(purchaseCompanyState),
      placeOfSupplyOverride: purchaseCompanyState,
    });
    appendCompanyGstContextXml(a, partyContext, companyGstRegistrations, false);
    a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
    a('     <ISINVOICE>Yes</ISINVOICE>');
    a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>');
    a('     <VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>');
    a('     <ISGSTOVERRIDDEN>Yes</ISGSTOVERRIDDEN>');
    const gstTxnType = gstTransactionType(partyContext.registration_type || '', partyContext.gstin || '');
    a(`     <GSTTRANSACTIONTYPE>${xmlEscape(gstTxnType)}</GSTTRANSACTIONTYPE>`);
    if (supplierInvoice) a(`     <REFERENCE>${supplierInvoice}</REFERENCE>`);
    if (narr) a(`     <NARRATION>${narr}</NARRATION>`);

    a('     <LEDGERENTRIES.LIST>');
    a(`      <LEDGERNAME>${party}</LEDGERNAME>`);
    a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
    a(`      <AMOUNT>${fmtAmt(partyTotal)}</AMOUNT>`);
    a('      <BILLALLOCATIONS.LIST>');
    a(`       <NAME>${supplierInvoice || vno}</NAME>`);
    a('       <BILLTYPE>New Ref</BILLTYPE>');
    a(`       <AMOUNT>${fmtAmt(partyTotal)}</AMOUNT>`);
    a('      </BILLALLOCATIONS.LIST>');
    a('     </LEDGERENTRIES.LIST>');

    for (const leg of purchaseLegs) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(leg.purchase_raw)}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(leg.taxable)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }

    for (const [ledName, amt] of gstCgst) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(ledName)}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(amt)}</AMOUNT>`);
      appendTaxObjectAllocationXml(a, 'CGST');
      a('     </LEDGERENTRIES.LIST>');
    }
    for (const [ledName, amt] of gstSgst) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(ledName)}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(amt)}</AMOUNT>`);
      appendTaxObjectAllocationXml(a, 'SGST');
      a('     </LEDGERENTRIES.LIST>');
    }
    for (const [ledName, amt] of gstIgst) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(ledName)}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(amt)}</AMOUNT>`);
      appendTaxObjectAllocationXml(a, 'IGST');
      a('     </LEDGERENTRIES.LIST>');
    }

    for (const [ledName, amt] of tdsByLed) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(ledName)}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(amt)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }

    appendRoundOffEntry(a, roundOffLedger, roAmtPa, true);
    a('    </VOUCHER>');
    a('   </TALLYMESSAGE>');
  });

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return lines.join('\n');
}

/** Purchase item invoice XML (inventory + accounting allocations). */
function generatePurchaseItemXml(rows, company, opts = {}) {
  const {
    useTodayDate = false, dateMode = '', customTallyDate = '', startVoucherNumber = null,
    fallbackPurchaseLedger = SUSPENSE_LEDGER, companyGstRegistrations = null, companyGstin = '',
    voucherType = 'Purchase', roundOffLedger = '', useExistingSeries = true,
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
  const purchaseCompanyState = resolveCompanyGstState(companyGstRegistrations || [], '');

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
    const supplierInvoiceRaw = rowInvoiceReference(r, vnoRaw);
    if (!vnoRaw) vnoRaw = supplierInvoiceRaw;
    const vno = xmlEscape(vnoRaw);
    const supplierInvoice = xmlEscape(supplierInvoiceRaw);

    const supplierInvDateRaw = rowGet(r, 'SupplierInvoiceDate', '') || rowGet(r, 'Supplier Invoice Date', '') || rowGet(r, 'ReferenceDate', '');
    const supplierInvDate = supplierInvDateRaw ? tallyDate(supplierInvDateRaw) : dt;

    const partyRaw = ledgerOrSuspense(resolvePartyLedger(r, true));
    const partyContext = collectPartyContext(r, partyRaw, false);
    const party = xmlEscape(partyRaw);
    const taxable = rowFloat(r, 'TaxableValue', 0.0);

    let invItems, itemTotal;
    const rawItems = r.items || [];
    if (rawItems.length) {
      invItems = rawItems.map((it) => {
        const itNameRaw = String(it.ItemName || '');
        if (!itNameRaw) throw new Error(`Purchase item row ${idx + 1}: item name missing in items list.`);
        const itQty = parseFloat(it.Quantity || 0);
        if (itQty <= 0) throw new Error(`Purchase item row ${idx + 1}: quantity must be > 0.`);
        const itRate = parseFloat(it.Rate || 0);
        const itPer = xmlEscape(normalizeStockUnitName(String(it.Per || 'Nos')) || 'Nos');
        const itGodown = xmlEscape(String(it.GodownName || 'Main Location'));
        const itPled = xmlEscape(ledgerOrSuspense(String(it.PurchaseLedger || fallbackPurchaseLedger)));
        const itAmt = itRate > 0 ? Math.round(itQty * itRate * 100) / 100 : 0.0;
        return {
          item_name: xmlEscape(itNameRaw), qty: itQty,
          rate: itRate > 0 ? itRate : (itQty > 0 ? itAmt / itQty : 0.0),
          per_unit: itPer, godown: itGodown, pled: itPled, amt: itAmt,
        };
      });
      itemTotal = invItems.reduce((s, it) => s + it.amt, 0);
    } else {
      const itemNameRaw = rowText(r, 'ItemName') || rowText(r, 'Item') || rowText(r, 'StockItem')
        || rowText(r, 'ProductName') || rowText(r, 'PurchaseLedger') || rowText(r, 'SalesLedger');
      if (!itemNameRaw) throw new Error(`Purchase item row ${idx + 1}: item name is missing.`);
      const itemNameKey = nameKey(itemNameRaw);

      const qty = rowFloat(r, 'Quantity', 0.0) || rowFloat(r, 'Qty', 0.0) || rowFloat(r, 'Unit', 0.0);
      if (qty <= 0) throw new Error(`Purchase item row ${idx + 1}: quantity is missing/zero.`);

      let rate = rowFloat(r, 'Rate', 0.0);
      if (rate <= 0 && taxable > 0 && qty > 0) rate = taxable / qty;
      const perUnitRaw = rowText(r, 'Per', '') || rowText(r, 'UOM', '') || rowText(r, 'Unit', '') || 'Nos';
      const perUnit = xmlEscape(normalizeStockUnitName(perUnitRaw) || 'Nos');
      const godown = xmlEscape(rowText(r, 'GodownName', 'Main Location') || 'Main Location');

      const explicitPurchaseLedger = rowText(r, 'PurchaseAccount') || rowText(r, 'Purchase Ledger')
        || rowText(r, 'ExpenseLedger') || rowText(r, 'PurchaseLedger');
      let defaultPurchaseLedger = ledgerOrSuspense(fallbackPurchaseLedger);
      if (nameKey(defaultPurchaseLedger) === itemNameKey) {
        for (const candidate of ['Purchase Account', 'Purchase', 'Purchase A/c', 'Purchase Ledger']) {
          if (nameKey(candidate) !== itemNameKey) { defaultPurchaseLedger = candidate; break; }
        }
      }

      let purchaseLedgerRaw = explicitPurchaseLedger || rowText(r, 'PurchaseLedger') || rowText(r, 'SalesLedger') || defaultPurchaseLedger;
      if (nameKey(purchaseLedgerRaw) === itemNameKey) purchaseLedgerRaw = defaultPurchaseLedger;
      purchaseLedgerRaw = ledgerOrSuspense(purchaseLedgerRaw, defaultPurchaseLedger);
      if (nameKey(purchaseLedgerRaw) === itemNameKey) {
        throw new Error(`Purchase item row ${idx + 1}: purchase ledger cannot match item '${itemNameRaw}'.`);
      }
      const sAmt = (qty && rate) ? Math.round(qty * rate * 100) / 100 : taxable;
      invItems = [{
        item_name: xmlEscape(itemNameRaw), qty, rate, per_unit: perUnit,
        godown, pled: xmlEscape(purchaseLedgerRaw), amt: sAmt,
      }];
      itemTotal = sAmt;
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

    const tdsLedgerRaw = rowText(r, 'TDSLedger') || rowText(r, 'TDS Ledger') || rowText(r, 'Tds Ledger');
    const tdsRate = rowFloat(r, 'TDSRate', 0.0) || rowFloat(r, 'TDS Rate', 0.0);
    const tdsAmountRaw = rowFloat(r, 'TDSAmount', 0.0) || rowFloat(r, 'TDS Amount', 0.0);
    let tdsAmount;
    if (tdsLedgerRaw && tdsAmountRaw <= 0 && tdsRate > 0) tdsAmount = Math.round(taxableForTax * tdsRate / 100 * 100) / 100;
    else tdsAmount = Math.abs(tdsAmountRaw);
    const tdsLed = xmlEscape(tdsLedgerRaw);
    const tdsVal = (tdsLed && tdsAmount > 0) ? tdsAmount : 0.0;
    const netPreRo = Math.round((total - tdsVal) * 100) / 100;
    const roAmtPi = roundOffLedger ? Math.round((Math.round(netPreRo) - netPreRo) * 100) / 100 : 0.0;
    const partyTotal = Math.round((netPreRo + roAmtPi) * 100) / 100;

    const vchTypeEsc = xmlEscape(String(voucherType || 'Purchase').trim() || 'Purchase');

    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a(`    <VOUCHER VCHTYPE="${vchTypeEsc}" ACTION="Create" OBJVIEW="Invoice Voucher View">`);
    a(`     <DATE>${dt}</DATE>`);
    a(`     <REFERENCEDATE>${supplierInvDate}</REFERENCEDATE>`);
    a(`     <VOUCHERTYPENAME>${vchTypeEsc}</VOUCHERTYPENAME>`);
    if (!useExistingSeries) a(`     <VOUCHERNUMBER>${vno}</VOUCHERNUMBER>`);
    a(`     <PARTYLEDGERNAME>${party}</PARTYLEDGERNAME>`);
    appendInvoicePartyContextXml(a, partyContext, {
      includeBasicBuyer: false,
      includePlaceOfSupply: Boolean(purchaseCompanyState),
      placeOfSupplyOverride: purchaseCompanyState,
    });
    appendCompanyGstContextXml(a, partyContext, companyGstRegistrations, false);
    if (!(companyGstRegistrations || []).length && companyGstin) {
      const coGstin = xmlEscape(companyGstin.trim().toUpperCase());
      const coStateName = stateNameFromGstin(companyGstin);
      const coRegName = xmlEscape(coStateName ? `${coStateName} Registration` : companyGstin.trim().toUpperCase());
      a(`     <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="${coGstin}">${coRegName}</GSTREGISTRATION>`);
      a(`     <CMPGSTIN>${coGstin}</CMPGSTIN>`);
      a('     <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>');
      if (coStateName) a(`     <CMPGSTSTATE>${xmlEscape(coStateName)}</CMPGSTSTATE>`);
    }
    a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
    a('     <ISINVOICE>Yes</ISINVOICE>');
    a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>');
    a('     <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>');
    a('     <ISGSTOVERRIDDEN>Yes</ISGSTOVERRIDDEN>');
    const gstTxnType = gstTransactionType(partyContext.registration_type || '', partyContext.gstin || '');
    a(`     <GSTTRANSACTIONTYPE>${xmlEscape(gstTxnType)}</GSTTRANSACTIONTYPE>`);
    if (supplierInvoice) a(`     <REFERENCE>${supplierInvoice}</REFERENCE>`);
    if (narr) a(`     <NARRATION>${narr}</NARRATION>`);

    for (const it of invItems) {
      a('     <ALLINVENTORYENTRIES.LIST>');
      a(`      <STOCKITEMNAME>${it.item_name}</STOCKITEMNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <RATE>${fmtAmt(it.rate)}/${it.per_unit}</RATE>`);
      a(`      <AMOUNT>-${fmtAmt(it.amt)}</AMOUNT>`);
      a(`      <ACTUALQTY>${fmtAmt(it.qty)} ${it.per_unit}</ACTUALQTY>`);
      a(`      <BILLEDQTY>${fmtAmt(it.qty)} ${it.per_unit}</BILLEDQTY>`);
      a('      <BATCHALLOCATIONS.LIST>');
      a(`       <GODOWNNAME>${it.godown}</GODOWNNAME>`);
      a(`       <AMOUNT>-${fmtAmt(it.amt)}</AMOUNT>`);
      a(`       <ACTUALQTY>${fmtAmt(it.qty)} ${it.per_unit}</ACTUALQTY>`);
      a(`       <BILLEDQTY>${fmtAmt(it.qty)} ${it.per_unit}</BILLEDQTY>`);
      a('      </BATCHALLOCATIONS.LIST>');
      a('      <ACCOUNTINGALLOCATIONS.LIST>');
      a(`       <LEDGERNAME>${it.pled}</LEDGERNAME>`);
      a('       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`       <AMOUNT>-${fmtAmt(it.amt)}</AMOUNT>`);
      a('      </ACCOUNTINGALLOCATIONS.LIST>');
      a('     </ALLINVENTORYENTRIES.LIST>');
    }

    a('     <LEDGERENTRIES.LIST>');
    a(`      <LEDGERNAME>${party}</LEDGERNAME>`);
    a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
    a(`      <AMOUNT>${fmtAmt(partyTotal)}</AMOUNT>`);
    a('      <BILLALLOCATIONS.LIST>');
    a(`       <NAME>${supplierInvoice || vno}</NAME>`);
    a('       <BILLTYPE>New Ref</BILLTYPE>');
    a(`       <AMOUNT>${fmtAmt(partyTotal)}</AMOUNT>`);
    a('      </BILLALLOCATIONS.LIST>');
    a('     </LEDGERENTRIES.LIST>');

    if (cgstAmt) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${cgstLed}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(cgstAmt)}</AMOUNT>`);
      appendTaxObjectAllocationXml(a, 'CGST');
      a('     </LEDGERENTRIES.LIST>');
    }
    if (sgstAmt) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${sgstLed}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(sgstAmt)}</AMOUNT>`);
      appendTaxObjectAllocationXml(a, 'SGST');
      a('     </LEDGERENTRIES.LIST>');
    }
    if (igstAmt) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${igstLed}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(igstAmt)}</AMOUNT>`);
      appendTaxObjectAllocationXml(a, 'IGST');
      a('     </LEDGERENTRIES.LIST>');
    }

    if (tdsLed && tdsAmount > 0) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${tdsLed}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(tdsAmount)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }

    appendRoundOffEntry(a, roundOffLedger, roAmtPi, true);
    a('    </VOUCHER>');
    a('   </TALLYMESSAGE>');
  });

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return lines.join('\n');
}

module.exports = { generatePurchaseAccountingXml, generatePurchaseItemXml };
