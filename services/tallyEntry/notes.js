/** Credit Note / Debit Note voucher builder (one shared panel in the
 * desktop tool — toggled between accounting/item entry mode, never between
 * sales/purchase since the same voucher type applies either way).
 * 1:1 port of Backend_Tally/tally_entry/notes.py. */
const {
  fmtAmt, tallyDate, normalizeManualDateToTally, appendRoundOffEntry,
  rowText, rowTextAny, rowFloat, rowVoucherNumber, ledgerOrSuspense,
  normalizeStockUnitName, cleanTaxLedger, resolveExcelDate, stateNameFromGstin,
  companyStaticBlock, xmlEscape,
} = require('./common');

function normalizeNoteType(value) {
  const text = String(value || '').replace(/\([^)]*\)/g, '').trim().toLowerCase();
  if (['debit note', 'debit', 'debitnote'].includes(text)) return 'Debit Note';
  return 'Credit Note';
}

/** Merge Credit/Debit Note rows with the same (date, GSTIN, voucher no,
 * party ledger) into one voucher; accumulated tax amounts stored as
 * _note_*_amt, item/accounting legs stored as items/accounting_entries. */
function consolidateNoteRows(rows, resolvedMode, resolvedCustomDate) {
  const consolidated = [];
  const groupMap = new Map();

  for (const r of rows) {
    const taxable = rowFloat(r, 'TaxableValue', 0.0);
    if (taxable <= 0) continue;

    let sourceDate;
    if (resolvedMode === 'current') sourceDate = new Date();
    else if (resolvedMode === 'custom') sourceDate = resolvedCustomDate;
    else sourceDate = resolveExcelDate(r);
    const dt = tallyDate(sourceDate);

    const vnoRaw = rowVoucherNumber(r, '');
    const gstinRaw = (rowText(r, 'GSTIN/UIN') || rowText(r, 'GSTIN') || rowText(r, 'PartyGSTIN')).toUpperCase();
    const partyRaw = String(rowText(r, 'PartyLedger') || rowText(r, 'Party Ledger')).trim().toLowerCase();
    const key = `${dt}|${gstinRaw.toLowerCase()}|${String(vnoRaw).trim().toLowerCase()}|${partyRaw}`;

    const cgstRate = rowFloat(r, 'CGSTRate', 0.0);
    const sgstRate = rowFloat(r, 'SGSTRate', 0.0);
    const igstRate = rowFloat(r, 'IGSTRate', 0.0);
    const cgstX = rowFloat(r, 'CGST Amount', 0.0);
    const sgstX = rowFloat(r, 'SGST Amount', 0.0);
    const igstX = rowFloat(r, 'IGST Amount', 0.0);
    const rowCgst = cgstRate > 0 ? Math.round(taxable * cgstRate / 100 * 100) / 100 : cgstX;
    const rowSgst = sgstRate > 0 ? Math.round(taxable * sgstRate / 100 * 100) / 100 : sgstX;
    const rowIgst = igstRate > 0 ? Math.round(taxable * igstRate / 100 * 100) / 100 : igstX;

    const itemName = rowTextAny(r, ['Item Name', 'ItemName', 'Item', 'StockItem', 'ProductName'], '');
    const particular = rowText(r, 'Particular') || rowText(r, 'Particulars') || '';

    if (groupMap.has(key)) {
      const base = groupMap.get(key);
      base.TaxableValue = rowFloat(base, 'TaxableValue', 0.0) + taxable;
      base._note_cgst_amt = (base._note_cgst_amt || 0.0) + rowCgst;
      base._note_sgst_amt = (base._note_sgst_amt || 0.0) + rowSgst;
      base._note_igst_amt = (base._note_igst_amt || 0.0) + rowIgst;
      base.CGSTRate = 0;
      base.SGSTRate = 0;
      base.IGSTRate = 0;

      if (!base.items) {
        base.items = [];
        const origTv = base._note_first_taxable !== undefined ? base._note_first_taxable : taxable;
        const origItem = {
          ItemName: rowTextAny(base, ['Item Name', 'ItemName', 'Item', 'StockItem', 'ProductName'], ''),
          Quantity: rowFloat(base, 'Quantity', 0.0) || rowFloat(base, 'Qty', 0.0) || 1.0,
          Rate: rowFloat(base, 'Rate', 0.0),
          Per: rowTextAny(base, ['Unit', 'UOM', 'Per'], '') || 'Nos',
          GodownName: rowText(base, 'GodownName') || 'Main Location',
          Particular: rowText(base, 'Particular') || rowText(base, 'Particulars') || '',
          TaxableValue: origTv,
        };
        if (origItem.ItemName) base.items.push(origItem);
        base.accounting_entries = [{ Particular: origItem.Particular, TaxableValue: origTv }];
      }

      const newItem = {
        ItemName: itemName,
        Quantity: rowFloat(r, 'Quantity', 0.0) || rowFloat(r, 'Qty', 0.0) || 1.0,
        Rate: rowFloat(r, 'Rate', 0.0),
        Per: rowTextAny(r, ['Unit', 'UOM', 'Per'], '') || 'Nos',
        GodownName: rowText(r, 'GodownName') || 'Main Location',
        Particular: particular,
        TaxableValue: taxable,
      };
      if (itemName) base.items.push(newItem);
      if (base.accounting_entries) base.accounting_entries.push({ Particular: particular, TaxableValue: taxable });
    } else {
      const newR = { ...r };
      newR._note_first_taxable = taxable;
      newR._note_cgst_amt = rowCgst;
      newR._note_sgst_amt = rowSgst;
      newR._note_igst_amt = rowIgst;
      groupMap.set(key, newR);
      consolidated.push(newR);
    }
  }

  return consolidated;
}

/** Credit/Debit Note, accounting or item mode.
 * Credit Note: party credited (ISDEEMEDPOSITIVE=No), particular/tax debited (Yes).
 * Debit Note:  party debited  (ISDEEMEDPOSITIVE=Yes), particular/tax credited (No).
 * Returns [xml, voucherCount] — mirrors the Python tuple return. */
function generateNoteXml(rows, company, opts = {}) {
  const {
    useTodayDate = false, dateMode = '', customTallyDate = '', voucherType = 'Credit Note',
    companyGstRegistrations = null, entryMode = 'accounting', roundOffLedger = '',
  } = opts;
  const normalizedType = normalizeNoteType(voucherType);
  const isDebitNote = normalizedType === 'Debit Note';
  const defaultParticularLedger = `${normalizedType} Account`;
  let resolvedEntryMode = String(entryMode || 'accounting').trim().toLowerCase();
  if (!['accounting', 'item'].includes(resolvedEntryMode)) resolvedEntryMode = 'accounting';
  const isItemMode = resolvedEntryMode === 'item';

  const linesOut = [];
  const a = (s) => linesOut.push(s);
  const companyStatic = companyStaticBlock(company);
  a('<?xml version="1.0" encoding="UTF-8"?>');
  a('<ENVELOPE>');
  a(' <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>');
  a(' <BODY><IMPORTDATA>');
  a('  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>');
  if (companyStatic) a(companyStatic);
  a('  </REQUESTDESC>');
  a('  <REQUESTDATA>');

  let voucherCount = 0;
  let resolvedMode = String(dateMode || (useTodayDate ? 'current' : 'excel')).trim().toLowerCase();
  if (!['current', 'excel', 'custom'].includes(resolvedMode)) resolvedMode = useTodayDate ? 'current' : 'excel';
  const resolvedCustomDate = resolvedMode === 'custom' ? normalizeManualDateToTally(customTallyDate) : '';

  rows = consolidateNoteRows(rows, resolvedMode, resolvedCustomDate);

  rows.forEach((r, idx) => {
    const taxable = rowFloat(r, 'TaxableValue', 0.0);
    if (taxable <= 0) return;

    let sourceDate;
    if (resolvedMode === 'current') sourceDate = new Date();
    else if (resolvedMode === 'custom') sourceDate = resolvedCustomDate;
    else sourceDate = resolveExcelDate(r);
    const dt = tallyDate(sourceDate);

    const vnoRaw = rowVoucherNumber(r, String(idx + 1));
    const partyRaw = ledgerOrSuspense(rowText(r, 'PartyLedger') || rowText(r, 'Party Ledger'));
    let particularRaw = rowText(r, 'Particular') || rowText(r, 'Particulars')
      || rowText(r, 'SalesLedger') || rowText(r, 'Sales Ledger')
      || rowText(r, 'Purchase Ledger') || rowText(r, 'PurchaseLedger') || defaultParticularLedger;
    particularRaw = ledgerOrSuspense(particularRaw) || defaultParticularLedger;

    let cgstLedgerRaw = cleanTaxLedger(rowText(r, 'CGSTLedger'));
    let sgstLedgerRaw = cleanTaxLedger(rowText(r, 'SGSTLedger'));
    let igstLedgerRaw = cleanTaxLedger(rowText(r, 'IGSTLedger'));

    const cgstRate = rowFloat(r, 'CGSTRate', 0.0);
    const sgstRate = rowFloat(r, 'SGSTRate', 0.0);
    const igstRate = rowFloat(r, 'IGSTRate', 0.0);

    let cgstAmt, sgstAmt, igstAmt;
    if (r._note_cgst_amt !== undefined) {
      cgstAmt = r._note_cgst_amt;
      sgstAmt = r._note_sgst_amt;
      igstAmt = r._note_igst_amt;
    } else {
      const cgstAmtExplicit = rowFloat(r, 'CGST Amount', 0.0);
      const sgstAmtExplicit = rowFloat(r, 'SGST Amount', 0.0);
      const igstAmtExplicit = rowFloat(r, 'IGST Amount', 0.0);
      cgstAmt = cgstRate > 0 ? Math.round(taxable * cgstRate / 100 * 100) / 100 : cgstAmtExplicit;
      sgstAmt = sgstRate > 0 ? Math.round(taxable * sgstRate / 100 * 100) / 100 : sgstAmtExplicit;
      igstAmt = igstRate > 0 ? Math.round(taxable * igstRate / 100 * 100) / 100 : igstAmtExplicit;
    }

    if (!cgstLedgerRaw && cgstAmt > 0) cgstLedgerRaw = 'CGST';
    if (!sgstLedgerRaw && sgstAmt > 0) sgstLedgerRaw = 'SGST';
    if (!igstLedgerRaw && igstAmt > 0) igstLedgerRaw = 'IGST';

    cgstAmt = cgstLedgerRaw ? cgstAmt : 0.0;
    sgstAmt = sgstLedgerRaw ? sgstAmt : 0.0;
    igstAmt = igstLedgerRaw ? igstAmt : 0.0;

    const total = taxable + cgstAmt + sgstAmt + igstAmt;
    const roAmtN = roundOffLedger ? Math.round((Math.round(total) - total) * 100) / 100 : 0.0;
    const roTotalN = Math.round((total + roAmtN) * 100) / 100;

    const vno = xmlEscape(vnoRaw);
    const party = xmlEscape(partyRaw);
    const particular = xmlEscape(particularRaw);
    const narration = xmlEscape(rowText(r, 'Narration'));
    const gstinRaw = (rowText(r, 'GSTIN/UIN') || rowText(r, 'GSTIN') || rowText(r, 'PartyGSTIN')).toUpperCase();
    const gstin = xmlEscape(gstinRaw);

    let itemName = '', rate = '', perUnit = '', godown = '';
    let qty = 0.0, itemAmt = 0.0;
    if (isItemMode) {
      const itemNameRaw = rowTextAny(r, ['Item Name', 'ItemName', 'Item', 'StockItem', 'ProductName'], '');
      if (!itemNameRaw) throw new Error(`${normalizedType} row ${idx + 1}: item name is missing.`);
      qty = rowFloat(r, 'Quantity', 0.0) || rowFloat(r, 'Qty', 0.0) || 1.0;
      rate = rowFloat(r, 'Rate', 0.0);
      if (rate <= 0 && taxable > 0 && qty > 0) rate = taxable / qty;
      const perUnitRaw = rowTextAny(r, ['Unit', 'UOM', 'Per'], '') || 'Nos';
      perUnit = xmlEscape(normalizeStockUnitName(perUnitRaw) || 'Nos');
      godown = xmlEscape(rowText(r, 'GodownName', 'Main Location') || 'Main Location');
      itemName = xmlEscape(itemNameRaw);
      itemAmt = (qty && rate) ? Math.round(qty * rate * 100) / 100 : taxable;
    }

    const placeRaw = rowText(r, 'PlaceOfSupply') || rowText(r, 'Place Of Supply') || stateNameFromGstin(gstinRaw);
    const stateXml = xmlEscape(placeRaw);

    const partyIsDeemedPositive = isDebitNote ? 'Yes' : 'No';
    const partyAmount = isDebitNote ? -roTotalN : roTotalN;
    const counterIsDeemedPositive = isDebitNote ? 'No' : 'Yes';
    const taxableAmount = isDebitNote ? taxable : -taxable;
    const cgstAmount = isDebitNote ? cgstAmt : -cgstAmt;
    const sgstAmount = isDebitNote ? sgstAmt : -sgstAmt;
    const igstAmount = isDebitNote ? igstAmt : -igstAmt;

    voucherCount += 1;
    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a(`    <VOUCHER VCHTYPE="${normalizedType}" ACTION="Create" OBJVIEW="Invoice Voucher View">`);
    a(`     <DATE>${dt}</DATE>`);
    a(`     <VOUCHERTYPENAME>${normalizedType}</VOUCHERTYPENAME>`);
    a(`     <VOUCHERNUMBER>${vno}</VOUCHERNUMBER>`);
    a(`     <PARTYLEDGERNAME>${party}</PARTYLEDGERNAME>`);
    a(`     <PARTYNAME>${party}</PARTYNAME>`);

    if (!isDebitNote) {
      a(`     <BASICBUYERNAME>${party}</BASICBUYERNAME>`);
      if (stateXml) {
        a(`     <STATENAME>${stateXml}</STATENAME>`);
        a(`     <PLACEOFSUPPLY>${stateXml}</PLACEOFSUPPLY>`);
      }
    } else {
      const cmpRegs = companyGstRegistrations || [];
      let cmpState = '', cmpGstin = '', cmpName = '';
      if (cmpRegs.length) {
        const cr = cmpRegs[0];
        cmpGstin = xmlEscape(String(cr.gstin || '').trim());
        cmpState = xmlEscape(String(cr.state || '').trim());
        cmpName = xmlEscape(String(cr.name || '').trim());
      }
      if (stateXml) a(`     <STATENAME>${stateXml}</STATENAME>`);
      if (cmpState) a(`     <PLACEOFSUPPLY>${cmpState}</PLACEOFSUPPLY>`);
      const dnReg = gstinRaw ? 'Regular' : 'Unregistered';
      a(`     <GSTREGISTRATIONTYPE>${dnReg}</GSTREGISTRATIONTYPE>`);
      if (gstinRaw) a('     <VATDEALERTYPE>Regular</VATDEALERTYPE>');
      if (cmpGstin && cmpName) {
        a(`     <GSTREGISTRATION TAXTYPE="GST" TAXREGISTRATION="${cmpGstin}">${cmpName}</GSTREGISTRATION>`);
        a(`     <CMPGSTIN>${cmpGstin}</CMPGSTIN>`);
        a('     <CMPGSTREGISTRATIONTYPE>Regular</CMPGSTREGISTRATIONTYPE>');
      }
      if (cmpState) a(`     <CMPGSTSTATE>${cmpState}</CMPGSTSTATE>`);
    }

    a('     <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>');
    a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
    a('     <ISINVOICE>Yes</ISINVOICE>');
    a('     <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>');
    a(`     <VCHENTRYMODE>${isItemMode ? 'Item Invoice' : 'Accounting Invoice'}</VCHENTRYMODE>`);
    a('     <NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>');
    a('     <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>');
    if (isDebitNote) {
      const dnGstTxn = gstinRaw ? 'Tax Invoice' : 'Unregistered';
      a(`     <GSTTRANSACTIONTYPE>${dnGstTxn}</GSTTRANSACTIONTYPE>`);
    }
    if (gstin) a(`     <PARTYGSTIN>${gstin}</PARTYGSTIN>`);
    if (narration) a(`     <NARRATION>${narration}</NARRATION>`);

    a('     <LEDGERENTRIES.LIST>');
    a(`      <LEDGERNAME>${party}</LEDGERNAME>`);
    a(`      <ISDEEMEDPOSITIVE>${partyIsDeemedPositive}</ISDEEMEDPOSITIVE>`);
    a(`      <AMOUNT>${fmtAmt(partyAmount)}</AMOUNT>`);
    a('     </LEDGERENTRIES.LIST>');

    if (isItemMode) {
      const invIsDeemed = counterIsDeemedPositive;
      const mergedItems = r.items;
      if (mergedItems && mergedItems.length) {
        for (const it of mergedItems) {
          const itName = xmlEscape(String(it.ItemName || '').trim());
          if (!itName) continue;
          const itQty = parseFloat(it.Quantity || 1.0);
          const itRate = parseFloat(it.Rate || 0.0);
          const itPer = xmlEscape(normalizeStockUnitName(String(it.Per || 'Nos')) || 'Nos');
          const itGodown = xmlEscape(String(it.GodownName || 'Main Location'));
          const itParticular = xmlEscape(String(it.Particular || particularRaw) || defaultParticularLedger);
          const itTv = parseFloat(it.TaxableValue || 0.0);
          const itAmtBase = (itQty && itRate) ? Math.round(itQty * itRate * 100) / 100 : itTv;
          const itInvAmount = invIsDeemed === 'Yes' ? -itAmtBase : itAmtBase;
          a('     <ALLINVENTORYENTRIES.LIST>');
          a(`      <STOCKITEMNAME>${itName}</STOCKITEMNAME>`);
          a(`      <ISDEEMEDPOSITIVE>${invIsDeemed}</ISDEEMEDPOSITIVE>`);
          a(`      <RATE>${fmtAmt(itRate)}/${itPer}</RATE>`);
          a(`      <AMOUNT>${fmtAmt(itInvAmount)}</AMOUNT>`);
          a(`      <ACTUALQTY>${fmtAmt(itQty)} ${itPer}</ACTUALQTY>`);
          a(`      <BILLEDQTY>${fmtAmt(itQty)} ${itPer}</BILLEDQTY>`);
          a('      <BATCHALLOCATIONS.LIST>');
          a(`       <GODOWNNAME>${itGodown}</GODOWNNAME>`);
          a(`       <AMOUNT>${fmtAmt(itInvAmount)}</AMOUNT>`);
          a(`       <ACTUALQTY>${fmtAmt(itQty)} ${itPer}</ACTUALQTY>`);
          a(`       <BILLEDQTY>${fmtAmt(itQty)} ${itPer}</BILLEDQTY>`);
          a('      </BATCHALLOCATIONS.LIST>');
          a('      <ACCOUNTINGALLOCATIONS.LIST>');
          a(`       <LEDGERNAME>${itParticular}</LEDGERNAME>`);
          a(`       <ISDEEMEDPOSITIVE>${invIsDeemed}</ISDEEMEDPOSITIVE>`);
          a(`       <AMOUNT>${fmtAmt(itInvAmount)}</AMOUNT>`);
          a('      </ACCOUNTINGALLOCATIONS.LIST>');
          a('     </ALLINVENTORYENTRIES.LIST>');
        }
      } else {
        const invAmount = invIsDeemed === 'Yes' ? -itemAmt : itemAmt;
        a('     <ALLINVENTORYENTRIES.LIST>');
        a(`      <STOCKITEMNAME>${itemName}</STOCKITEMNAME>`);
        a(`      <ISDEEMEDPOSITIVE>${invIsDeemed}</ISDEEMEDPOSITIVE>`);
        a(`      <RATE>${fmtAmt(rate)}/${perUnit}</RATE>`);
        a(`      <AMOUNT>${fmtAmt(invAmount)}</AMOUNT>`);
        a(`      <ACTUALQTY>${fmtAmt(qty)} ${perUnit}</ACTUALQTY>`);
        a(`      <BILLEDQTY>${fmtAmt(qty)} ${perUnit}</BILLEDQTY>`);
        a('      <BATCHALLOCATIONS.LIST>');
        a(`       <GODOWNNAME>${godown}</GODOWNNAME>`);
        a(`       <AMOUNT>${fmtAmt(invAmount)}</AMOUNT>`);
        a(`       <ACTUALQTY>${fmtAmt(qty)} ${perUnit}</ACTUALQTY>`);
        a(`       <BILLEDQTY>${fmtAmt(qty)} ${perUnit}</BILLEDQTY>`);
        a('      </BATCHALLOCATIONS.LIST>');
        a('      <ACCOUNTINGALLOCATIONS.LIST>');
        a(`       <LEDGERNAME>${particular}</LEDGERNAME>`);
        a(`       <ISDEEMEDPOSITIVE>${invIsDeemed}</ISDEEMEDPOSITIVE>`);
        a(`       <AMOUNT>${fmtAmt(invAmount)}</AMOUNT>`);
        a('      </ACCOUNTINGALLOCATIONS.LIST>');
        a('     </ALLINVENTORYENTRIES.LIST>');
      }
    } else {
      const acctEntries = r.accounting_entries;
      if (acctEntries && acctEntries.length > 1) {
        for (const ae of acctEntries) {
          const aeParticular = xmlEscape(String(ae.Particular || particularRaw) || defaultParticularLedger);
          const aeTv = parseFloat(ae.TaxableValue || 0.0);
          const aeAmount = isDebitNote ? aeTv : -aeTv;
          a('     <LEDGERENTRIES.LIST>');
          a(`      <LEDGERNAME>${aeParticular}</LEDGERNAME>`);
          a(`      <ISDEEMEDPOSITIVE>${counterIsDeemedPositive}</ISDEEMEDPOSITIVE>`);
          a(`      <AMOUNT>${fmtAmt(aeAmount)}</AMOUNT>`);
          a('     </LEDGERENTRIES.LIST>');
        }
      } else {
        a('     <LEDGERENTRIES.LIST>');
        a(`      <LEDGERNAME>${particular}</LEDGERNAME>`);
        a(`      <ISDEEMEDPOSITIVE>${counterIsDeemedPositive}</ISDEEMEDPOSITIVE>`);
        a(`      <AMOUNT>${fmtAmt(taxableAmount)}</AMOUNT>`);
        a('     </LEDGERENTRIES.LIST>');
      }
    }

    if (cgstAmt > 0 && cgstLedgerRaw) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(cgstLedgerRaw)}</LEDGERNAME>`);
      a(`      <ISDEEMEDPOSITIVE>${counterIsDeemedPositive}</ISDEEMEDPOSITIVE>`);
      a(`      <AMOUNT>${fmtAmt(cgstAmount)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }
    if (sgstAmt > 0 && sgstLedgerRaw) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(sgstLedgerRaw)}</LEDGERNAME>`);
      a(`      <ISDEEMEDPOSITIVE>${counterIsDeemedPositive}</ISDEEMEDPOSITIVE>`);
      a(`      <AMOUNT>${fmtAmt(sgstAmount)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }
    if (igstAmt > 0 && igstLedgerRaw) {
      a('     <LEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${xmlEscape(igstLedgerRaw)}</LEDGERNAME>`);
      a(`      <ISDEEMEDPOSITIVE>${counterIsDeemedPositive}</ISDEEMEDPOSITIVE>`);
      a(`      <AMOUNT>${fmtAmt(igstAmount)}</AMOUNT>`);
      a('     </LEDGERENTRIES.LIST>');
    }

    appendRoundOffEntry(a, roundOffLedger, roAmtN, isDebitNote);

    a('    </VOUCHER>');
    a('   </TALLYMESSAGE>');
  });

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return [linesOut.join('\n'), voucherCount];
}

module.exports = { generateNoteXml, normalizeNoteType, consolidateNoteRows };
