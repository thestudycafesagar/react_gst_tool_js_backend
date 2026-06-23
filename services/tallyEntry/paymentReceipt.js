/** Payment / Receipt (and auto-detected Contra) single-voucher-per-row builder.
 * 1:1 port of Backend_Tally/tally_entry/payment_receipt.py. */
const { fmtAmt, tallyDate, normalizeManualDateToTally, companyStaticBlock, resolveExcelDate, rowText, rowFloat, xmlEscape } = require('./common');

const PAYMENT_HEADERS = ['DATE', 'DESCRIPTION', 'CHEQUE NO.', 'Amount', 'LEDGER'];
const RECEIPT_HEADERS = ['DATE', 'DESCRIPTION', 'CHEQUE NO.', 'Amount', 'LEDGER'];

function titleCase(s) {
  return String(s).trim().replace(/\w\S*/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase());
}

/** Payment, Receipt, or Contra voucher XML from rows with an AMOUNT column.
 * voucher_type="Payment": dest-ledger DEBIT + bank CREDIT (money going out).
 * voucher_type="Receipt": bank DEBIT + src-ledger CREDIT (money coming in).
 * Auto-switches to "Contra" when the row's LEDGER is a known bank/cash account.
 * Returns [xml, voucherCount] — mirrors the Python tuple return. */
function generateSingleVoucherXml(rows, company, bankLedger, voucherType, opts = {}) {
  const { dateMode = 'excel', customTallyDate = '', startVno = null, contraLedgerNames = null } = opts;
  const vtype = titleCase(voucherType);
  const contraSet = new Set((contraLedgerNames || []).filter(Boolean).map((n) => n.trim().toUpperCase()));
  let resolvedMode = String(dateMode || 'excel').trim().toLowerCase();
  if (!['current', 'excel', 'custom'].includes(resolvedMode)) resolvedMode = 'excel';
  const resolvedCustom = resolvedMode === 'custom' ? normalizeManualDateToTally(customTallyDate) : '';

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

  const bankEsc = xmlEscape(String(bankLedger || '').trim());
  let vchCounter = 0;

  for (const r of rows) {
    let dt;
    let sourceDate = null;
    if (resolvedMode === 'current') {
      const today = new Date();
      dt = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
    } else if (resolvedMode === 'custom') {
      dt = resolvedCustom;
    } else {
      sourceDate = resolveExcelDate(r);
      dt = tallyDate(sourceDate);
      if (!sourceDate) continue;
    }

    const amt = rowFloat(r, 'AMOUNT', 0.0) || rowFloat(r, 'Amount', 0.0);
    if (amt <= 0) continue;

    const contraRaw = (rowText(r, 'LEDGER') || rowText(r, 'Ledger')).trim() || 'Suspense A/c';
    const contraEsc = xmlEscape(contraRaw);
    const isContraRow = contraSet.size > 0 && contraSet.has(contraRaw.toUpperCase());
    const rowVtype = isContraRow ? 'Contra' : vtype;

    const description = (rowText(r, 'DESCRIPTION') || rowText(r, 'Description')).trim();
    const chequeNo = (rowText(r, 'CHEQUE NO.') || rowText(r, 'ChequeNo') || rowText(r, 'Cheque No')).trim();
    const narrationParts = [description, chequeNo ? `Chq: ${chequeNo}` : ''].filter(Boolean);
    const narration = xmlEscape(narrationParts.join(' | '));

    vchCounter += 1;
    const vno = startVno != null ? String(startVno + vchCounter - 1) : String(vchCounter);

    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a(`    <VOUCHER VCHTYPE="${rowVtype}" ACTION="Create" OBJVIEW="Accounting Voucher View">`);
    a(`     <DATE>${dt}</DATE>`);
    a(`     <VOUCHERTYPENAME>${rowVtype}</VOUCHERTYPENAME>`);
    a(`     <VOUCHERNUMBER>${xmlEscape(vno)}</VOUCHERNUMBER>`);
    a(`     <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>`);
    a('     <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>');
    if (narration) a(`     <NARRATION>${narration}</NARRATION>`);

    if (rowVtype === 'Payment' || rowVtype === 'Contra') {
      a('     <ALLLEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${contraEsc}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(amt)}</AMOUNT>`);
      a('     </ALLLEDGERENTRIES.LIST>');
      a('     <ALLLEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${bankEsc}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(amt)}</AMOUNT>`);
      a('     </ALLLEDGERENTRIES.LIST>');
    } else {
      a('     <ALLLEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${bankEsc}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>-${fmtAmt(amt)}</AMOUNT>`);
      a('     </ALLLEDGERENTRIES.LIST>');
      a('     <ALLLEDGERENTRIES.LIST>');
      a(`      <LEDGERNAME>${contraEsc}</LEDGERNAME>`);
      a('      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      a(`      <AMOUNT>${fmtAmt(amt)}</AMOUNT>`);
      a('     </ALLLEDGERENTRIES.LIST>');
    }

    a('    </VOUCHER>');
    a('   </TALLYMESSAGE>');
  }

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return [lines.join('\n'), vchCounter];
}

module.exports = { generateSingleVoucherXml, PAYMENT_HEADERS, RECEIPT_HEADERS };
