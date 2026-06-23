/** Standalone Create Stock Item voucher builder.
 * 1:1 port of Backend_Tally/tally_entry/stock_item.py. */
const { companyStaticBlock, normalizeStockGroupName, normalizeStockUnitName, xmlEscape } = require('./common');

/** items = list of dict: Name, Parent (stock group), Unit, HSNCode, GSTRate, Description. */
function generateStockitemXml(items, company) {
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

  for (const item of items) {
    const name = xmlEscape(item.Name);
    const parentRaw = normalizeStockGroupName(item.Parent || 'Primary');
    const parent = xmlEscape(parentRaw);
    const unitRaw = normalizeStockUnitName(item.Unit || 'Nos');
    const unit = xmlEscape(unitRaw);
    const hsn = xmlEscape(item.HSNCode || '');
    const gstR = String(item.GSTRate || '').replace(/%/g, '').trim();
    const desc = xmlEscape(item.Description || '');
    const supplyType = xmlEscape(item.TypeOfSupply || 'Goods');

    a('   <TALLYMESSAGE xmlns:UDF="TallyUDF">');
    a(`    <STOCKITEM NAME="${name}" ACTION="Create">`);
    a(`     <NAME>${name}</NAME>`);
    if (parentRaw && parentRaw.toLowerCase() !== 'primary') a(`     <PARENT>${parent}</PARENT>`);
    a(`     <BASEUNITS>${unit}</BASEUNITS>`);
    a('     <ISADDITIONALUNITS>NO</ISADDITIONALUNITS>');

    if (hsn || gstR) {
      a('     <GSTAPPLICABLE>Applicable</GSTAPPLICABLE>');
      a('     <GSTDETAILS.LIST>');
      if (hsn) a(`      <HSNCODE>${hsn}</HSNCODE>`);
      a('      <TAXABILITY>Taxable</TAXABILITY>');
      a(`      <SUPPLYTYPENAME>${supplyType}</SUPPLYTYPENAME>`);
      if (gstR) {
        let igst, half;
        const rateVal = parseFloat(gstR);
        if (Number.isFinite(rateVal)) {
          igst = rateVal.toFixed(2);
          half = (rateVal / 2).toFixed(2);
        } else {
          igst = gstR;
          half = gstR;
        }
        a('      <STATEWISEDETAILS.LIST>');
        a('       <STATENAME>Not Applicable</STATENAME>');
        a('       <RATEDETAILS.LIST>');
        a('        <GSTRATEDUTYHEAD>Integrated Tax</GSTRATEDUTYHEAD>');
        a(`        <GSTRATE>${igst}</GSTRATE>`);
        a('       </RATEDETAILS.LIST>');
        a('       <RATEDETAILS.LIST>');
        a('        <GSTRATEDUTYHEAD>Central Tax</GSTRATEDUTYHEAD>');
        a(`        <GSTRATE>${half}</GSTRATE>`);
        a('       </RATEDETAILS.LIST>');
        a('       <RATEDETAILS.LIST>');
        a('        <GSTRATEDUTYHEAD>State Tax</GSTRATEDUTYHEAD>');
        a(`        <GSTRATE>${half}</GSTRATE>`);
        a('       </RATEDETAILS.LIST>');
        a('      </STATEWISEDETAILS.LIST>');
      }
      a('     </GSTDETAILS.LIST>');
    }

    if (desc) a(`     <DESCRIPTION>${desc}</DESCRIPTION>`);

    a('    </STOCKITEM>');
    a('   </TALLYMESSAGE>');
  }

  a('  </REQUESTDATA>');
  a(' </IMPORTDATA></BODY>');
  a('</ENVELOPE>');
  return lines.join('\n');
}

module.exports = { generateStockitemXml };
