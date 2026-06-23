/** Excel template definitions (headers + sample row) and the generic
 * header-row -> dict reader shared by every Tally Entry mode.
 * 1:1 port of Backend_Tally/tally_entry/templates.py. */
const ExcelJS = require('exceljs');

const MODE_TEMPLATES = {
  accounting: {
    label: 'Sales Accounting Invoice',
    headers: ['VoucherType', 'Date', 'InvoiceNo', 'PartyLedger', 'PlaceOfSupply', 'GSTIN/UIN',
      'SalesLedger', 'TaxableValue', 'CGSTLedger', 'CGST Amount', 'SGSTLedger', 'SGST Amount',
      'IGSTLedger', 'IGST Amount', 'Narration'],
    sample: ['Sales', '01-Apr-2024', 'INV-001', 'Acme Corp', 'Maharashtra', '27AADCB2230M1Z2',
      'Sales Account', 10000, 'CGST', 900, 'SGST', 900, '', '', 'Sample sale'],
  },
  item: {
    label: 'Sales Item Invoice',
    headers: ['VoucherType', 'Date', 'InvoiceNo', 'GSTIN', 'PartyLedger', 'SalesLedger',
      'ItemName', 'Unit', 'Quantity', 'Rate', 'TaxableValue', 'CGSTLedger', 'CGST Amount',
      'SGSTLedger', 'SGST Amount', 'IGSTLedger', 'IGST Amount', 'Narration'],
    sample: ['Sales', '01-Apr-2024', 'INV-001', '27AADCB2230M1Z2', 'Acme Corp', 'Sales Account',
      'Widget', 'Nos', 5, 2000, 10000, 'CGST', 900, 'SGST', 900, '', '', 'Sample sale'],
  },
  purchase_accounting: {
    label: 'Purchase Accounting Invoice',
    headers: ['VoucherType', 'Voucher Date', 'SupplierInvoiceNo', 'SupplierInvoiceDate', 'PartyLedger',
      'GSTIN/UIN', 'PurchaseLedger', 'TaxableValue', 'CGSTLedger', 'CGST Rate', 'SGSTLedger',
      'SGST Rate', 'IGSTLedger', 'IGST Rate', 'TDS Ledger', 'TDS Rate', 'Narration'],
    sample: ['Purchase', '01-Apr-2024', 'P-001', '01-Apr-2024', 'Vendor Pvt Ltd', '07AAACE1234F1Z5',
      'Purchase Account', 5000, 'CGST', 9, 'SGST', 9, '', '', '', '', 'Sample purchase'],
  },
  purchase_item: {
    label: 'Purchase Item Invoice',
    headers: ['VoucherType', 'Date', 'InvoiceNo', 'SupplierInvoiceNo', 'SupplierInvoiceDate', 'GSTIN',
      'PartyLedger', 'Purchase Ledger', 'Item Name', 'Unit', 'Quantity', 'Rate', 'TaxableValue',
      'CGSTLedger', 'CGST Amount', 'SGSTLedger', 'SGST Amount', 'IGSTLedger', 'IGST Amount', 'Narration'],
    sample: ['Purchase', '01-Apr-2024', 'P-001', 'P-001', '01-Apr-2024', '07AAACE1234F1Z5', 'Vendor Pvt Ltd',
      'Purchase Account', 'Raw Material', 'Nos', 10, 500, 5000, 'CGST', 450, 'SGST', 450, '', '', 'Sample purchase'],
  },
  credit_note_accounting: {
    label: 'Credit Note (Accounting)',
    headers: ['Date', 'InvoiceNo', 'PartyLedger', 'GSTIN', 'Particular', 'TaxableValue',
      'CGSTLedger', 'CGSTRate', 'SGSTLedger', 'SGSTRate', 'IGSTLedger', 'IGSTRate', 'Narration'],
    sample: ['01-Apr-2024', 'CN-001', 'Acme Corp', '27AADCB2230M1Z2', 'Sales Account',
      1000, 'CGST', 9, 'SGST', 9, '', '', 'Sales return'],
  },
  credit_note_item: {
    label: 'Credit Note (Item)',
    headers: ['Date', 'InvoiceNo', 'PartyLedger', 'GSTIN', 'Item Name', 'Unit', 'Quantity', 'Rate',
      'TaxableValue', 'CGSTLedger', 'CGSTRate', 'SGSTLedger', 'SGSTRate', 'IGSTLedger', 'IGSTRate', 'Narration'],
    sample: ['01-Apr-2024', 'CN-001', 'Acme Corp', '27AADCB2230M1Z2', 'Widget', 'Nos', 5, 200,
      1000, 'CGST', 9, 'SGST', 9, '', '', 'Sales return'],
  },
  debit_note_accounting: {
    label: 'Debit Note (Accounting)',
    headers: ['Date', 'InvoiceNo', 'PartyLedger', 'GSTIN', 'Particular', 'TaxableValue',
      'CGSTLedger', 'CGSTRate', 'SGSTLedger', 'SGSTRate', 'IGSTLedger', 'IGSTRate', 'Narration'],
    sample: ['01-Apr-2024', 'DN-001', 'Acme Corp', '27AADCB2230M1Z2', 'Sales Account',
      1000, 'CGST', 9, 'SGST', 9, '', '', 'Price escalation'],
  },
  debit_note_item: {
    label: 'Debit Note (Item)',
    headers: ['Date', 'InvoiceNo', 'PartyLedger', 'GSTIN', 'Item Name', 'Unit', 'Quantity', 'Rate',
      'TaxableValue', 'CGSTLedger', 'CGSTRate', 'SGSTLedger', 'SGSTRate', 'IGSTLedger', 'IGSTRate', 'Narration'],
    sample: ['01-Apr-2024', 'DN-001', 'Acme Corp', '27AADCB2230M1Z2', 'Widget', 'Nos', 5, 200,
      1000, 'CGST', 9, 'SGST', 9, '', '', 'Price escalation'],
  },
  sales_journal: {
    label: 'Sales Journal',
    headers: ['Date', 'InvoiceNo', 'PartyLedger', 'GSTIN', 'Particular', 'TaxableValue', 'CGSTLedger',
      'CGSTRate', 'SGSTLedger', 'SGSTRate', 'IGSTLedger', 'IGSTRate', 'Narration'],
    sample: ['01-Apr-2024', 'J-001', 'Acme Corp', '27AADCB2230M1Z2', 'Sales Account', 1000,
      'CGST', 9, 'SGST', 9, '', '', 'Sales journal adjustment'],
  },
  purchase_journal: {
    label: 'Purchase Journal',
    headers: ['Date', 'InvoiceNo', 'PartyLedger', 'GSTIN', 'Particular', 'TaxableValue', 'CGSTLedger',
      'CGSTRate', 'SGSTLedger', 'SGSTRate', 'IGSTLedger', 'IGSTRate', 'TDSLedger', 'TDSRate', 'Narration'],
    sample: ['01-Apr-2024', 'J-001', 'Vendor Pvt Ltd', '07AAACE1234F1Z5', 'Purchase Account', 1000,
      'CGST', 9, 'SGST', 9, '', '', '', '', 'Purchase journal adjustment'],
  },
  payment: {
    label: 'Payment Voucher',
    headers: ['DATE', 'DESCRIPTION', 'CHEQUE NO.', 'Amount', 'LEDGER'],
    sample: ['01-Apr-2024', 'Office rent', '', 5000, 'Rent Expense'],
  },
  receipt: {
    label: 'Receipt Voucher',
    headers: ['DATE', 'DESCRIPTION', 'CHEQUE NO.', 'Amount', 'LEDGER'],
    sample: ['01-Apr-2024', 'Advance from customer', '', 5000, 'Acme Corp'],
  },
  sales_export: {
    label: 'Sales Export Accounting',
    headers: ['Date', 'InvoiceNo', 'PartyLedger', 'PartyName', 'Country', 'SalesLedger', 'Amount',
      'HSNCode', 'HSNDescription', 'Narration'],
    sample: ['01-Apr-2024', 'EXP-001', 'Overseas Client', 'Overseas Client Inc.', 'United States',
      'Export Sales', 50000, '9983', 'Consulting services', 'Export invoice'],
  },
};

/** Generic header-row -> dict reader. Row 1 = headers, data from row 2.
 * Keys keep the exact header text from the sheet; rowGet/rowText match
 * them tolerantly (whitespace/case-insensitive) so any of the header
 * variants the voucher builders look for will resolve correctly. */
async function readExcelRows(content) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(content);
  const ws = wb.worksheets[0];

  const headerRow = ws.getRow(1);
  const headers = [];
  for (let i = 1; i <= ws.columnCount; i++) headers.push(String(headerRow.getCell(i).value ?? '').trim());

  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const values = [];
    for (let i = 1; i <= ws.columnCount; i++) values.push(row.getCell(i).value);
    if (!values.length || values.every((v) => v == null || String(v).trim() === '')) continue;
    const rowDict = {};
    headers.forEach((header, idx) => {
      if (!header) return;
      rowDict[header] = idx < values.length ? values[idx] : null;
    });
    rows.push(rowDict);
  }
  return rows;
}

async function generateTemplateWorkbook(mode) {
  const template = MODE_TEMPLATES[mode];
  if (!template) throw new Error(`Unknown mode '${mode}' for template generation.`);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(template.label.slice(0, 31));
  template.headers.forEach((header, idx) => {
    const cell = ws.getCell(1, idx + 1);
    cell.value = header;
    cell.font = { bold: true };
  });
  template.sample.forEach((value, idx) => {
    ws.getCell(2, idx + 1).value = value === '' ? null : value;
  });
  for (let col = 1; col <= template.headers.length; col++) ws.getColumn(col).width = 20;
  return wb.xlsx.writeBuffer();
}

module.exports = { MODE_TEMPLATES, readExcelRows, generateTemplateWorkbook };
