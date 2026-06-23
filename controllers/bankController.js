/** Health check + bank-statement Excel template/upload endpoints.
 * 1:1 port of the corresponding routes in Backend_Tally/main.py. */
const excelBridge = require('../services/excelBridge');

function health(req, res) {
  res.json({ status: 'ok' });
}

async function downloadTemplate(req, res) {
  const content = await excelBridge.buildTemplateWorkbook();
  res.set({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': 'attachment; filename=BankStatement_Template.xlsx',
  });
  res.send(Buffer.from(content));
}

function isExcelFilename(filename) {
  const lower = String(filename || '').toLowerCase();
  return lower.endsWith('.xlsx') || lower.endsWith('.xlsm') || lower.endsWith('.xls');
}

async function processExcel(req, res) {
  const file = req.file;
  if (!file || !isExcelFilename(file.originalname)) {
    return res.status(400).json({ detail: 'Only Excel files (.xlsx, .xlsm, .xls) are supported.' });
  }
  const content = file.buffer;
  if (!content || !content.length) {
    return res.status(400).json({ detail: 'Uploaded file is empty.' });
  }

  let rows;
  try {
    rows = await excelBridge.readExcelRows(content);
  } catch (exc) {
    return res.status(422).json({ detail: `Could not read this Excel file: ${exc.message}` });
  }

  let warning = null;
  if (!rows.length) {
    warning = 'No rows could be read. Make sure the file\'s first row has headers matching '
      + 'the template (DATE, DESCRIPTION, CHEQUE NO., Debit, Credit, LEDGER).';
  }

  res.json({ success: true, row_count: rows.length, transactions: rows, warning });
}

module.exports = { health, downloadTemplate, processExcel, isExcelFilename };
