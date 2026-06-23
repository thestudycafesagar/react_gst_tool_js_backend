/** Mounts every route from Backend_Tally/main.py at the exact same path,
 * method, and request/response contract. */
const express = require('express');
const upload = require('../middleware/upload');

const bankController = require('../controllers/bankController');
const tallyController = require('../controllers/tallyController');
const tallyEntryController = require('../controllers/tallyEntryController');
const gstr2bController = require('../controllers/gstr2bController');
const gstPortalController = require('../controllers/gstPortalController');

const router = express.Router();

// ─── Bank statement ──────────────────────────────────────────────────────
router.get('/health', bankController.health);
router.get('/template/download', bankController.downloadTemplate);
router.post('/process-excel', upload.single('file'), bankController.processExcel);

// ─── Tally bridge ────────────────────────────────────────────────────────
router.post('/tally/test-connection', tallyController.testConnection);
router.post('/tally/companies', tallyController.companies);
router.post('/tally/bank-ledgers', tallyController.bankLedgers);
router.post('/tally/all-ledgers', tallyController.allLedgers);
router.post('/tally/stock-items', tallyController.stockItems);
router.post('/tally/create-ledger', tallyController.createLedger);
router.post('/tally-entry/create-stock-item', tallyEntryController.createStockItem);
router.post('/tally/generate-xml', tallyController.generateXml);
router.post('/tally/push-vouchers', tallyController.pushVouchers);

// ─── GSTR-2B bridge ──────────────────────────────────────────────────────
router.post('/gstr2b/process-excel', upload.single('file'), gstr2bController.processExcel);
router.post('/gstr2b/download-template', upload.single('file'), gstr2bController.downloadTemplate);
router.post('/gstr2b/upload-template', upload.single('file'), gstr2bController.uploadTemplate);
router.post('/gstr2b/validate-tax', gstr2bController.validateTax);
router.post('/gstr2b/generate-xml', gstr2bController.generateXml);
router.post('/gstr2b/push-vouchers', gstr2bController.pushVouchers);
router.post('/gstr2b/push-manual-voucher', gstr2bController.pushManualVoucher);

// ─── Tally Entry (generic Sales/Purchase voucher entry) ─────────────────
router.post('/tally-entry/process-excel', upload.single('file'), tallyEntryController.processExcel);
router.get('/tally-entry/download-template', tallyEntryController.downloadTemplate);
router.post('/tally-entry/generate-xml', tallyEntryController.generateXml);
router.post('/tally-entry/push-vouchers', tallyEntryController.pushVouchers);
router.post('/tally-entry/push-xml', tallyEntryController.pushXml);

// ─── GST Portal Search (Selenium captcha scrape) ────────────────────────
router.post('/gst-portal/start-session', gstPortalController.startSession);
router.post('/gst-portal/reload-captcha', gstPortalController.reloadCaptcha);
router.post('/gst-portal/fetch-details', gstPortalController.fetchDetails);
router.post('/gst-portal/cancel-session', gstPortalController.cancelSession);

module.exports = router;
