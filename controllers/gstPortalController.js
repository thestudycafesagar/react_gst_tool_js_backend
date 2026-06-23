/** GST Portal Search (Selenium captcha scrape) endpoints (/gst-portal/*).
 * 1:1 port of the corresponding routes in Backend_Tally/main.py.
 *
 * Session-based REST flow: start-session opens headless Chrome and returns
 * a captcha image; fetch-details submits the typed captcha, scrapes the
 * result, and closes the browser (single-use, same as the desktop dialog). */
const gstPortalBridge = require('../services/gstPortalBridge');

async function startSession(req, res) {
  const gstin = String(req.body.gstin || '').trim().toUpperCase();
  if (!gstin) return res.json({ success: false, sessionId: '', captchaPngBase64: '', error: 'GSTIN is required.' });
  const result = await gstPortalBridge.startSession(gstin);
  res.json({
    success: result.success, sessionId: result.sessionId || '',
    captchaPngBase64: result.captchaPngBase64 || '', error: result.error ?? null,
  });
}

async function reloadCaptcha(req, res) {
  const gstin = String(req.body.gstin || '').trim().toUpperCase();
  if (!gstin) return res.json({ success: false, captchaPngBase64: '', error: 'GSTIN is required.' });
  const result = await gstPortalBridge.reloadCaptcha(req.body.sessionId, gstin);
  res.json({
    success: result.success, captchaPngBase64: result.captchaPngBase64 || '', error: result.error ?? null,
  });
}

async function fetchDetails(req, res) {
  const gstin = String(req.body.gstin || '').trim().toUpperCase();
  const captchaText = String(req.body.captchaText || '').trim();
  if (!gstin) return res.json({ success: false, extra: {}, error: 'GSTIN is required.' });
  if (!captchaText) return res.json({ success: false, extra: {}, error: 'Enter the captcha text first.' });
  const result = await gstPortalBridge.fetchDetails(req.body.sessionId, gstin, captchaText, req.body.ledgerName || '');
  res.json({ success: result.success, extra: result.extra || {}, error: result.error ?? null });
}

async function cancelSession(req, res) {
  const result = await gstPortalBridge.cancelSession(req.body.sessionId);
  res.json(result);
}

module.exports = { startSession, reloadCaptcha, fetchDetails, cancelSession };
