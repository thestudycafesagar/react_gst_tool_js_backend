/**
 * GST portal (services.gst.gov.in) captcha + taxpayer-search scraper.
 * 1:1 port of Backend_Tally/gst_portal_bridge.py.
 *
 * Web port of the desktop tool's _GSTPortalSearcher / _GSTFetchDialog /
 * _parse_portal_data_to_extra — used to pre-fill the Create Ledger form
 * during missing/unmapped-party resolution. A headless Chrome instance is
 * kept alive server-side between the start-session (load captcha) and fetch
 * (submit + scrape) calls, keyed by a session id, with idle sessions swept
 * after a few minutes.
 */
const { Builder, By, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const { v4: uuidv4 } = require('uuid');
const { normalizeStateName, stateFromGstin } = require('./gstr2bBridge');

const GST_SEARCH_URL = 'https://services.gst.gov.in/services/searchtp';
const GST_MOBILE_UA = (
  'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) '
  + 'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Mobile Safari/537.36'
);

const SESSION_TTL_SECONDS = 300; // idle sessions auto-quit after 5 minutes

const PORTAL_NAME_BLANKS = new Set([
  'na', 'n/a', 'none', 'null', 'nil', 'not applicable',
  '* not applicable', '-', '--',
]);

const REG_TYPE_MAP = {
  regular: 'Regular', composition: 'Composition',
  unregistered: 'Unregistered/Consumer', consumer: 'Unregistered/Consumer',
  'input service distributor': 'Input Service Distributor', isd: 'Input Service Distributor',
  'sez unit': 'SEZ', 'sez developer': 'SEZ',
  'embassy / consulate': 'Overseas', overseas: 'Overseas',
};

function resolvePortalName(tradeName, legalName, fallback = '') {
  const clean = (v) => {
    const text = String(v || '').trim();
    return PORTAL_NAME_BLANKS.has(text.toLowerCase()) ? '' : text;
  };
  return clean(tradeName) || clean(legalName) || String(fallback || '').trim();
}

function normalizeStateForPortal(value) {
  const text = String(value || '').trim();
  if (['not applicable', '* not applicable', 'na', 'n/a'].includes(text.toLowerCase())) return '';
  return text;
}

/** Converts a scraped GST portal result into the lowercase 'extra' dict
 * shape consumed by the Create Ledger form (mailing_name, gstin, etc.). */
function parsePortalDataToExtra(portalData, ledgerName = '') {
  const details = portalData.details || {};
  const keyD = details.key_details || {};

  const tradeName = String(keyD['Trade Name'] || '').trim();
  const legalName = String(keyD['Legal Name of Business'] || '').trim();
  const mailing = resolvePortalName(tradeName, legalName, ledgerName);

  const ppob = String(keyD['Principal Place of Business'] || '').trim();
  let address1 = '', address2 = '', stateRaw = '', pincodeRaw = '';
  if (ppob) {
    let parts = ppob.split(',').map((p) => p.trim()).filter(Boolean);
    if (parts.length && /^\d{6}$/.test(parts[parts.length - 1])) {
      pincodeRaw = parts.pop();
    }
    if (parts.length) stateRaw = parts[parts.length - 1];
    if (parts.length > 1) {
      const mid = parts.slice(0, -1);
      address1 = mid.slice(0, 3).join(', ');
      address2 = mid.slice(3).join(', ');
    } else if (parts.length) {
      address1 = parts[0];
    }
  }

  const taxpayerType = String(keyD['Taxpayer Type'] || '').trim();
  const gstinVal = String(portalData.gstin || '').trim().toUpperCase();
  const regType = REG_TYPE_MAP[taxpayerType.toLowerCase()] || (gstinVal ? 'Regular' : 'Unregistered/Consumer');

  let normalizedState = normalizeStateName(normalizeStateForPortal(stateRaw));
  if (!normalizedState && gstinVal) normalizedState = normalizeStateName(stateFromGstin(gstinVal));

  return {
    mailing_name: mailing,
    gstin: gstinVal,
    gst_applicable: gstinVal ? 'Applicable' : 'Not Applicable',
    reg_type: regType,
    state: normalizedState,
    address1, address2,
    pincode: pincodeRaw,
    country: 'India',
    billwise: 'Yes',
  };
}

/** Selenium-based GST portal searcher — loads captcha, submits, scrapes. */
class GstPortalSearcher {
  constructor() {
    this.driver = null;
    this.lastUsed = Date.now() / 1000;
  }

  async ensureDriver() {
    if (this.driver !== null) return this.driver;
    const options = new chrome.Options();
    options.addArguments(
      '--headless=new',
      `--user-agent=${GST_MOBILE_UA}`,
      '--disable-blink-features=AutomationControlled',
      '--window-size=1200,900',
      '--no-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--lang=en-US',
    );
    options.excludeSwitches('enable-automation');
    options.setUserPreferences({});
    this.driver = await new Builder().forBrowser('chrome').setChromeOptions(options).build();
    await this.driver.manage().setTimeouts({ pageLoad: 60000 });
    try {
      await this.driver.sendDevToolsCommand('Page.addScriptToEvaluateOnNewDocument', {
        source: "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});",
      });
    } catch {
      // non-fatal, matches Python's bare except
    }
    return this.driver;
  }

  async triggerEvents(driver, element) {
    try {
      await driver.executeScript(
        "const e=arguments[0];"
        + "e.dispatchEvent(new Event('input',{bubbles:true}));"
        + "e.dispatchEvent(new Event('change',{bubbles:true}));"
        + "e.dispatchEvent(new Event('blur',{bubbles:true}));",
        element,
      );
    } catch {
      // non-fatal
    }
  }

  async findReadyCaptcha(driver) {
    try {
      const el = await driver.findElement(By.id('imgCaptcha'));
      const src = (await el.getAttribute('src') || '').toLowerCase();
      if (await el.isDisplayed() && src.includes('captcha')) return el;
    } catch {
      // not ready yet
    }
    return false;
  }

  async captureCaptchaBytes(driver, el) {
    try {
      await driver.wait(async () => driver.executeScript(
        'const i=arguments[0];return i&&i.complete&&i.naturalWidth>0;', el,
      ), 15000);
    } catch {
      // proceed to screenshot attempt regardless
    }
    try {
      const pngBase64 = await el.takeScreenshot();
      const buf = Buffer.from(pngBase64, 'base64');
      if (buf.length > 200) return buf;
    } catch {
      // fall through to empty buffer
    }
    return Buffer.alloc(0);
  }

  async findCaptchaInput(driver) {
    const locators = [
      [By.id('captcha')], [By.id('captchaCode')], [By.id('captchaText')],
      [By.name('captcha')], [By.name('captchaCode')],
      [By.css("input[ng-model*='captcha' i]")],
      [By.css("input[id*='captcha' i]")],
    ];
    for (const [locator] of locators) {
      try {
        const el = await driver.findElement(locator);
        if (await el.isDisplayed()) return el;
      } catch {
        // try next locator
      }
    }
    return null;
  }

  async loadCaptcha(gstin) {
    this.lastUsed = Date.now() / 1000;
    const driver = await this.ensureDriver();
    await driver.get(GST_SEARCH_URL);
    const inp = await driver.wait(until.elementLocated(By.id('for_gstin')), 35000);
    await driver.wait(until.elementIsVisible(inp), 35000);
    await inp.clear();
    await inp.sendKeys(gstin);
    await this.triggerEvents(driver, inp);
    const el = await driver.wait(async () => this.findReadyCaptcha(driver), 35000);
    if (!el) throw new Error('Timed out waiting for captcha on GST portal.');
    return this.captureCaptchaBytes(driver, el);
  }

  async fetch(gstin, captchaText) {
    this.lastUsed = Date.now() / 1000;
    const driver = await this.ensureDriver();
    const currentUrl = (await driver.getCurrentUrl() || '').toLowerCase();
    if (!currentUrl.includes('searchtp')) await driver.get(GST_SEARCH_URL);

    const inp = await driver.wait(until.elementLocated(By.id('for_gstin')), 25000);
    await driver.wait(until.elementIsVisible(inp), 25000);
    await inp.clear();
    await inp.sendKeys(gstin);
    await this.triggerEvents(driver, inp);

    let capInp = null;
    for (let i = 0; i < 40; i++) {
      capInp = await this.findCaptchaInput(driver);
      if (capInp) break;
      await new Promise((r) => setTimeout(r, 300));
    }
    if (!capInp) throw new Error('Captcha input not found on page.');
    await capInp.clear();
    await capInp.sendKeys(captchaText);

    const btn = await driver.wait(until.elementLocated(By.id('lotsearch')), 25000);
    await driver.wait(until.elementIsVisible(btn), 25000);
    await btn.click();

    const end = Date.now() + 35000;
    while (Date.now() < end) {
      let found = false;
      for (const sel of ['#lottable', '#searchResult', '.panel-body', 'table']) {
        try {
          const el = await driver.findElement(By.css(sel));
          const text = (await el.getText() || '').trim();
          if (text) { found = true; break; }
        } catch {
          // try next selector
        }
      }
      if (found) break;
      await new Promise((r) => setTimeout(r, 500));
    }

    const details = await driver.executeScript(`
            const getText=e=>e?(e.textContent||'').replace(/\\s+/g,' ').trim():'';
            const result={};const container=document.querySelector('#lottable');
            if(!container)return result;
            const pairs={};
            const cols=container.querySelectorAll('.tbl-format .inner .col-sm-4,.tbl-format .inner .col-sm-12');
            cols.forEach(col=>{
                const label=getText(col.querySelector('strong'));if(!label)return;
                const list=col.querySelector('ul.jurisdictList');
                if(list){const items=Array.from(list.querySelectorAll('li')).map(getText).filter(Boolean);
                    if(items.length){pairs[label]=items;return;}}
                const word=col.querySelector('.wordCls');
                if(word){pairs[label]=getText(word);return;}
                const ps=Array.from(col.querySelectorAll('p')).map(getText).filter(Boolean);
                if(ps.length>1)pairs[label]=ps[ps.length-1];
                else if(ps.length===1)pairs[label]=ps[0];
            });
            result.key_details=pairs;
            return result;
        `) || {};

    return { gstin, details };
  }

  async quit() {
    if (this.driver) {
      try {
        await this.driver.quit();
      } catch {
        // best-effort cleanup
      }
      this.driver = null;
    }
  }
}

const sessions = new Map();

function sweepExpiredSessions() {
  const now = Date.now() / 1000;
  const expired = [];
  for (const [sid, searcher] of sessions.entries()) {
    if (now - searcher.lastUsed > SESSION_TTL_SECONDS) expired.push([sid, searcher]);
  }
  for (const [sid] of expired) sessions.delete(sid);
  return Promise.all(expired.map(([, searcher]) => searcher.quit()));
}

/** Launches a headless Chrome session, loads the GSTIN search page, and
 * returns the captcha image as base64 PNG along with a session id. */
async function startSession(gstin) {
  await sweepExpiredSessions();
  const searcher = new GstPortalSearcher();
  let png;
  try {
    png = await searcher.loadCaptcha(gstin);
  } catch (exc) {
    await searcher.quit();
    return { success: false, error: exc.message };
  }
  if (!png || !png.length) {
    await searcher.quit();
    return { success: false, error: 'Could not capture the captcha image. Try again.' };
  }

  const sessionId = uuidv4().replace(/-/g, '');
  sessions.set(sessionId, searcher);
  return { success: true, sessionId, captchaPngBase64: png.toString('base64') };
}

/** Reloads a fresh captcha on an existing session (e.g. after a wrong
 * captcha guess) without spinning up a new browser instance. */
async function reloadCaptcha(sessionId, gstin) {
  const searcher = sessions.get(sessionId);
  if (!searcher) return { success: false, error: 'Session expired. Start a new GST portal search.' };
  let png;
  try {
    png = await searcher.loadCaptcha(gstin);
  } catch (exc) {
    return { success: false, error: exc.message };
  }
  if (!png || !png.length) return { success: false, error: 'Could not capture the captcha image. Try again.' };
  return { success: true, captchaPngBase64: png.toString('base64') };
}

/** Submits the captcha, scrapes the result table, and ends the session
 * (single-use, matching the desktop dialog's one-shot fetch-then-close). */
async function fetchDetails(sessionId, gstin, captchaText, ledgerName = '') {
  const searcher = sessions.get(sessionId);
  sessions.delete(sessionId);
  if (!searcher) return { success: false, error: 'Session expired. Start a new GST portal search.' };
  let data;
  try {
    data = await searcher.fetch(gstin, captchaText);
  } catch (exc) {
    await searcher.quit();
    return { success: false, error: exc.message };
  }

  const keyDetails = (data.details || {}).key_details || {};
  if (!Object.keys(keyDetails).length) {
    await searcher.quit();
    return { success: false, error: 'No result found — check the GSTIN and captcha text, then try again.' };
  }

  const extra = parsePortalDataToExtra(data, ledgerName);
  await searcher.quit();
  return { success: true, extra };
}

async function cancelSession(sessionId) {
  const searcher = sessions.get(sessionId);
  sessions.delete(sessionId);
  if (searcher) await searcher.quit();
  return { success: true };
}

module.exports = {
  GST_SEARCH_URL, GST_MOBILE_UA, SESSION_TTL_SECONDS,
  resolvePortalName, normalizeStateForPortal, parsePortalDataToExtra,
  GstPortalSearcher, startSession, reloadCaptcha, fetchDetails, cancelSession,
};
