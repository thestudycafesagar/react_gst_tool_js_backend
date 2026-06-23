/** Shared string/date helpers used across services — ports of small
 * standalone functions duplicated in several Python bridge modules. */

function xmlEscape(s) {
  if (!s) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function fmtAmt(num) {
  return Number(num).toFixed(2);
}

/** Strips control chars / invalid numeric character refs that break XML
 * parsers on some real-world Tally exports. */
function sanitizeTallyXml(text) {
  if (!text) return '';
  let cleaned = text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  cleaned = cleaned.replace(/&#(x[0-9A-Fa-f]+|\d+);/g, (match, token) => {
    let cp;
    try {
      cp = token.toLowerCase().startsWith('x') ? parseInt(token.slice(1), 16) : parseInt(token, 10);
    } catch {
      return '';
    }
    const isValid = cp === 0x9 || cp === 0xa || cp === 0xd
      || (cp >= 0x20 && cp <= 0xd7ff)
      || (cp >= 0xe000 && cp <= 0xfffd)
      || (cp >= 0x10000 && cp <= 0x10ffff);
    return isValid ? match : '';
  });
  return cleaned;
}

function normalizeCompanyName(value) {
  if (value == null) return '';
  let text = String(value);
  // html.unescape equivalent for the handful of entities Tally responses use
  text = text
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#39;/g, "'");
  text = text.replace(/\x00/g, '');
  text = text.replace(/[\x01-\x1F\x7F]/g, ' ');
  return text.replace(/\s+/g, ' ').trim();
}

function isValidCompanyName(value) {
  const name = normalizeCompanyName(value);
  if (!name) return false;
  return !/^\d+$/.test(name);
}

function panFromGstin(gstin) {
  const g = String(gstin || '').trim().toUpperCase();
  return g.length === 15 ? g.slice(2, 12) : '';
}

function isGstinLike(value) {
  return /^\d{2}[A-Z0-9]{13}$/.test(String(value || '').trim().toUpperCase());
}

const GSTIN_STATE_MAP = {
  '01': 'Jammu And Kashmir', '02': 'Himachal Pradesh', '03': 'Punjab', '04': 'Chandigarh',
  '05': 'Uttarakhand', '06': 'Haryana', '07': 'Delhi', '08': 'Rajasthan', '09': 'Uttar Pradesh',
  10: 'Bihar', 11: 'Sikkim', 12: 'Arunachal Pradesh', 13: 'Nagaland', 14: 'Manipur',
  15: 'Mizoram', 16: 'Tripura', 17: 'Meghalaya', 18: 'Assam', 19: 'West Bengal',
  20: 'Jharkhand', 21: 'Odisha', 22: 'Chhattisgarh', 23: 'Madhya Pradesh',
  24: 'Gujarat', 25: 'Daman And Diu', 26: 'Dadra And Nagar Haveli And Daman And Diu',
  27: 'Maharashtra', 29: 'Karnataka', 30: 'Goa', 31: 'Lakshadweep', 32: 'Kerala',
  33: 'Tamil Nadu', 34: 'Puducherry', 35: 'Andaman And Nicobar Islands', 36: 'Telangana',
  37: 'Andhra Pradesh', 38: 'Ladakh', 97: 'Other Territory', 99: 'Centre Jurisdiction',
};

function stateNameFromGstin(gstin) {
  const code = String(gstin || '').trim().toUpperCase().slice(0, 2);
  return GSTIN_STATE_MAP[code] || '';
}

function normalizeStateForLedger(value) {
  const text = String(value || '').trim();
  if (['not applicable', '* not applicable', 'na', 'n/a'].includes(text.toLowerCase())) return '';
  return text;
}

function normalizeGstApplicable(value, gstin = '') {
  const raw = String(value || '').trim();
  const key = raw.toLowerCase();
  if (['applicable', 'yes', 'y', 'true', '1', 'registered', 'regular', 'gst applicable'].includes(key)) return 'Applicable';
  if (['not applicable', 'no', 'n', 'false', '0', 'na', 'n/a', 'notapplicable'].includes(key)) return 'Not Applicable';
  if (gstin) return 'Applicable';
  return raw;
}

function normalizeGstRegistrationType(value, gstin = '', gstApplicable = '') {
  const raw = String(value || '').trim();
  if (!raw) {
    if (gstin || String(gstApplicable).trim().toLowerCase() === 'applicable') return 'Regular';
    return '';
  }
  const mapping = {
    regular: 'Regular', registered: 'Regular', composition: 'Composition',
    consumer: 'Consumer', unregistered: 'Unregistered',
    sez: 'SEZ', 'sez unit': 'SEZ', 'sez developer': 'SEZ', overseas: 'Overseas',
  };
  return mapping[raw.toLowerCase()] || raw;
}

function currentFyStart() {
  const today = new Date();
  const fyStartYear = today.getMonth() + 1 >= 4 ? today.getFullYear() : today.getFullYear() - 1;
  return `${fyStartYear}0401`;
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

function toTallyDateString(d) {
  return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`;
}

const DATE_FORMAT_REGEXES = [
  // [regex, (m) => {year, month, day}]
  [/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/, (m) => ({ d: +m[1], mo: +m[2], y: +m[3] })],
  [/^(\d{1,2})[-/](\d{1,2})[-/](\d{2})$/, (m) => ({ d: +m[1], mo: +m[2], y: 2000 + +m[3] })],
  [/^(\d{4})-(\d{1,2})-(\d{1,2})$/, (m) => ({ y: +m[1], mo: +m[2], d: +m[3] })],
];
const MONTH_NAMES = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

function monthFromName(name) {
  const idx = MONTH_NAMES.findIndex((m) => name.toLowerCase().startsWith(m));
  return idx === -1 ? null : idx + 1;
}

/** Parses a bank-statement-style date in many formats — mirrors
 * parse_statement_datetime's DATE_FORMATS list. Returns a Date or null. */
function parseStatementDatetime(dt) {
  if (dt == null || dt === '') return null;
  if (dt instanceof Date) return dt;
  const text = String(dt).replace(/\s+/g, ' ').trim();
  if (!text) return null;

  for (const [regex, extract] of DATE_FORMAT_REGEXES) {
    const m = text.match(regex);
    if (m) {
      const { y, mo, d } = extract(m);
      const date = new Date(y, mo - 1, d);
      if (date.getFullYear() === y && date.getMonth() === mo - 1 && date.getDate() === d) return date;
    }
  }
  // "DD-Mon-YYYY" / "DD Mon YYYY" / 2-digit-year variants, with full or abbreviated month names
  const m2 = text.match(/^(\d{1,2})[-\s]([A-Za-z]+)[-\s](\d{2,4})$/);
  if (m2) {
    const day = +m2[1];
    const month = monthFromName(m2[2]);
    let year = +m2[3];
    if (year < 100) year += 2000;
    if (month) {
      const date = new Date(year, month - 1, day);
      if (date.getDate() === day) return date;
    }
  }
  return null;
}

/** Validates + reformats a manually-typed date to Tally's YYYYMMDD —
 * mirrors _normalize_manual_date_to_tally, throwing the same message. */
function normalizeManualDateToTally(dateText) {
  const text = String(dateText || '').trim();
  if (!text) throw new Error('Custom date is empty.');
  const patterns = [
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,
    /^(\d{1,2})-(\d{1,2})-(\d{4})$/,
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/,
    /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/,
    /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/,
  ];
  const orders = ['dmy', 'dmy', 'ymd', 'ymd', 'dmy'];
  for (let i = 0; i < patterns.length; i++) {
    const m = text.match(patterns[i]);
    if (!m) continue;
    let y, mo, d;
    if (orders[i] === 'dmy') { d = +m[1]; mo = +m[2]; y = +m[3]; } else { y = +m[1]; mo = +m[2]; d = +m[3]; }
    const date = new Date(y, mo - 1, d);
    if (date.getFullYear() === y && date.getMonth() === mo - 1 && date.getDate() === d) {
      return toTallyDateString(date);
    }
  }
  throw new Error('Invalid custom date format. Use DD/MM/YYYY, DD-MM-YYYY, or YYYY-MM-DD.');
}

module.exports = {
  xmlEscape, fmtAmt, sanitizeTallyXml, normalizeCompanyName, isValidCompanyName,
  panFromGstin, isGstinLike, stateNameFromGstin, normalizeStateForLedger,
  normalizeGstApplicable, normalizeGstRegistrationType, currentFyStart,
  parseStatementDatetime, normalizeManualDateToTally, toTallyDateString, pad2,
};
