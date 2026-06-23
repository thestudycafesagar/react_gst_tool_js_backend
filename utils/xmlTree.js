/**
 * Minimal ElementTree-like XML tree for parsing Tally's XML responses.
 *
 * The Python backend uses xml.etree.ElementTree, walking `root.iter()` /
 * `root.find(".//TAG")` / child text. fast-xml-parser's default object
 * shape doesn't preserve element order/attributes the same way, so this
 * module normalizes its `preserveOrder` output into a small uniform tree
 * `{ tag, attrs, text, children }` that supports the same three
 * operations the Python code relies on, with the same semantics.
 */
const { XMLParser } = require('fast-xml-parser');

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
});

// fast-xml-parser's numeric-character-reference decoding (even with
// htmlEntities: true) is inconsistent on Tally's large multi-megabyte
// ledger-list responses — some &#13;/&#10; refs (common in ledger names
// pasted from Word/Excel) silently survive as literal text. Decoding them
// ourselves before parsing guarantees the same result Python's
// ElementTree gives every time, regardless of document size.
function decodeNumericEntities(text) {
  return text
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)));
}

const ATTR_KEY = ':@';
const TEXT_KEY = '#text';

function normalizeNode(tag, rawNode) {
  const attrs = {};
  const children = [];
  let text = '';

  if (rawNode && rawNode[ATTR_KEY]) {
    for (const [k, v] of Object.entries(rawNode[ATTR_KEY])) {
      if (k.startsWith('@_')) attrs[k.slice(2)] = String(v);
    }
  }

  const childArray = Array.isArray(rawNode) ? rawNode : (rawNode && rawNode[tag]) || [];
  for (const entry of childArray) {
    if (entry == null) continue;
    const keys = Object.keys(entry).filter((k) => k !== ATTR_KEY);
    for (const childTag of keys) {
      if (childTag === TEXT_KEY) {
        text += String(entry[TEXT_KEY]);
        continue;
      }
      children.push(normalizeNode(childTag, entry));
    }
  }

  return { tag, attrs, text: text.trim(), children };
}

/**
 * Parses an XML string into the root node, or throws on malformed XML —
 * mirrors xml.etree.ElementTree.fromstring(text) raising ET.ParseError.
 */
// XML 1.0 forbids literal control characters other than tab/LF/CR (0x09/0x0A/0x0D)
// in element/attribute content — ElementTree raises ParseError on these. Some
// Tally ledger names (pasted from elsewhere) contain raw bytes like BEL (0x07);
// fast-xml-parser silently accepts them, so without this check we'd succeed
// where Python fails and never fall through to the caller's sanitized retry.
const ILLEGAL_XML_CHAR_RE = /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/;

function parseXml(text) {
  if (!text || !String(text).trim()) {
    throw new Error('Empty XML document');
  }
  const decoded = decodeNumericEntities(text);
  if (ILLEGAL_XML_CHAR_RE.test(decoded)) {
    throw new Error('XML parse error: illegal literal control character');
  }
  let parsed;
  try {
    parsed = parser.parse(decoded);
  } catch (exc) {
    throw new Error(`XML parse error: ${exc.message}`);
  }
  // Skip any leading declaration/comment pseudo-nodes to find the real root element.
  for (const entry of parsed) {
    const keys = Object.keys(entry).filter((k) => k !== ATTR_KEY && k !== '?xml');
    if (keys.length) {
      const tag = keys[0];
      return normalizeNode(tag, entry);
    }
  }
  throw new Error('No root element found');
}

/** Depth-first traversal yielding every node including the root — mirrors root.iter(). */
function* iterAll(node) {
  yield node;
  for (const child of node.children) {
    yield* iterAll(child);
  }
}

/** All descendant (not self) nodes whose tag matches — mirrors root.findall(f".//{tag}"). */
function findAll(node, tag) {
  const upper = tag.toUpperCase();
  const out = [];
  for (const child of node.children) {
    for (const n of iterAll(child)) {
      if (n.tag.toUpperCase() === upper) out.push(n);
    }
  }
  return out;
}

/** First descendant (not self) matching tag, or null — mirrors root.find(f".//{tag}"). */
function find(node, tag) {
  const all = findAll(node, tag);
  return all.length ? all[0] : null;
}

/** Direct child's text by tag, or default — mirrors node.findtext(tag, default). */
function findText(node, tag, defaultValue = '') {
  const upper = tag.toUpperCase();
  for (const child of node.children) {
    if (child.tag.toUpperCase() === upper) return child.text;
  }
  return defaultValue;
}

module.exports = { parseXml, iterAll, findAll, find, findText };
