/**
 * Minimal ElementTree-style XML tree builder, used only by gstr2bBridge.js
 * (the one Python module that builds vouchers via xml.etree.ElementTree
 * instead of raw string templating). Supports the two exact serializations
 * the Python code relies on:
 *  - toCompactXml: matches ET.tostring(root, encoding="unicode") —
 *    no whitespace, self-closes as `<TAG />` (space before slash).
 *  - toPrettyXml: matches minidom.parseString(...).toprettyxml(indent="  ") —
 *    2-space indent, self-closes as `<TAG/>` (no space).
 * Both escape text as &amp;/&lt;/&gt; and attributes as &amp;/&lt;/&gt;/&quot;.
 */

function Element(tag) {
  return { tag, attrs: {}, attrOrder: [], children: [], text: null };
}

function SubElement(parent, tag) {
  const el = Element(tag);
  parent.children.push(el);
  return el;
}

function setAttr(el, key, value) {
  if (!(key in el.attrs)) el.attrOrder.push(key);
  el.attrs[key] = value;
}

function escapeText(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function attrsToString(el) {
  return el.attrOrder.map((k) => ` ${k}="${escapeAttr(el.attrs[k])}"`).join('');
}

/** Matches ET.tostring(root, encoding="unicode"): no whitespace, `<TAG />` self-close. */
function toCompactXml(el) {
  const attrStr = attrsToString(el);
  if (el.children.length) {
    const inner = el.children.map(toCompactXml).join('');
    return `<${el.tag}${attrStr}>${inner}</${el.tag}>`;
  }
  if (el.text != null && el.text !== '') {
    return `<${el.tag}${attrStr}>${escapeText(el.text)}</${el.tag}>`;
  }
  return `<${el.tag}${attrStr} />`;
}

/** Matches minidom.parseString(ET.tostring(root)).toprettyxml(indent="  "):
 * 2-space indent per level, `<TAG/>` self-close (no space), one element per line. */
function toPrettyXml(el, depth = 0) {
  const pad = '  '.repeat(depth);
  const attrStr = attrsToString(el);
  if (el.children.length) {
    const inner = el.children.map((c) => toPrettyXml(c, depth + 1)).join('\n');
    return `${pad}<${el.tag}${attrStr}>\n${inner}\n${pad}</${el.tag}>`;
  }
  if (el.text != null && el.text !== '') {
    return `${pad}<${el.tag}${attrStr}>${escapeText(el.text)}</${el.tag}>`;
  }
  return `${pad}<${el.tag}${attrStr}/>`;
}

module.exports = { Element, SubElement, setAttr, toCompactXml, toPrettyXml };
