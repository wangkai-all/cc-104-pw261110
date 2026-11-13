/**
 * Very small, dependency-free XML reader for OOXML parts.
 *
 * It deliberately implements only what DOCX parsing needs:
 *   - elements, attributes, text, CDATA, comments, processing instructions
 *   - namespace prefixes (resolved against xmlns declarations in scope)
 *   - skipping <mc:Fallback> subtrees so alternate-content isn't read twice
 *
 * It never resolves external entities. Unknown entities are kept literally.
 */

const MAX_DEPTH = 200;
const MAX_NODES = 2_000_000;

function decodeEntities(input) {
  if (!input.includes('&')) return input;
  return input.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (m, body) => {
    if (body[0] === '#') {
      const hex = body[1] === 'x' || body[1] === 'X';
      const code = parseInt(body.slice(hex ? 2 : 1), hex ? 16 : 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return m;
      try {
        return String.fromCodePoint(code);
      } catch {
        return m;
      }
    }
    switch (body) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
      default: return m;
    }
  });
}

function readTag(xml, start) {
  // start points at '<'. Returns {raw, end} where end is index after '>'.
  let quote = null;
  for (let i = start + 1; i < xml.length; i++) {
    const ch = xml[i];
    if (quote) {
      if (ch === quote) quote = null;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === '>') {
      return { raw: xml.slice(start + 1, i), end: i + 1 };
    }
  }
  return { raw: xml.slice(start + 1), end: xml.length };
}

function parseAttrs(raw) {
  const attrs = {};
  const re = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(raw))) {
    const name = m[1];
    if (name === '/' || name.endsWith('/')) continue;
    const value = decodeEntities(m[2] !== undefined ? m[2] : m[3]);
    attrs[name] = value;
  }
  return attrs;
}

export function parseXml(xml) {
  if (typeof xml !== 'string') throw new TypeError('XML input must be a string');
  const doc = { name: '#document', local: '#document', prefix: '', ns: '', attrs: {}, children: [] };
  const stack = [doc];
  const nsStack = [Object.create(null)];
  let pos = 0;
  let nodeCount = 0;
  const len = xml.length;

  const currentNs = () => nsStack[nsStack.length - 1];

  const pushText = (text) => {
    if (!text) return;
    stack[stack.length - 1].children.push({ type: 'text', text: decodeEntities(text) });
  };

  while (pos < len) {
    const lt = xml.indexOf('<', pos);
    if (lt === -1) {
      pushText(xml.slice(pos));
      break;
    }
    if (lt > pos) pushText(xml.slice(pos, lt));

    if (xml.startsWith('<!--', lt)) {
      const end = xml.indexOf('-->', lt + 4);
      pos = end === -1 ? len : end + 3;
      continue;
    }
    if (xml.startsWith('<![CDATA[', lt)) {
      const end = xml.indexOf(']]>', lt + 9);
      const text = xml.slice(lt + 9, end === -1 ? len : end);
      stack[stack.length - 1].children.push({ type: 'text', text });
      pos = end === -1 ? len : end + 3;
      continue;
    }
    if (xml.startsWith('<?', lt)) {
      const end = xml.indexOf('?>', lt + 2);
      pos = end === -1 ? len : end + 2;
      continue;
    }
    if (xml.startsWith('<!', lt)) {
      // DOCTYPE / declarations. DOCX parts don't normally contain them.
      const tag = readTag(xml, lt);
      pos = tag.end;
      continue;
    }
    if (xml.startsWith('</', lt)) {
      const tag = readTag(xml, lt);
      if (stack.length > 1) {
        stack.pop();
        nsStack.pop();
      }
      pos = tag.end;
      continue;
    }

    const tag = readTag(xml, lt);
    pos = tag.end;
    let raw = tag.raw;
    const selfClosing = raw.trimEnd().endsWith('/');
    if (selfClosing) raw = raw.trimEnd().slice(0, -1);
    const spaceIdx = raw.search(/\s/);
    const qname = (spaceIdx === -1 ? raw : raw.slice(0, spaceIdx)).trim();
    if (!qname) continue;
    const attrs = parseAttrs(spaceIdx === -1 ? '' : raw.slice(spaceIdx));
    const colon = qname.indexOf(':');
    const prefix = colon === -1 ? '' : qname.slice(0, colon);
    const local = colon === -1 ? qname : qname.slice(colon + 1);

    const parentNs = currentNs();
    const scope = Object.create(parentNs);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'xmlns') scope[''] = v;
      else if (k.startsWith('xmlns:')) scope[k.slice(6)] = v;
    }
    const ns = scope[prefix] || '';

    if (prefix === 'mc' && local === 'Fallback') {
      // Skip the whole fallback subtree (alternate content is read from Choice).
      if (selfClosing) continue;
      const close = `</${qname}>`;
      const end = xml.indexOf(close, pos);
      pos = end === -1 ? len : end + close.length;
      continue;
    }

    if (stack.length > MAX_DEPTH) throw new Error('XML nesting too deep');
    if (++nodeCount > MAX_NODES) throw new Error('XML document too large');

    const node = { name: qname, local, prefix, ns, attrs, children: [] };
    stack[stack.length - 1].children.push(node);
    if (!selfClosing) {
      stack.push(node);
      nsStack.push(scope);
    }
  }
  return doc;
}

export function attr(node, localName) {
  if (!node || !node.attrs) return undefined;
  if (Object.prototype.hasOwnProperty.call(node.attrs, localName)) return node.attrs[localName];
  for (const [k, v] of Object.entries(node.attrs)) {
    const colon = k.indexOf(':');
    if ((colon === -1 ? k : k.slice(colon + 1)) === localName) return v;
  }
  return undefined;
}

export function firstChild(node, local, ns) {
  if (!node) return null;
  for (const child of node.children) {
    if (child.type === 'text') continue;
    if (child.local === local && (ns === undefined || child.ns === ns)) return child;
  }
  return null;
}

export function childrenOf(node, local, ns) {
  if (!node) return [];
  return node.children.filter((c) => c.type !== 'text' && c.local === local && (ns === undefined || c.ns === ns));
}

/** Depth-first search for an element with a local name (and optional namespace). */
export function findFirst(node, local, ns) {
  if (!node || node.type === 'text') return null;
  if (node.local === local && (ns === undefined || node.ns === ns)) return node;
  for (const child of node.children) {
    if (child.type === 'text') continue;
    const found = findFirst(child, local, ns);
    if (found) return found;
  }
  return null;
}

export function textContent(node) {
  if (!node) return '';
  if (node.type === 'text') return node.text;
  let out = '';
  for (const child of node.children) out += textContent(child);
  return out;
}
