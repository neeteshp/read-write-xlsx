// Tiny streaming (SAX-style) XML tokenizer, sufficient for SpreadsheetML.
// Element and attribute names are reported without namespace prefixes
// (attributes are additionally available under their full prefixed name).

const ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }

export function decodeEntities(s) {
  if (s.indexOf('&') === -1) return s
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, e) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)
      return Number.isFinite(code) ? String.fromCodePoint(code) : m
    }
    return ENTITIES[e] ?? m
  })
}

// Characters that are not allowed in XML 1.0 documents.
// eslint-disable-next-line no-control-regex
const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g

export function escapeXml(s) {
  return String(s)
    .replace(INVALID_XML_CHARS, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

const localName = (name) => {
  const i = name.indexOf(':')
  return i === -1 ? name : name.slice(i + 1)
}

const ATTR_RE = /([^\s=/]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

function parseAttrs(body) {
  const attrs = {}
  ATTR_RE.lastIndex = 0
  let m
  while ((m = ATTR_RE.exec(body))) {
    const value = decodeEntities(m[2] ?? m[3])
    attrs[m[1]] = value
    const ln = localName(m[1])
    if (!(ln in attrs)) attrs[ln] = value
  }
  return attrs
}

/**
 * @param {string} xml
 * @param {{
 *   open?: (name: string, attrs: Record<string,string>) => void,
 *   close?: (name: string) => void,
 *   text?: (text: string) => void
 * }} h
 */
export function parseXml(xml, h) {
  const len = xml.length
  let i = 0
  while (i < len) {
    const lt = xml.indexOf('<', i)
    if (lt === -1) {
      if (h.text) h.text(decodeEntities(xml.slice(i)))
      break
    }
    if (lt > i && h.text) h.text(decodeEntities(xml.slice(i, lt)))

    const next = xml.charCodeAt(lt + 1)
    if (next === 63 /* ? */) {
      i = xml.indexOf('?>', lt) + 2
      if (i === 1) break
      continue
    }
    if (next === 33 /* ! */) {
      if (xml.startsWith('<!--', lt)) {
        const end = xml.indexOf('-->', lt)
        i = end === -1 ? len : end + 3
      } else if (xml.startsWith('<![CDATA[', lt)) {
        const end = xml.indexOf(']]>', lt)
        const stop = end === -1 ? len : end
        if (h.text) h.text(xml.slice(lt + 9, stop))
        i = end === -1 ? len : end + 3
      } else {
        const end = xml.indexOf('>', lt)
        i = end === -1 ? len : end + 1
      }
      continue
    }
    if (next === 47 /* / */) {
      const end = xml.indexOf('>', lt)
      if (h.close) h.close(localName(xml.slice(lt + 2, end).trim()))
      i = end + 1
      continue
    }

    // Opening tag: find the closing '>' while respecting quoted attribute values.
    let j = lt + 1
    let quote = 0
    for (; j < len; j++) {
      const c = xml.charCodeAt(j)
      if (quote) {
        if (c === quote) quote = 0
      } else if (c === 34 || c === 39) {
        quote = c
      } else if (c === 62 /* > */) {
        break
      }
    }
    let body = xml.slice(lt + 1, j)
    const selfClosing = body.endsWith('/')
    if (selfClosing) body = body.slice(0, -1)
    const ws = body.search(/\s/)
    const rawName = ws === -1 ? body : body.slice(0, ws)
    const name = localName(rawName)
    if (h.open) h.open(name, ws === -1 ? {} : parseAttrs(body.slice(ws)))
    if (selfClosing && h.close) h.close(name)
    i = j + 1
  }
}

/**
 * Parses small XML documents into a simple element tree.
 * @returns {{name: string, attrs: Record<string,string>, children: any[], text: string}}
 */
export function parseXmlTree(xml) {
  const root = { name: '#root', attrs: {}, children: [], text: '' }
  const stack = [root]
  parseXml(xml, {
    open(name, attrs) {
      const el = { name, attrs, children: [], text: '' }
      stack[stack.length - 1].children.push(el)
      stack.push(el)
    },
    close() {
      if (stack.length > 1) stack.pop()
    },
    text(t) {
      stack[stack.length - 1].text += t
    }
  })
  return root
}

/** Depth-first search for elements with the given local name. */
export function findAll(el, name, out = []) {
  for (const child of el.children) {
    if (child.name === name) out.push(child)
    findAll(child, name, out)
  }
  return out
}

export function findFirst(el, name) {
  for (const child of el.children) {
    if (child.name === name) return child
    const found = findFirst(child, name)
    if (found) return found
  }
  return null
}
