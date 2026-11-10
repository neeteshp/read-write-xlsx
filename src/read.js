import { readZip } from './zip.js'
import { parseXml, parseXmlTree, findAll, findFirst } from './xml.js'
import { parseCellRef, excelSerialToDate, isDateFormat } from './util.js'
import { parseWithSchema } from './schema.js'

const textDecoder = new TextDecoder('utf-8')

/** Normalises any supported input into a Uint8Array. */
export async function toBytes(input) {
  if (input instanceof Uint8Array) return input
  if (input instanceof ArrayBuffer) return new Uint8Array(input)
  if (ArrayBuffer.isView(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength)
  }
  if (input && typeof input.arrayBuffer === 'function') {
    // Blob / File / Response
    return new Uint8Array(await input.arrayBuffer())
  }
  throw new TypeError(
    'Unsupported input. Pass a Uint8Array, ArrayBuffer, Buffer, Blob or File ' +
      '(or a file path when using the "/node" entry).'
  )
}

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1)
  const parts = base.split('/').slice(0, -1)
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop()
    else if (seg !== '.') parts.push(seg)
  }
  return parts.join('/')
}

function relsPathFor(path) {
  const i = path.lastIndexOf('/')
  return path.slice(0, i + 1) + '_rels/' + path.slice(i + 1) + '.rels'
}

/** Opens a workbook and returns lazy accessors for its parts. */
async function openWorkbook(input) {
  const entries = readZip(await toBytes(input))
  // Some producers vary the case of part names, so match case-insensitively.
  const lower = new Map()
  for (const k of entries.keys()) lower.set(k.toLowerCase(), k)
  const readText = async (path) => {
    const key = entries.has(path) ? path : lower.get(path.toLowerCase())
    if (!key) return null
    return textDecoder.decode(await entries.get(key)())
  }

  // Find the main workbook part via the package relationships.
  let workbookPath = 'xl/workbook.xml'
  const rootRels = await readText('_rels/.rels')
  if (rootRels) {
    const rel = findAll(parseXmlTree(rootRels), 'Relationship').find((r) =>
      /\/officeDocument$/.test(r.attrs.Type || '')
    )
    if (rel) workbookPath = resolvePath('', rel.attrs.Target)
  }
  const workbookXml = await readText(workbookPath)
  if (!workbookXml) throw new Error('Not a valid .xlsx file (workbook part missing)')

  const workbook = parseXmlTree(workbookXml)
  const workbookPr = findFirst(workbook, 'workbookPr')
  const date1904 = ['1', 'true'].includes(workbookPr?.attrs.date1904)

  const relsXml = await readText(relsPathFor(workbookPath))
  const rels = new Map()
  let stylesPath = null
  let sharedStringsPath = null
  if (relsXml) {
    for (const r of findAll(parseXmlTree(relsXml), 'Relationship')) {
      const target = resolvePath(workbookPath, r.attrs.Target)
      rels.set(r.attrs.Id, target)
      if (/\/styles$/.test(r.attrs.Type)) stylesPath = target
      if (/\/sharedStrings$/.test(r.attrs.Type)) sharedStringsPath = target
    }
  }

  const sheets = findAll(workbook, 'sheet').map((s, i) => ({
    name: s.attrs.name,
    path: rels.get(s.attrs['r:id'] ?? s.attrs.id) ?? `xl/worksheets/sheet${i + 1}.xml`,
    state: s.attrs.state || 'visible'
  }))

  let shared = null
  let styles = null
  return {
    sheets,
    date1904,
    readText,
    async sharedStrings() {
      if (!shared) {
        const xml = sharedStringsPath ? await readText(sharedStringsPath) : null
        shared = xml ? parseSharedStrings(xml) : []
      }
      return shared
    },
    async dateStyles() {
      if (!styles) {
        const xml = stylesPath ? await readText(stylesPath) : null
        styles = xml ? parseDateStyles(xml) : []
      }
      return styles
    }
  }
}

function parseSharedStrings(xml) {
  const out = []
  let inSi = false
  let inT = false
  let phonetic = 0
  let buf = ''
  parseXml(xml, {
    open(name) {
      if (name === 'si') {
        inSi = true
        buf = ''
      } else if (name === 'rPh') phonetic++
      else if (name === 't' && inSi && !phonetic) inT = true
    },
    close(name) {
      if (name === 'si') {
        out.push(buf)
        inSi = false
      } else if (name === 'rPh') phonetic--
      else if (name === 't') inT = false
    },
    text(t) {
      if (inT) buf += t
    }
  })
  return out
}

/** Returns an array indexed by cellXfs style id: true when it's a date format. */
function parseDateStyles(xml) {
  const tree = parseXmlTree(xml)
  const custom = new Map()
  for (const f of findAll(tree, 'numFmt')) {
    custom.set(Number(f.attrs.numFmtId), f.attrs.formatCode)
  }
  const cellXfs = findFirst(tree, 'cellXfs')
  if (!cellXfs) return []
  return cellXfs.children
    .filter((x) => x.name === 'xf')
    .map((xf) => {
      const id = Number(xf.attrs.numFmtId || 0)
      return isDateFormat(id, custom.get(id))
    })
}

function pickSheet(wb, sheet) {
  if (sheet == null) {
    const visible = wb.sheets.find((s) => s.state === 'visible')
    return visible || wb.sheets[0]
  }
  if (typeof sheet === 'number') {
    const s = wb.sheets[sheet - 1]
    if (!s) throw new Error(`Sheet #${sheet} not found (workbook has ${wb.sheets.length})`)
    return s
  }
  const s = wb.sheets.find((x) => x.name === sheet)
  if (!s) {
    throw new Error(
      `Sheet "${sheet}" not found. Available: ${wb.sheets.map((x) => x.name).join(', ')}`
    )
  }
  return s
}

async function readSheetRows(wb, sheetInfo, options = {}) {
  const xml = await wb.readText(sheetInfo.path)
  if (xml == null) throw new Error(`Worksheet part missing: ${sheetInfo.path}`)
  const sst = await wb.sharedStrings()
  const dateStyles = await wb.dateStyles()
  const parseDates = options.dates !== false
  const date1904 = wb.date1904

  const rows = []
  let maxCol = -1
  let rowIdx = -1
  let colCursor = -1
  let cell = null
  let collect = null // 'v' | 't'
  let inIs = false
  let phonetic = 0
  let buf = ''
  const merges = []

  parseXml(xml, {
    open(name, a) {
      switch (name) {
        case 'row':
          rowIdx = a.r ? Number(a.r) - 1 : rowIdx + 1
          colCursor = -1
          break
        case 'c': {
          let col = colCursor + 1
          let r = rowIdx
          if (a.r) {
            const ref = parseCellRef(a.r)
            col = ref.col
            if (ref.row >= 0) r = ref.row
          }
          colCursor = col
          cell = { row: r, col, t: a.t || 'n', s: a.s ? Number(a.s) : 0, v: null, is: '' }
          break
        }
        case 'v':
          if (cell) {
            collect = 'v'
            buf = ''
          }
          break
        case 'is':
          inIs = true
          break
        case 'rPh':
          phonetic++
          break
        case 't':
          if (inIs && !phonetic) {
            collect = 't'
            buf = ''
          }
          break
        case 'mergeCell':
          if (a.ref) merges.push(a.ref)
          break
      }
    },
    close(name) {
      switch (name) {
        case 'v':
          if (collect === 'v') cell.v = buf
          collect = null
          break
        case 't':
          if (collect === 't') cell.is += buf
          collect = null
          break
        case 'is':
          inIs = false
          break
        case 'rPh':
          phonetic--
          break
        case 'c': {
          const value = cellValue(cell, sst, dateStyles, parseDates, date1904)
          if (value !== null && cell.row >= 0) {
            ;(rows[cell.row] ||= [])[cell.col] = value
            if (cell.col > maxCol) maxCol = cell.col
          }
          cell = null
          break
        }
      }
    },
    text(t) {
      if (collect) buf += t
    }
  })

  // Densify: fill holes with null and pad every row to the same width.
  const width = maxCol + 1
  const out = new Array(rows.length)
  for (let r = 0; r < rows.length; r++) {
    const src = rows[r] || []
    const row = new Array(width)
    for (let c = 0; c < width; c++) row[c] = src[c] ?? null
    out[r] = row
  }
  if (options.trim !== false) trimEmpty(out)
  return { rows: out, merges }
}

function trimEmpty(rows) {
  const isEmpty = (v) => v === null || v === ''
  while (rows.length && rows[rows.length - 1].every(isEmpty)) rows.pop()
  if (!rows.length) return
  let width = rows[0].length
  while (width > 0 && rows.every((r) => isEmpty(r[width - 1]))) width--
  for (const r of rows) r.length = width
}

function cellValue(cell, sst, dateStyles, parseDates, date1904) {
  const { t, v } = cell
  switch (t) {
    case 's':
      return v == null ? null : (sst[Number(v)] ?? null)
    case 'inlineStr':
      return cell.is
    case 'str':
      return v ?? ''
    case 'b':
      return v == null ? null : v === '1' || v === 'true'
    case 'e':
      return v // e.g. "#N/A", "#DIV/0!"
    case 'd':
      return v == null ? null : new Date(v)
    default: {
      if (v == null || v === '') return null
      const n = Number(v)
      if (!Number.isFinite(n)) return v
      if (parseDates && dateStyles[cell.s]) return excelSerialToDate(n, date1904)
      return n
    }
  }
}

// ---------------- Public API ----------------

/**
 * Reads one sheet as an array of rows (each row an array of cell values).
 * With `options.schema`, returns `{ rows: object[], errors: [] }` instead.
 */
export async function readXlsx(input, options = {}) {
  const wb = await openWorkbook(input)
  if (!wb.sheets.length) throw new Error('Workbook has no sheets')
  const { rows } = await readSheetRows(wb, pickSheet(wb, options.sheet), options)
  if (options.schema) {
    return parseWithSchema(rows, options.schema, { ...options, date1904: wb.date1904 })
  }
  return rows
}

/** Returns the names of all sheets, in workbook order. */
export async function readSheetNames(input) {
  const wb = await openWorkbook(input)
  return wb.sheets.map((s) => s.name)
}

/** Reads every sheet: `[{ name, data }]`. */
export async function readXlsxSheets(input, options = {}) {
  const wb = await openWorkbook(input)
  const out = []
  for (const s of wb.sheets) {
    const { rows } = await readSheetRows(wb, s, options)
    out.push({ name: s.name, data: options.schema
        ? parseWithSchema(rows, options.schema, { ...options, date1904: wb.date1904 })
        : rows })
  }
  return out
}
