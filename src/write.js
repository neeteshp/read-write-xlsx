import { writeZip } from './zip.js'
import { escapeXml } from './xml.js'
import { cellRef, columnLetter, dateToExcelSerial } from './util.js'

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

const BUILTIN_NUM_FORMATS = {
  General: 0,
  '0': 1,
  '0.00': 2,
  '#,##0': 3,
  '#,##0.00': 4,
  '0%': 9,
  '0.00%': 10,
  '0.00E+00': 11,
  'mm-dd-yy': 14,
  'h:mm': 20,
  'h:mm:ss': 21,
  'm/d/yy h:mm': 22,
  '@': 49
}

const ALIGN = new Set(['left', 'center', 'right', 'fill', 'justify', 'general'])
const VALIGN = { top: 'top', center: 'center', middle: 'center', bottom: 'bottom' }
const BORDER_STYLES = new Set([
  'thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair',
  'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'slantDashDot'
])

function argb(color) {
  if (!color) return null
  let c = String(color).replace('#', '').trim()
  if (/^[0-9a-f]{3}$/i.test(c)) c = c.replace(/./g, (x) => x + x)
  if (/^[0-9a-f]{6}$/i.test(c)) return 'FF' + c.toUpperCase()
  if (/^[0-9a-f]{8}$/i.test(c)) return c.toUpperCase()
  throw new Error(`Invalid color "${color}". Use "#RRGGBB".`)
}

/** Collects unique fonts/fills/borders/number formats and the cell formats combining them. */
class StyleRegistry {
  constructor({ fontFamily = 'Calibri', fontSize = 11 } = {}) {
    this.defaultFont = { name: fontFamily, size: fontSize }
    this.fonts = [this.defaultFont]
    this.fills = [{ pattern: 'none' }, { pattern: 'gray125' }]
    this.borders = [{}]
    this.numFmts = [] // { id, code }
    this.xfs = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }]
    this.keys = new Map([['font:' + JSON.stringify(this.defaultFont), 0]])
  }

  intern(list, item, prefix) {
    const key = prefix + JSON.stringify(item)
    let id = this.keys.get(key)
    if (id === undefined) {
      id = list.push(item) - 1
      this.keys.set(key, id)
    }
    return id
  }

  numFmtId(code) {
    if (code == null) return 0
    if (code in BUILTIN_NUM_FORMATS) return BUILTIN_NUM_FORMATS[code]
    const found = this.numFmts.find((f) => f.code === code)
    if (found) return found.id
    const id = 164 + this.numFmts.length
    this.numFmts.push({ id, code })
    return id
  }

  /** Returns the cellXfs index for a cell's style properties (0 = default). */
  styleId(c, format) {
    const font = {
      name: c.fontFamily ?? this.defaultFont.name,
      size: c.fontSize ?? this.defaultFont.size,
      bold: c.fontWeight === 'bold' || undefined,
      italic: c.fontStyle === 'italic' || undefined,
      underline: c.textDecoration === 'underline' || undefined,
      strike: c.textDecoration === 'line-through' || undefined,
      color: argb(c.color) || undefined
    }
    const fill = c.backgroundColor ? { pattern: 'solid', color: argb(c.backgroundColor) } : null
    const border = this.borderFor(c)
    const alignment = {
      horizontal: c.align && ALIGN.has(c.align) ? c.align : undefined,
      vertical: c.alignVertical ? VALIGN[c.alignVertical] : undefined,
      wrapText: c.wrap ? true : undefined,
      indent: c.indent || undefined,
      textRotation: c.textRotation || undefined
    }
    const hasAlignment = Object.values(alignment).some((v) => v !== undefined)

    const xf = {
      numFmtId: this.numFmtId(format),
      fontId: this.intern(this.fonts, font, 'font:'),
      fillId: fill ? this.intern(this.fills, fill, 'fill:') : 0,
      borderId: border ? this.intern(this.borders, border, 'border:') : 0,
      alignment: hasAlignment ? alignment : undefined
    }
    if (
      xf.numFmtId === 0 && xf.fontId === 0 && xf.fillId === 0 && xf.borderId === 0 && !hasAlignment
    ) {
      return 0
    }
    return this.intern(this.xfs, xf, 'xf:')
  }

  borderFor(c) {
    const sides = {}
    const base = c.borderStyle ?? (c.borderColor ? 'thin' : null)
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const cap = side[0].toUpperCase() + side.slice(1)
      const style = c[`border${cap}Style`] ?? (c[`border${cap}Color`] ? 'thin' : base)
      if (!style) continue
      if (!BORDER_STYLES.has(style)) throw new Error(`Invalid border style "${style}"`)
      sides[side] = { style, color: argb(c[`border${cap}Color`] ?? c.borderColor) || undefined }
    }
    return Object.keys(sides).length ? sides : null
  }

  toXml() {
    const fontXml = (f) =>
      '<font>' +
      (f.bold ? '<b/>' : '') +
      (f.italic ? '<i/>' : '') +
      (f.strike ? '<strike/>' : '') +
      (f.underline ? '<u/>' : '') +
      `<sz val="${f.size}"/>` +
      (f.color ? `<color rgb="${f.color}"/>` : '') +
      `<name val="${escapeXml(f.name)}"/><family val="2"/>` +
      '</font>'
    const fillXml = (f) =>
      f.pattern === 'solid'
        ? `<fill><patternFill patternType="solid"><fgColor rgb="${f.color}"/><bgColor indexed="64"/></patternFill></fill>`
        : `<fill><patternFill patternType="${f.pattern}"/></fill>`
    const borderXml = (b) =>
      '<border>' +
      ['left', 'right', 'top', 'bottom']
        .map((s) =>
          b[s]
            ? `<${s} style="${b[s].style}">` +
              (b[s].color ? `<color rgb="${b[s].color}"/>` : '<color indexed="64"/>') +
              `</${s}>`
            : `<${s}/>`
        )
        .join('') +
      '<diagonal/></border>'
    const xfXml = (x) => {
      let attrs = `numFmtId="${x.numFmtId}" fontId="${x.fontId}" fillId="${x.fillId}" borderId="${x.borderId}" xfId="0"`
      if (x.numFmtId) attrs += ' applyNumberFormat="1"'
      if (x.fontId) attrs += ' applyFont="1"'
      if (x.fillId) attrs += ' applyFill="1"'
      if (x.borderId) attrs += ' applyBorder="1"'
      if (!x.alignment) return `<xf ${attrs}/>`
      const a = x.alignment
      const al =
        (a.horizontal ? ` horizontal="${a.horizontal}"` : '') +
        (a.vertical ? ` vertical="${a.vertical}"` : '') +
        (a.wrapText ? ' wrapText="1"' : '') +
        (a.indent ? ` indent="${a.indent}"` : '') +
        (a.textRotation ? ` textRotation="${a.textRotation}"` : '')
      return `<xf ${attrs} applyAlignment="1"><alignment${al}/></xf>`
    }

    return (
      XML_HEADER +
      `<styleSheet xmlns="${NS_MAIN}">` +
      (this.numFmts.length
        ? `<numFmts count="${this.numFmts.length}">` +
          this.numFmts
            .map((f) => `<numFmt numFmtId="${f.id}" formatCode="${escapeXml(f.code)}"/>`)
            .join('') +
          '</numFmts>'
        : '') +
      `<fonts count="${this.fonts.length}">${this.fonts.map(fontXml).join('')}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.map(fillXml).join('')}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.map(borderXml).join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.map(xfXml).join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/><tableStyles count="0"/>' +
      '</styleSheet>'
    )
  }
}

class SharedStrings {
  constructor() {
    this.list = []
    this.index = new Map()
    this.count = 0
  }
  add(s) {
    this.count++
    let i = this.index.get(s)
    if (i === undefined) {
      i = this.list.push(s) - 1
      this.index.set(s, i)
    }
    return i
  }
  toXml() {
    const items = this.list.map((s) => {
      const preserve = /^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ''
      return `<si><t${preserve}>${escapeXml(s)}</t></si>`
    })
    return (
      XML_HEADER +
      `<sst xmlns="${NS_MAIN}" count="${this.count}" uniqueCount="${this.list.length}">` +
      items.join('') +
      '</sst>'
    )
  }
}

const isCellObject = (c) =>
  c !== null && typeof c === 'object' && !(c instanceof Date) && !Array.isArray(c)

function inferType(value) {
  if (typeof value === 'number') return Number
  if (typeof value === 'boolean') return Boolean
  if (value instanceof Date) return Date
  return String
}

function displayLength(value, type, format) {
  if (value == null) return 0
  if (type === Date) return (format || 'yyyy-mm-dd').length
  return String(value)
    .split('\n')
    .reduce((m, line) => Math.max(m, line.length), 0)
}

/** Turns a schema + objects into plain rows of cell objects. */
function rowsFromSchema(objects, schema, headerStyle) {
  const header = schema.map((col) => ({
    value: col.column ?? '',
    ...headerStyle,
    ...(col.headerStyle || {})
  }))
  const body = objects.map((obj) =>
    schema.map((col) => {
      const { column, value, width, headerStyle: _h, getCellStyle, ...style } = col
      const v = typeof value === 'function' ? value(obj) : obj[column]
      const extra = typeof getCellStyle === 'function' ? getCellStyle(obj) : null
      return { ...style, ...(extra || {}), value: v }
    })
  )
  return [header, ...body]
}

function sheetXml(rows, sheetOpts, ctx) {
  const { styles, sst, dateFormat } = ctx
  const merges = []
  const widths = []
  const rowHeights = []
  let maxCol = 0
  const rowParts = []

  rows.forEach((row, r) => {
    if (!row || !row.length) return
    const cells = []
    row.forEach((raw, c) => {
      const cell = isCellObject(raw) ? raw : { value: raw }
      if (cell.height) rowHeights[r] = Math.max(rowHeights[r] || 0, cell.height)
      const span = cell.columnSpan ?? cell.span ?? 1
      const rowSpan = cell.rowSpan ?? 1
      if (span > 1 || rowSpan > 1) {
        merges.push(`${cellRef(r, c)}:${cellRef(r + rowSpan - 1, c + span - 1)}`)
      }
      let { value } = cell
      if (value === undefined || value === null || value === '') {
        // Still emit styled empty cells (e.g. a coloured blank in a merged range).
        const s = isCellObject(raw) ? styles.styleId(cell, cell.format) : 0
        if (s) cells.push(`<c r="${cellRef(r, c)}" s="${s}"/>`)
        return
      }
      if (c + 1 > maxCol) maxCol = c + 1

      const type = cell.type ?? inferType(value)
      let format = cell.format
      let body
      const ref = cellRef(r, c)

      if (type === 'Formula') {
        const f = String(value).replace(/^=/, '')
        body = { t: null, xml: `<f>${escapeXml(f)}</f>` }
      } else if (type === Number || type === 'Integer') {
        const n = typeof value === 'number' ? value : Number(value)
        if (!Number.isFinite(n)) throw new Error(`Cell ${ref}: "${value}" is not a finite number`)
        body = { t: null, xml: `<v>${n}</v>` }
      } else if (type === Boolean) {
        body = { t: 'b', xml: `<v>${value ? 1 : 0}</v>` }
      } else if (type === Date) {
        const d = value instanceof Date ? value : new Date(value)
        if (Number.isNaN(d.getTime())) throw new Error(`Cell ${ref}: invalid date "${value}"`)
        format = format || dateFormat
        body = { t: null, xml: `<v>${dateToExcelSerial(d)}</v>` }
      } else {
        body = { t: 's', xml: `<v>${sst.add(String(value))}</v>` }
      }

      const s = styles.styleId(cell, format)
      if (sheetOpts.autoWidth) {
        widths[c] = Math.max(widths[c] || 0, displayLength(value, type, format))
      }
      cells.push(
        `<c r="${ref}"${s ? ` s="${s}"` : ''}${body.t ? ` t="${body.t}"` : ''}>${body.xml}</c>`
      )
    })
    if (cells.length || rowHeights[r]) {
      const ht = rowHeights[r] ? ` ht="${rowHeights[r]}" customHeight="1"` : ''
      rowParts.push(`<row r="${r + 1}"${ht}>${cells.join('')}</row>`)
    }
  })

  // Column widths: explicit `columns` win over auto-width.
  const cols = []
  const colCount = Math.max(maxCol, sheetOpts.columns?.length || 0, widths.length)
  for (let c = 0; c < colCount; c++) {
    const explicit = sheetOpts.columns?.[c]?.width
    const auto = widths[c] ? Math.min(Math.max(widths[c] + 2, 8), 80) : null
    const w = explicit ?? auto
    if (w) cols.push(`<col min="${c + 1}" max="${c + 1}" width="${w}" customWidth="1"/>`)
  }

  const ySplit = sheetOpts.stickyRowsCount || 0
  const xSplit = sheetOpts.stickyColumnsCount || 0
  let pane = ''
  if (ySplit || xSplit) {
    const activePane = ySplit && xSplit ? 'bottomRight' : ySplit ? 'bottomLeft' : 'topRight'
    pane =
      `<pane${xSplit ? ` xSplit="${xSplit}"` : ''}${ySplit ? ` ySplit="${ySplit}"` : ''}` +
      ` topLeftCell="${cellRef(ySplit, xSplit)}" activePane="${activePane}" state="frozen"/>` +
      `<selection pane="${activePane}"/>`
  }
  const viewAttrs =
    (ctx.first ? ' tabSelected="1"' : '') +
    (sheetOpts.showGridLines === false ? ' showGridLines="0"' : '') +
    (sheetOpts.rightToLeft ? ' rightToLeft="1"' : '')

  const dim = maxCol && rows.length ? `A1:${columnLetter(maxCol - 1)}${rows.length}` : 'A1'
  return (
    XML_HEADER +
    `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
    `<dimension ref="${dim}"/>` +
    `<sheetViews><sheetView workbookViewId="0"${viewAttrs}>${pane}</sheetView></sheetViews>` +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    (cols.length ? `<cols>${cols.join('')}</cols>` : '') +
    `<sheetData>${rowParts.join('')}</sheetData>` +
    (merges.length
      ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`
      : '') +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '</worksheet>'
  )
}

function sanitizeSheetNames(names) {
  const used = new Set()
  return names.map((n, i) => {
    let name = String(n ?? `Sheet${i + 1}`).replace(/[[\]:*?/\\]/g, '_').trim().slice(0, 31)
    if (!name) name = `Sheet${i + 1}`
    let unique = name
    for (let k = 2; used.has(unique.toLowerCase()); k++) {
      const suffix = ` (${k})`
      unique = name.slice(0, 31 - suffix.length) + suffix
    }
    used.add(unique.toLowerCase())
    return unique
  })
}

/**
 * Writes several sheets into one workbook.
 * @param {{ name?: string, data: any[], schema?: any[], columns?: any[], stickyRowsCount?: number }[]} sheets
 * @param {object} [options] shared defaults (fontFamily, fontSize, dateFormat, autoWidth, ...)
 * @returns {Promise<Uint8Array>}
 */
export async function writeXlsxSheets(sheets, options = {}) {
  if (!Array.isArray(sheets) || !sheets.length) {
    throw new TypeError('writeXlsxSheets() expects a non-empty array of { name, data }')
  }
  const styles = new StyleRegistry(options)
  const sst = new SharedStrings()
  const dateFormat = options.dateFormat || 'yyyy-mm-dd'
  const names = sanitizeSheetNames(sheets.map((s) => s.name))
  const headerStyle = options.headerStyle ?? { fontWeight: 'bold' }

  const sheetFiles = sheets.map((s, i) => {
    const opts = { ...options, ...s }
    const schema = s.schema ?? (sheets.length === 1 ? options.schema : undefined)
    const rows = schema ? rowsFromSchema(s.data || [], schema, opts.headerStyle ?? headerStyle) : s.data || []
    if (!Array.isArray(rows)) throw new TypeError(`Sheet "${names[i]}": data must be an array`)
    if (schema) {
      if (!opts.columns) opts.columns = schema.map((c) => (c.width ? { width: c.width } : {}))
      if (opts.stickyRowsCount === undefined && opts.stickyHeader !== false) opts.stickyRowsCount = 1
    }
    return {
      name: `xl/worksheets/sheet${i + 1}.xml`,
      data: sheetXml(rows, opts, { styles, sst, dateFormat, first: i === 0 })
    }
  })

  const contentTypes =
    XML_HEADER +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    sheetFiles
      .map(
        (_, i) =>
          `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
      )
      .join('') +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    '</Types>'

  const rootRels =
    XML_HEADER +
    `<Relationships xmlns="${NS_PKG_REL}">` +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
    '</Relationships>'

  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z')
  const core =
    XML_HEADER +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    (options.title ? `<dc:title>${escapeXml(options.title)}</dc:title>` : '') +
    (options.author ? `<dc:creator>${escapeXml(options.author)}</dc:creator>` : '') +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
    '</cp:coreProperties>'
  const app =
    XML_HEADER +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>read-write-xlsx</Application></Properties>'

  const workbook =
    XML_HEADER +
    `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
    '<workbookPr/>' +
    '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16000" windowHeight="9000"/></bookViews>' +
    '<sheets>' +
    names.map((n, i) => `<sheet name="${escapeXml(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
    '</sheets>' +
    '<calcPr calcId="191029" fullCalcOnLoad="1"/>' +
    '</workbook>'

  const n = sheetFiles.length
  const workbookRels =
    XML_HEADER +
    `<Relationships xmlns="${NS_PKG_REL}">` +
    sheetFiles
      .map(
        (_, i) =>
          `<Relationship Id="rId${i + 1}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`
      )
      .join('') +
    `<Relationship Id="rId${n + 1}" Type="${NS_REL}/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId${n + 2}" Type="${NS_REL}/sharedStrings" Target="sharedStrings.xml"/>` +
    '</Relationships>'

  return writeZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'docProps/core.xml', data: core },
    { name: 'docProps/app.xml', data: app },
    { name: 'xl/workbook.xml', data: workbook },
    { name: 'xl/_rels/workbook.xml.rels', data: workbookRels },
    ...sheetFiles,
    { name: 'xl/styles.xml', data: styles.toXml() },
    { name: 'xl/sharedStrings.xml', data: sst.toXml() }
  ])
}

/**
 * Writes a single-sheet workbook.
 * `data` is an array of rows (arrays of cells), or an array of objects when `options.schema` is set.
 * @returns {Promise<Uint8Array>}
 */
export async function writeXlsx(data, options = {}) {
  const { sheet, ...rest } = options
  return writeXlsxSheets([{ name: sheet, data, schema: options.schema }], rest)
}
