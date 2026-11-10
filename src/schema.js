import { excelSerialToDate } from './util.js'

/** Marker type for whole numbers: `{ type: Integer }`. */
export function Integer(value) {
  const n = parseNumber(value)
  if (!Number.isInteger(n)) throw new Error('invalid')
  return n
}

function parseNumber(value) {
  if (typeof value === 'number') return value
  if (typeof value === 'string') {
    const s = value.trim()
    if (s !== '') {
      const n = Number(s)
      if (Number.isFinite(n)) return n
    }
  }
  throw new Error('invalid')
}

function convert(value, type, date1904) {
  if (type === undefined || type === null) return value
  if (type === String) {
    if (value instanceof Date) return value.toISOString()
    return String(value)
  }
  if (type === Number) return parseNumber(value)
  if (type === 'Integer') return Integer(value)
  if (type === Boolean) {
    if (typeof value === 'boolean') return value
    if (value === 1 || value === 0) return value === 1
    if (typeof value === 'string') {
      const s = value.trim().toLowerCase()
      if (s === 'true' || s === '1') return true
      if (s === 'false' || s === '0') return false
    }
    throw new Error('invalid')
  }
  if (type === Date) {
    if (value instanceof Date) return value
    if (typeof value === 'number') return excelSerialToDate(value, date1904)
    if (typeof value === 'string') {
      const t = Date.parse(value.trim())
      if (!Number.isNaN(t)) return new Date(t)
    }
    throw new Error('invalid')
  }
  if (typeof type === 'function') return type(value)
  throw new Error(`Unsupported schema type: ${String(type)}`)
}

const isEmpty = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '')

/**
 * Converts sheet rows into objects using a schema keyed by column title:
 *   { 'Name': { prop: 'name', type: String, required: true } }
 * @returns {{ rows: object[], errors: {row: number, column: string, value: any, error: string}[] }}
 */
export function parseWithSchema(sheetRows, schema, options = {}) {
  const errors = []
  const result = []
  const headerIndex = sheetRows.findIndex((r) => r.some((v) => !isEmpty(v)))
  if (headerIndex === -1) return { rows: result, errors }

  const header = sheetRows[headerIndex].map((h) => (h == null ? '' : String(h).trim()))
  const columns = Object.entries(schema).map(([title, def]) => {
    const spec = typeof def === 'string' ? { prop: def } : def
    return { title, index: header.indexOf(title.trim()), ...spec, prop: spec.prop ?? title }
  })
  for (const col of columns) {
    if (col.index === -1 && col.required) {
      errors.push({ row: headerIndex + 1, column: col.title, value: undefined, error: 'missing column' })
    }
  }

  for (let r = headerIndex + 1; r < sheetRows.length; r++) {
    const row = sheetRows[r]
    if (row.every(isEmpty)) continue
    const obj = {}
    let hasValue = false
    for (const col of columns) {
      const raw = col.index === -1 ? null : row[col.index]
      if (isEmpty(raw)) {
        if (col.required) errors.push({ row: r + 1, column: col.title, value: raw, error: 'required' })
        continue
      }
      let value
      try {
        value = convert(raw, col.type, options.date1904)
        if (col.oneOf && !col.oneOf.includes(value)) throw new Error('invalid')
        if (col.validate) col.validate(value)
      } catch (e) {
        errors.push({ row: r + 1, column: col.title, value: raw, error: e?.message || 'invalid' })
        continue
      }
      setPath(obj, col.prop, value)
      hasValue = true
    }
    if (hasValue) result.push(obj)
  }
  return { rows: result, errors }
}

/** Supports dotted props: "address.city". */
function setPath(obj, path, value) {
  const keys = String(path).split('.')
  let o = obj
  for (let i = 0; i < keys.length - 1; i++) o = o[keys[i]] ??= {}
  o[keys[keys.length - 1]] = value
}
