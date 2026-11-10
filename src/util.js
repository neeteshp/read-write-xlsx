// Shared helpers: cell references and Excel date serials.

const MS_PER_DAY = 86400000

/** 0-based column index -> letters ("A", "B", ..., "AA"). */
export function columnLetter(index) {
  let n = index + 1
  let s = ''
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/** "C12" -> { col: 2, row: 11 } (0-based). */
export function parseCellRef(ref) {
  let col = 0
  let i = 0
  for (; i < ref.length; i++) {
    const c = ref.charCodeAt(i)
    if (c >= 65 && c <= 90) col = col * 26 + (c - 64)
    else if (c >= 97 && c <= 122) col = col * 26 + (c - 96)
    else if (c === 36 /* $ */) continue
    else break
  }
  const row = parseInt(ref.slice(i).replace('$', ''), 10)
  return { col: col - 1, row: Number.isFinite(row) ? row - 1 : -1 }
}

export const cellRef = (row, col) => columnLetter(col) + (row + 1)

/** Excel serial number -> JS Date (UTC). Handles the 1900 leap-year bug. */
export function excelSerialToDate(serial, date1904 = false) {
  let days
  if (date1904) days = serial - 24107
  else days = serial < 60 ? serial - 25568 : serial - 25569
  return new Date(Math.round(days * MS_PER_DAY))
}

/** JS Date -> Excel serial number (uses the date's UTC fields). */
export function dateToExcelSerial(date, date1904 = false) {
  const days = date.getTime() / MS_PER_DAY
  if (date1904) return days + 24107
  const serial = days + 25569
  return serial < 61 ? serial - 1 : serial
}

// Built-in number format ids that represent dates/times.
const BUILTIN_DATE_FORMATS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50,
  51, 52, 53, 54, 55, 56, 57, 58
])

export function isDateFormat(numFmtId, formatCode) {
  if (formatCode == null) return BUILTIN_DATE_FORMATS.has(numFmtId)
  let f = String(formatCode).split(';')[0]
  if (/\[(h+|m+|s+)\]/i.test(f)) return true // elapsed time, e.g. [h]:mm
  f = f
    .replace(/"[^"]*"/g, '') // quoted literals
    .replace(/\\./g, '') // escaped chars
    .replace(/\[[^\]]*\]/g, '') // [Red], [$-409] etc.
    .replace(/_.|\*./g, '') // padding/repeat
  return /[dmyhs]/i.test(f) && !/^general$/i.test(f.trim())
}
