// Universal entry (browsers, Node >= 18, Deno, Bun, edge runtimes).
export { readXlsx, readSheetNames, readXlsxSheets } from './read.js'
export { writeXlsx, writeXlsxSheets } from './write.js'
export { Integer } from './schema.js'
export { excelSerialToDate, dateToExcelSerial } from './util.js'
export { setZipCodec } from './zip.js'
