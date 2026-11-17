// Universal entry (browsers, Node >= 18, Deno, Bun, edge runtimes).
import { readXlsx, readSheetNames, readXlsxSheets } from './read.js'
import { writeXlsx, writeXlsxSheets } from './write.js'

export { readXlsx, readSheetNames, readXlsxSheets, writeXlsx, writeXlsxSheets }
export { Integer } from './schema.js'
export { excelSerialToDate, dateToExcelSerial } from './util.js'
export { setZipCodec } from './zip.js'

// The *File functions have the same names as in Node, so `import { readXlsxFile } from 'read-write-xlsx'`
// works everywhere. Outside Node they take a File/Blob/ArrayBuffer/Uint8Array instead of a path.
function noPaths(input) {
  if (typeof input === 'string' || input instanceof URL) {
    throw new TypeError(
      'File paths are only supported in Node.js. In the browser, pass a File, Blob, ArrayBuffer or Uint8Array.'
    )
  }
  return input
}

function noFilePath(options) {
  const { filePath, ...rest } = options
  if (filePath) {
    throw new TypeError(
      'The filePath option is only supported in Node.js. In the browser, use the returned bytes, e.g. to create a Blob for download.'
    )
  }
  return rest
}

export const readXlsxFile = async (input, options) => readXlsx(noPaths(input), options)
export const readSheetNamesFromFile = async (input) => readSheetNames(noPaths(input))
export const readXlsxFileSheets = async (input, options) => readXlsxSheets(noPaths(input), options)
export const writeXlsxFile = async (data, options = {}) => writeXlsx(data, noFilePath(options))
export const writeXlsxFileSheets = async (sheets, options = {}) => writeXlsxSheets(sheets, noFilePath(options))
