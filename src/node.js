// Node.js entry: same API as the universal entry, plus file-path helpers.
// Uses the built-in node:zlib for faster (de)compression.
import { readFile, writeFile } from 'node:fs/promises'
import { inflateRawSync, deflateRawSync } from 'node:zlib'
import { setZipCodec } from './zip.js'
import { readXlsx, readSheetNames, readXlsxSheets } from './read.js'
import { writeXlsx, writeXlsxSheets } from './write.js'

setZipCodec({
  inflateRaw: (u8) => inflateRawSync(u8),
  deflateRaw: (u8) => deflateRawSync(u8, { level: 6 })
})

async function streamToBuffer(stream) {
  const chunks = []
  for await (const chunk of stream) chunks.push(chunk)
  return Buffer.concat(chunks)
}

async function resolveInput(input) {
  if (typeof input === 'string' || input instanceof URL) return readFile(input)
  if (input && typeof input.pipe === 'function' && typeof input[Symbol.asyncIterator] === 'function') {
    return streamToBuffer(input)
  }
  return input
}

/** Reads a sheet from a file path, Buffer, or readable stream. */
export async function readXlsxFile(input, options) {
  return readXlsx(await resolveInput(input), options)
}

export async function readSheetNamesFromFile(input) {
  return readSheetNames(await resolveInput(input))
}

export async function readXlsxFileSheets(input, options) {
  return readXlsxSheets(await resolveInput(input), options)
}

/** Writes a workbook. With `options.filePath` saves it to disk; always returns a Buffer. */
export async function writeXlsxFile(data, options = {}) {
  const { filePath, ...rest } = options
  const buffer = Buffer.from(await writeXlsx(data, rest))
  if (filePath) await writeFile(filePath, buffer)
  return buffer
}

export async function writeXlsxFileSheets(sheets, options = {}) {
  const { filePath, ...rest } = options
  const buffer = Buffer.from(await writeXlsxSheets(sheets, rest))
  if (filePath) await writeFile(filePath, buffer)
  return buffer
}

export * from './index.js'
