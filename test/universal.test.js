// Runs the universal entry without node:zlib (uses CompressionStream),
// the same code path browsers and edge runtimes take.
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { writeXlsx, readXlsx, readSheetNames } from '../src/index.js'

test('universal entry: round trip via CompressionStream', async () => {
  const rows = Array.from({ length: 200 }, (_, i) => [`row ${i}`, i, i % 2 === 0])
  const u8 = await writeXlsx(rows, { sheet: 'Universal' })
  assert.ok(u8 instanceof Uint8Array)
  assert.deepEqual(await readSheetNames(u8), ['Universal'])
  assert.deepEqual(await readXlsx(u8), rows)
})

test('universal entry: reads Blob input and files from other producers', async () => {
  const bytes = await readFile(new URL('./fixtures/libreoffice.xlsx', import.meta.url))
  const rows = await readXlsx(new Blob([bytes]))
  assert.equal(rows[1][0], 'Alice')
})

test('universal entry: *File helpers take bytes and reject Node-only inputs', async () => {
  const { readXlsxFile, readSheetNamesFromFile, readXlsxFileSheets, writeXlsxFile, writeXlsxFileSheets } =
    await import('../src/index.js')
  const u8 = await writeXlsxFile([['a', 1]], { sheet: 'S' })
  assert.ok(u8 instanceof Uint8Array)
  assert.deepEqual(await readXlsxFile(u8), [['a', 1]])
  assert.deepEqual(await readSheetNamesFromFile(new Blob([u8])), ['S'])
  assert.equal((await readXlsxFileSheets(u8)).length, 1)
  assert.ok((await writeXlsxFileSheets([{ name: 'T', data: [['x']] }])) instanceof Uint8Array)

  await assert.rejects(readXlsxFile('report.xlsx'), /only supported in Node/)
  await assert.rejects(readSheetNamesFromFile(new URL('file:///x.xlsx')), /only supported in Node/)
  await assert.rejects(writeXlsxFile([['x']], { filePath: 'x.xlsx' }), /only supported in Node/)
})
