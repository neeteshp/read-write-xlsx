import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, readFile } from 'node:fs/promises'
import { createReadStream } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  readXlsxFile,
  readSheetNamesFromFile,
  readXlsxFileSheets,
  writeXlsxFile,
  writeXlsxFileSheets,
  readXlsx,
  Integer
} from '../src/node.js'

const fixture = (name) => new URL(`./fixtures/${name}`, import.meta.url)
const utc = (...a) => new Date(Date.UTC(...a))

test('reads an openpyxl workbook: types, dates, rich text, sparse cells', async () => {
  const rows = await readXlsxFile(fixture('openpyxl.xlsx'))
  assert.deepEqual(rows[0].slice(0, 6), ['Name', 'Age', 'Born', 'Active', 'Score', 'Time'])
  assert.deepEqual(rows[1].slice(0, 5), ['Alice', 34, utc(1992, 4, 17), true, 0.875])
  assert.deepEqual(rows[2].slice(0, 4), ['Bob', '41', utc(1985, 0, 2, 8, 15), false])
  assert.equal(rows[3][0], 'Carol') // rich text runs are concatenated
  assert.equal(rows.length, 10)
  assert.equal(rows[9][7], 'far away')
  assert.ok(rows.every((r) => r.length === 8))
})

test('reads a LibreOffice workbook including cached formula errors', async () => {
  const rows = await readXlsxFile(fixture('libreoffice.xlsx'))
  assert.deepEqual(rows[1].slice(0, 4), ['Alice', 34, utc(1992, 4, 17), true])
  assert.equal(rows[3][4], '#DIV/0!')
})

test('1904 date system', async () => {
  assert.deepEqual(await readXlsxFile(fixture('date1904.xlsx')), [['Date'], [utc(2020, 1, 29)]])
})

test('sheet names and sheet selection by name / number', async () => {
  const names = await readSheetNamesFromFile(fixture('openpyxl.xlsx'))
  assert.deepEqual(names, ['People', 'Hidden stuff', 'Ünïcødé & <xml>'])
  assert.deepEqual(await readXlsxFile(fixture('openpyxl.xlsx'), { sheet: 2 }), [['second sheet']])
  assert.deepEqual(await readXlsxFile(fixture('openpyxl.xlsx'), { sheet: 'Ünïcødé & <xml>' }), [
    [null, null],
    [null, 'ok']
  ])
  await assert.rejects(readXlsxFile(fixture('openpyxl.xlsx'), { sheet: 'Nope' }), /not found/)
  const all = await readXlsxFileSheets(fixture('openpyxl.xlsx'))
  assert.equal(all.length, 3)
})

test('accepts Buffer and stream input', async () => {
  const buf = await readFile(fixture('date1904.xlsx'))
  assert.equal((await readXlsx(buf)).length, 2)
  assert.equal((await readXlsxFile(createReadStream(fixture('date1904.xlsx')))).length, 2)
})

test('schema parsing with types, required, oneOf, custom parser and errors', async () => {
  const { rows, errors } = await readXlsxFile(fixture('openpyxl.xlsx'), {
    schema: {
      Name: { prop: 'name', type: String, required: true },
      Age: { prop: 'age', type: Integer },
      Born: { prop: 'born', type: Date },
      Active: { prop: 'flags.active', type: Boolean },
      Score: { prop: 'score', type: Number },
      Missing: { prop: 'missing', required: false }
    }
  })
  assert.deepEqual(rows[0], {
    name: 'Alice',
    age: 34,
    born: utc(1992, 4, 17),
    flags: { active: true },
    score: 0.875
  })
  assert.equal(rows[1].age, 41) // numeric string converted
  assert.deepEqual(errors, [{ row: 10, column: 'Name', value: null, error: 'required' }])
})

test('round trip: write then read back', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'xlsx-'))
  const filePath = join(dir, 'out.xlsx')
  const data = [
    ['Text', 'Number', 'Bool', 'Date', 'Empty'],
    ['Äpfel & <Birnen> "quoted"', 42.5, true, utc(2026, 10, 10, 9, 30), null],
    ['  leading/trailing  ', -1e-7, false, utc(1900, 0, 1), ''],
    ['line\nbreak', 1234567890123, null, utc(1900, 2, 1), null],
    [{ value: '=SUM(B2:B3)', type: 'Formula' }, { value: '7', type: Number }, null, null, null]
  ]
  const buffer = await writeXlsxFile(data, { filePath, sheet: 'Data' })
  assert.ok(Buffer.isBuffer(buffer))
  const rows = await readXlsxFile(filePath)
  assert.deepEqual(
    rows.slice(0, 4).map((r) => r.slice(0, 4)),
    data.slice(0, 4).map((r) => r.slice(0, 4))
  )
  assert.deepEqual(rows[0][4], 'Empty')
  assert.equal(rows[4][1], 7)
  assert.equal(rows[4][0], null) // formulas have no cached value until opened in Excel
  assert.deepEqual(await readSheetNamesFromFile(filePath), ['Data'])
})

test('writes from schema and reads back with the same column titles', async () => {
  const people = [
    { name: 'Alice', age: 34, born: utc(1992, 4, 17), active: true },
    { name: 'Bob', age: 41, born: utc(1985, 0, 2), active: false }
  ]
  const buffer = await writeXlsxFile(people, {
    schema: [
      { column: 'Name', type: String, value: (p) => p.name, width: 20 },
      { column: 'Age', type: Number, value: (p) => p.age, align: 'right' },
      { column: 'Born', type: Date, value: (p) => p.born, format: 'dd.mm.yyyy' },
      { column: 'Active', type: Boolean, value: (p) => p.active }
    ]
  })
  const { rows, errors } = await readXlsx(buffer, {
    schema: {
      Name: { prop: 'name', type: String },
      Age: { prop: 'age', type: Number },
      Born: { prop: 'born', type: Date },
      Active: { prop: 'active', type: Boolean }
    }
  })
  assert.deepEqual(errors, [])
  assert.deepEqual(rows, people)
})

test('multiple sheets with sanitised, unique names', async () => {
  const buffer = await writeXlsxFileSheets([
    { name: 'Q1/Q2: [Report]', data: [['a']] },
    { name: 'Q1/Q2: [Report]', data: [['b']] },
    { data: [['c']] }
  ])
  const names = await readSheetNamesFromFile(buffer)
  assert.deepEqual(names, ['Q1_Q2_ _Report_', 'Q1_Q2_ _Report_ (2)', 'Sheet3'])
  assert.deepEqual(await readXlsx(buffer, { sheet: 3 }), [['c']])
})

test('invalid input is rejected with a clear message', async () => {
  await assert.rejects(readXlsx(new Uint8Array([1, 2, 3])), /Not a valid \.xlsx/)
  await assert.rejects(readXlsx(42), /Unsupported input/)
  await assert.rejects(writeXlsxFile([[{ value: 'abc', type: Number }]]), /not a finite number/)
  await assert.rejects(writeXlsxFile([[{ value: 'x', color: 'blue' }]]), /Invalid color/)
})
