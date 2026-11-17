# read-write-xlsx

Read **and** write `.xlsx` files in the browser and Node.js — with **zero dependencies**.

**[▶ Try the live demo](https://neeteshp.github.io/read-write-xlsx/demo/)**: read, write and validate spreadsheets right in your browser.

The API follows the two popular single-purpose packages
[`read-excel-file`](https://www.npmjs.com/package/read-excel-file) and
[`write-excel-file`](https://www.npmjs.com/package/write-excel-file), combined in one package.
It's a fresh implementation, not a copy of their source. Unzipping, zipping and XML parsing use only
what's built into the platform:

| Need        | Uses                                                                 |
| ----------- | -------------------------------------------------------------------- |
| Deflate     | `node:zlib` in Node; `CompressionStream('deflate-raw')` everywhere else |
| ZIP + CRC32 | ~200 lines in `src/zip.js`                                           |
| XML         | streaming tokenizer in `src/xml.js`                                  |

- ~70 KB of plain ES modules, no build step, no `node_modules` at runtime
- Works in Node ≥ 18, modern browsers, Deno, Bun and edge runtimes
- Reads files from Excel, LibreOffice, Google Sheets, openpyxl and more: shared/inline/rich-text
  strings, booleans, errors, dates (1900 and 1904 systems), sparse cells
- Writes styles (fonts, colours, fills, borders, alignment, number formats), merged cells, frozen panes,
  column widths, formulas, and multiple sheets
- Schema support in both directions: rows ⇄ typed objects with validation errors
- TypeScript types included

## Install

```bash
npm install read-write-xlsx
```

## Reading

```js
// Node.js (accepts a file path, URL, Buffer or readable stream)
import { readXlsxFile, readSheetNamesFromFile } from 'read-write-xlsx'

const rows = await readXlsxFile('report.xlsx')            // first visible sheet
const rows2 = await readXlsxFile('report.xlsx', { sheet: 'Q3' }) // or { sheet: 2 }
const names = await readSheetNamesFromFile('report.xlsx')
```

```js
// Browser (File, Blob, ArrayBuffer or Uint8Array)
import { readXlsx, readSheetNames } from 'read-write-xlsx'

input.addEventListener('change', async () => {
  const rows = await readXlsx(input.files[0])
})
```

`rows` is an array of rows, and each row is an array of cell values (`string | number | boolean | Date | null`).
Row `i` is spreadsheet row `i + 1`; empty cells are `null`; trailing empty rows/columns are trimmed
(`{ trim: false }` keeps them). Date-formatted cells come back as `Date` objects in UTC
(`{ dates: false }` returns the raw serial numbers instead).

Formula cells return the value Excel last calculated and saved. A file written by this library has no
saved value until someone opens and saves it in Excel or LibreOffice.

`readXlsxSheets(input)` / `readXlsxFileSheets(path)` read every sheet as `[{ name, data }]`.

### Reading with a schema

```js
import { readXlsxFile, Integer } from 'read-write-xlsx'

const { rows, errors } = await readXlsxFile('people.xlsx', {
  schema: {
    // column title in the header row → how to read it
    'Name':    { prop: 'name', type: String, required: true },
    'Age':     { prop: 'age', type: Integer },
    'Born':    { prop: 'born', type: Date },
    'Active':  { prop: 'flags.active', type: Boolean },     // dot paths create nested objects
    'Status':  { prop: 'status', oneOf: ['open', 'closed'] },
    'E-mail':  { prop: 'email', validate: (v) => { if (!v.includes('@')) throw new Error('invalid email') } },
    'Country': { prop: 'country', type: (v) => String(v).toUpperCase() } // custom parser
  }
})
// errors: [{ row: 7, column: 'Age', value: 'abc', error: 'invalid' }, ...]   (row = Excel row number)
```

Types: `String`, `Number`, `Integer`, `Boolean`, `Date`, or any function. Strings such as `"41"`
are converted when the type is `Number`. Rows that are completely empty are skipped. If a column
marked `required` is missing from the header, you get a `missing column` error.

## Writing

```js
// Node.js
import { writeXlsxFile } from 'read-write-xlsx'

await writeXlsxFile([
  ['Name', 'Amount', 'Paid', 'Due'],
  ['Alice', 1200.5, true, new Date('2026-12-01')],
  [{ value: 'Total', fontWeight: 'bold' }, { type: 'Formula', value: 'SUM(B2:B2)', format: '#,##0.00' }, null, null]
], {
  filePath: 'invoice.xlsx',  // omit to just get the Buffer back
  sheet: 'Invoices',
  stickyRowsCount: 1,
  autoWidth: true
})
```

```js
// Browser
import { writeXlsx } from 'read-write-xlsx'

const bytes = await writeXlsx(rows, { sheet: 'Export' })
const blob = new Blob([bytes], {
  type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
})
const a = Object.assign(document.createElement('a'), {
  href: URL.createObjectURL(blob),
  download: 'export.xlsx'
})
a.click()
URL.revokeObjectURL(a.href)
```

A cell is either a plain value (`string | number | boolean | Date | null`) or an object:

| Property | Example | Notes |
| --- | --- | --- |
| `value` | `42` | |
| `type` | `Number`, `String`, `Boolean`, `Date`, `'Integer'`, `'Formula'` | inferred when omitted |
| `format` | `'#,##0.00'`, `'0%'`, `'dd.mm.yyyy'` | Excel number format |
| `fontWeight` / `fontStyle` / `textDecoration` | `'bold'` / `'italic'` / `'underline'` \| `'line-through'` | |
| `fontFamily` / `fontSize` | `'Arial'` / `12` | |
| `color` / `backgroundColor` | `'#C00000'` / `'#FFF2CC'` | |
| `align` / `alignVertical` / `wrap` | `'center'` / `'top'` / `true` | |
| `borderStyle` / `borderColor` | `'thin'` / `'#999999'` | also `borderTopStyle`, `borderLeftColor`, … |
| `columnSpan` / `rowSpan` | `3` | merges cells; put `null` in the covered cells |
| `height` | `30` | row height in points |

Options: `sheet`, `columns: [{ width }]`, `autoWidth`, `stickyRowsCount`, `stickyColumnsCount`,
`showGridLines`, `rightToLeft`, `fontFamily`, `fontSize`, `dateFormat` (default `yyyy-mm-dd`), `title`, `author`.

Dates are written using their UTC date and time.

### Writing objects with a schema

```js
await writeXlsxFile(people, {
  filePath: 'people.xlsx',
  schema: [
    { column: 'Name',  type: String,  value: (p) => p.name, width: 24 },
    { column: 'Age',   type: Number,  value: (p) => p.age, align: 'right' },
    { column: 'Born',  type: Date,    value: (p) => p.born, format: 'dd.mm.yyyy' },
    { column: 'Active', type: Boolean, value: (p) => p.active,
      getCellStyle: (p) => (p.active ? null : { color: '#999999' }) }
  ]
})
```

The header row is bold and frozen by default (`headerStyle`, `stickyHeader: false` to change).
A file written with a schema can be read back with the matching read schema, so the round trip is lossless.

### Multiple sheets

```js
import { writeXlsxFileSheets } from 'read-write-xlsx'

await writeXlsxFileSheets([
  { name: 'Summary', data: summaryRows, stickyRowsCount: 1 },
  { name: 'People', data: people, schema: peopleSchema }
], { filePath: 'report.xlsx', fontFamily: 'Arial' })
```

In the browser, use `writeXlsxSheets(sheets, options)` → `Uint8Array`. Invalid characters in sheet names
(`[]:*?/\`) are replaced, names are cut to 31 characters, and duplicate names get a suffix.

## Entry points

Import everything from `read-write-xlsx`. The right build is picked automatically:

| Environment | What you get |
| --- | --- |
| Node.js | file paths, URLs, Buffers and streams; files are saved with `filePath`; uses the faster `node:zlib` |
| Browsers, workers, edge, Deno, Bun | `File`, `Blob`, `ArrayBuffer` or `Uint8Array` input; `CompressionStream` for deflate |

`readXlsxFile`, `readSheetNamesFromFile`, `readXlsxFileSheets`, `writeXlsxFile` and `writeXlsxFileSheets`
exist in both. Outside Node, passing a file path or the `filePath` option throws an error that says so, and
`writeXlsxFile` returns a `Uint8Array` instead of a `Buffer`.

To pin one build, import `read-write-xlsx/node` or `read-write-xlsx/browser` explicitly.

The package is ESM-only. From CommonJS, use `await import('read-write-xlsx')`, or `require()` on
Node ≥ 20.19 / 22.12, which can load ES modules directly.

## Limits (by design, to stay small)

- `.xlsx` / `.xlsm` only. Legacy `.xls`, `.ods`, password-protected files and ZIP64 (>4 GB) aren't supported.
- The reader returns values only. It ignores styles, comments, images and charts.
- Files are processed in memory. Reading 100 000 rows × 5 columns takes about 0.7 s, and writing them takes about
  1.4 s in Node 22.
- The writer doesn't produce charts, images, data validation or conditional formatting.

## Development

```bash
npm test            # node:test, no dev dependencies
npm run fixtures    # regenerate test files (needs python3 + openpyxl, optionally LibreOffice)
```

To run the demo locally, serve the repo root (for example `python3 -m http.server`) and open `/demo/`.
It imports `../src/index.js` directly, with no build step. Pushes to `main` publish it to GitHub Pages
through `.github/workflows/pages.yml`.

## License

MIT
