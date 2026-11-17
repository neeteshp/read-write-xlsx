// Live demo for read-write-xlsx. Loads the library's own source directly; no build step.
import {
  readXlsx,
  readXlsxSheets,
  readSheetNames,
  writeXlsx,
  writeXlsxSheets,
  Integer
} from '../src/index.js'

const $ = (id) => document.getElementById(id)
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// ---------------------------------------------------------------- support check
;(() => {
  const note = $('support')
  try {
    new CompressionStream('deflate-raw')
    new DecompressionStream('deflate-raw')
    note.textContent = 'Your browser supports CompressionStream("deflate-raw"), so everything below runs on this page.'
  } catch {
    note.className = 'note bad'
    note.textContent =
      'This browser does not support CompressionStream("deflate-raw"), which the library needs. ' +
      'Try a current version of Chrome, Edge, Firefox or Safari.'
  }
})()

// ---------------------------------------------------------------- helpers
const columnLetter = (i) => {
  let s = ''
  for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s
  return s
}

function formatDate(d) {
  const iso = d.toISOString()
  return iso.endsWith('T00:00:00.000Z') ? iso.slice(0, 10) : iso.slice(0, 19).replace('T', ' ')
}

function formatBytes(n) {
  return n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1048576).toFixed(2)} MB`
}

const ms = (t) => `${(performance.now() - t).toFixed(0)} ms`

/** Renders rows as a spreadsheet-like grid. `bad` is a Set of "r,c" keys to highlight. */
function renderGrid(container, rows, { maxRows = 500, bad } = {}) {
  container.textContent = ''
  if (!rows.length) {
    container.innerHTML = '<div class="sheet"><div class="empty">This sheet is empty.</div></div>'
    return
  }
  const cols = Math.max(1, ...rows.map((r) => r.length))
  const table = document.createElement('table')
  table.className = 'sheet'
  const head = table.createTHead().insertRow()
  head.appendChild(document.createElement('th'))
  for (let c = 0; c < cols; c++) {
    const th = document.createElement('th')
    th.textContent = columnLetter(c)
    head.appendChild(th)
  }
  const body = table.createTBody()
  const shown = Math.min(rows.length, maxRows)
  for (let r = 0; r < shown; r++) {
    const tr = body.insertRow()
    const th = document.createElement('th')
    th.textContent = r + 1
    tr.appendChild(th)
    for (let c = 0; c < cols; c++) {
      const td = tr.insertCell()
      const v = rows[r][c]
      if (v == null) continue
      if (v instanceof Date) { td.textContent = formatDate(v); td.className = 'd' }
      else if (typeof v === 'number') { td.textContent = String(v); td.className = 'n' }
      else if (typeof v === 'boolean') { td.textContent = v ? 'TRUE' : 'FALSE'; td.className = 'b' }
      else td.textContent = v
      td.title = td.textContent
      if (bad?.has(`${r},${c}`)) td.classList.add('bad')
    }
  }
  if (rows.length > shown) {
    const tr = body.insertRow()
    const td = tr.insertCell()
    td.colSpan = cols + 1
    td.className = 'empty'
    td.textContent = `… ${rows.length - shown} more rows not shown`
  }
  container.appendChild(table)
}

function renderTabs(container, names, onSelect) {
  container.textContent = ''
  names.forEach((name, i) => {
    const b = document.createElement('button')
    b.type = 'button'
    b.role = 'tab'
    b.textContent = name
    b.setAttribute('aria-selected', String(i === 0))
    b.onclick = () => {
      for (const x of container.children) x.setAttribute('aria-selected', String(x === b))
      onSelect(i)
    }
    container.appendChild(b)
  })
  onSelect(0)
}

function showError(el, err) {
  el.hidden = !err
  if (err) el.textContent = err instanceof Error ? `${err.name}: ${err.message}` : String(err)
}

// ---------------------------------------------------------------- 1. READ
let readView = 'grid'
let lastRead = null // { input, name }

async function readFile(input, name) {
  lastRead = { input, name }
  showError($('read-error'), null)
  try {
    const opts = { dates: $('opt-dates').checked, trim: $('opt-trim').checked }
    const bytes = new Uint8Array(await (input.arrayBuffer ? input.arrayBuffer() : input))
    const t = performance.now()
    const sheets = await readXlsxSheets(bytes, opts)
    const took = ms(t)
    const cells = sheets.reduce((n, s) => n + s.data.reduce((m, r) => m + r.filter((v) => v != null).length, 0), 0)
    $('read-meta').textContent =
      `${name} · ${formatBytes(bytes.length)} · ${sheets.length} sheet${sheets.length === 1 ? '' : 's'} · ` +
      `${cells.toLocaleString()} non-empty cells · parsed in ${took}`
    $('read-out').hidden = false
    renderTabs($('read-tabs'), sheets.map((s) => s.name), (i) => {
      renderGrid($('read-grid'), sheets[i].data)
      $('read-json').textContent = JSON.stringify(sheets[i].data.slice(0, 200), null, 1)
        .replace(/\n {2,}/g, ' ').replace(/\n \]/g, ' ]')
    })
  } catch (err) {
    $('read-out').hidden = true
    showError($('read-error'), err)
  }
}

const drop = $('drop')
$('file').addEventListener('change', (e) => {
  const f = e.target.files[0]
  if (f) readFile(f, f.name)
})
;['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over') }))
;['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, () => drop.classList.remove('over')))
drop.addEventListener('drop', (e) => {
  e.preventDefault()
  const f = e.dataTransfer.files[0]
  if (f) readFile(f, f.name)
})
for (const id of ['opt-dates', 'opt-trim']) {
  $(id).addEventListener('change', () => lastRead && readFile(lastRead.input, lastRead.name))
}
for (const b of document.querySelectorAll('.view-toggle .seg')) {
  b.addEventListener('click', () => {
    readView = b.dataset.view
    for (const x of document.querySelectorAll('.view-toggle .seg')) x.classList.toggle('on', x === b)
    $('read-grid').hidden = readView !== 'grid'
    $('read-json').hidden = readView !== 'json'
  })
}
$('sample').addEventListener('click', async () => {
  const bytes = await writeXlsxSheets([
    {
      name: 'Orders',
      stickyRowsCount: 1,
      data: [
        ['Order', 'Customer', 'Placed', 'Qty', 'Unit price', 'Shipped'],
        ...Array.from({ length: 40 }, (_, i) => [
          1001 + i,
          ['Acme', 'Globex', 'Initech', 'Umbrella', 'Hooli'][i % 5],
          new Date(Date.UTC(2026, 0, 3 + i * 4)),
          (i * 7) % 13 + 1,
          Math.round((9.99 + i * 3.37) * 100) / 100,
          i % 3 !== 0
        ])
      ]
    },
    { name: 'Notes', data: [['Generated in your browser by writeXlsxSheets()'], [], ['Sparse cell →', null, null, 'here']] }
  ])
  readFile(bytes, 'sample.xlsx')
})

// ---------------------------------------------------------------- 2. WRITE
const EXAMPLES = [
  {
    id: 'styled',
    label: 'Styled invoice',
    file: 'invoice.xlsx',
    code: `// Cells can be plain values or objects with styles.
const header = (value) => ({
  value, fontWeight: 'bold', color: '#FFFFFF', backgroundColor: '#1D7A46',
  align: 'center', borderStyle: 'thin', borderColor: '#145C34'
})

const items = [
  ['Design work',       12, 85],
  ['Development',       40, 95],
  ['Hosting (annual)',   1, 240],
  ['Support retainer',   3, 150]
]

return writeXlsx([
  [{ value: 'INVOICE #2026-114', fontSize: 16, fontWeight: 'bold', columnSpan: 4 }, null, null, null],
  ['Issued', new Date('2026-11-17'), 'Due', { value: new Date('2026-12-17'), color: '#C00000' }],
  [],
  [header('Item'), header('Hours'), header('Rate'), header('Amount')],
  ...items.map(([item, hours, rate], i) => [
    item,
    { value: hours, type: 'Integer' },
    { value: rate, format: '$#,##0.00' },
    { type: 'Formula', value: \`B\${i + 5}*C\${i + 5}\`, format: '$#,##0.00' }
  ]),
  [
    { value: 'Total', fontWeight: 'bold', columnSpan: 3, align: 'right' }, null, null,
    { type: 'Formula', value: \`SUM(D5:D\${items.length + 4})\`, format: '$#,##0.00',
      fontWeight: 'bold', backgroundColor: '#FFF2CC', borderTopStyle: 'double' }
  ]
], {
  sheet: 'Invoice',
  columns: [{ width: 24 }, { width: 10 }, { width: 12 }, { width: 14 }],
  dateFormat: 'dd mmm yyyy'
})`
  },
  {
    id: 'schema',
    label: 'Objects + schema',
    file: 'people.xlsx',
    code: `// Write an array of objects. The schema defines columns, types and styles.
const people = [
  { name: 'Ada Lovelace',    age: 36, born: new Date('1815-12-10'), active: true,  score: 0.98 },
  { name: 'Alan Turing',     age: 41, born: new Date('1912-06-23'), active: true,  score: 0.95 },
  { name: 'Grace Hopper',    age: 85, born: new Date('1906-12-09'), active: false, score: 0.91 },
  { name: 'Edsger Dijkstra', age: 72, born: new Date('1930-05-11'), active: false, score: 0.89 }
]

return writeXlsx(people, {
  sheet: 'People',
  schema: [
    { column: 'Name',   type: String,  value: (p) => p.name, width: 22 },
    { column: 'Age',    type: Number,  value: (p) => p.age, align: 'right' },
    { column: 'Born',   type: Date,    value: (p) => p.born, format: 'dd.mm.yyyy', width: 14 },
    { column: 'Score',  type: Number,  value: (p) => p.score, format: '0%' },
    { column: 'Active', type: Boolean, value: (p) => p.active,
      getCellStyle: (p) => (p.active ? null : { color: '#999999' }) }
  ]
})`
  },
  {
    id: 'sheets',
    label: 'Multiple sheets',
    file: 'report.xlsx',
    code: `// Several sheets in one workbook, each with its own options.
const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
const regions = ['North', 'South', 'East', 'West']
const sales = regions.map((r, i) => [r, ...months.map((_, m) => 1000 + ((i + 1) * 137 * (m + 3)) % 900)])

return writeXlsxSheets([
  {
    name: 'Summary',
    data: [
      [{ value: 'H1 2026 sales', fontWeight: 'bold', fontSize: 14 }],
      ['Regions', regions.length],
      ['Grand total', { type: 'Formula', value: "SUM('Sales by region'!B2:G5)", format: '#,##0' }]
    ],
    columns: [{ width: 16 }, { width: 14 }]
  },
  {
    name: 'Sales by region',
    stickyRowsCount: 1,
    stickyColumnsCount: 1,
    data: [
      ['Region', ...months].map((v) => ({ value: v, fontWeight: 'bold', backgroundColor: '#E5F3EA' })),
      ...sales
    ]
  },
  { name: 'Notes', showGridLines: false, data: [['Sheet names with [ ] : * ? / \\\\ are cleaned up automatically.']] }
], { fontFamily: 'Arial' })`
  },
  {
    id: 'big',
    label: '50,000 rows',
    file: 'large.xlsx',
    code: `// Speed test: how long does your browser take to build 50k rows × 6 columns?
const start = new Date('2026-01-01').getTime()
const rows = [['ID', 'SKU', 'Date', 'Qty', 'Price', 'In stock']]
for (let i = 1; i <= 50000; i++) {
  rows.push([
    i,
    'SKU-' + (i * 7919 % 100000).toString().padStart(5, '0'),
    new Date(start + (i % 365) * 86400000),
    i % 97,
    Math.round(i * 1.37 % 500 * 100) / 100,
    i % 4 !== 0
  ])
}
return writeXlsx(rows, { sheet: 'Inventory', stickyRowsCount: 1, autoWidth: true })`
  }
]

let current = EXAMPLES[0]
const edited = new Map() // example id -> edited code
let downloadUrl = null

function selectExample(ex) {
  if (current) edited.set(current.id, $('code').value)
  current = ex
  $('code').value = edited.get(ex.id) ?? ex.code
  $('code-label').textContent = ex.file.replace('.xlsx', '.js')
  for (const b of $('examples').children) b.setAttribute('aria-selected', String(b.dataset.id === ex.id))
}

for (const ex of EXAMPLES) {
  const b = document.createElement('button')
  b.type = 'button'
  b.role = 'tab'
  b.dataset.id = ex.id
  b.textContent = ex.label
  b.onclick = () => selectExample(ex)
  $('examples').appendChild(b)
}
selectExample(EXAMPLES[0])

const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor

async function run() {
  showError($('write-error'), null)
  $('write-meta').textContent = 'Running…'
  $('run').disabled = true
  try {
    const fn = new AsyncFunction('writeXlsx', 'writeXlsxSheets', 'readXlsx', 'readXlsxSheets', 'readSheetNames', 'Integer', $('code').value)
    const t = performance.now()
    const bytes = await fn(writeXlsx, writeXlsxSheets, readXlsx, readXlsxSheets, readSheetNames, Integer)
    if (!(bytes instanceof Uint8Array)) throw new TypeError('The code must return the Uint8Array from writeXlsx() / writeXlsxSheets().')
    const wrote = ms(t)

    if (downloadUrl) URL.revokeObjectURL(downloadUrl)
    downloadUrl = URL.createObjectURL(new Blob([bytes], { type: XLSX_MIME }))
    const a = $('download')
    a.href = downloadUrl
    a.download = current.file
    $('dl-name').textContent = `${current.file} (${formatBytes(bytes.length)})`
    a.hidden = false

    const t2 = performance.now()
    const sheets = await readXlsxSheets(bytes)
    $('write-meta').textContent = `Written in ${wrote}, read back in ${ms(t2)}`
    $('write-out').hidden = false
    renderTabs($('write-tabs'), sheets.map((s) => s.name), (i) => renderGrid($('write-grid'), sheets[i].data, { maxRows: 200 }))
  } catch (err) {
    $('write-meta').textContent = ''
    showError($('write-error'), err)
  } finally {
    $('run').disabled = false
  }
}

$('run').addEventListener('click', run)
$('reset').addEventListener('click', () => { edited.delete(current.id); $('code').value = current.code })
$('code').addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); run() }
  if (e.key === 'Tab') {
    e.preventDefault()
    const el = e.target
    el.setRangeText('  ', el.selectionStart, el.selectionEnd, 'end')
  }
})

// ---------------------------------------------------------------- 3. SCHEMA
const SCHEMA_SRC = `{
  'Name':   { prop: 'name', type: String, required: true },
  'Age':    { prop: 'age', type: Integer },
  'Joined': { prop: 'joined', type: Date },
  'Active': { prop: 'flags.active', type: Boolean },
  'Status': { prop: 'status', oneOf: ['open', 'closed'] },
  'E-mail': { prop: 'email', validate: (v) => {
    if (!String(v).includes('@')) throw new Error('invalid email')
  } }
}`

;(async () => {
  const sheet = [
    ['Name', 'Age', 'Joined', 'Active', 'Status', 'E-mail'],
    ['Ada', 36, new Date('2024-03-01'), true, 'open', 'ada@example.com'],
    ['Linus', 'thirty', new Date('2025-07-15'), false, 'closed', 'linus@example.com'],
    [null, 28, new Date('2023-11-30'), true, 'open', 'anon@example.com'],
    ['Grace', 41.5, new Date('2022-01-09'), true, 'pending', 'grace@example.com'],
    ['Alan', 41, new Date('2021-06-23'), true, 'closed', 'alan.example.com']
  ]
  const bytes = await writeXlsx(sheet)
  const schema = new Function('Integer', `return (${SCHEMA_SRC})`)(Integer)
  const { rows, errors } = await readXlsx(bytes, { schema })

  const header = sheet[0]
  const bad = new Set(errors.map((e) => `${e.row - 1},${header.indexOf(e.column)}`))
  renderGrid($('schema-sheet'), sheet, { bad })
  $('schema-code').textContent = `const { rows, errors } = await readXlsx(file, {\n  schema: ${SCHEMA_SRC.replace(/\n/g, '\n  ')}\n})`
  $('schema-rows').textContent = JSON.stringify(rows, null, 2)

  const t = document.createElement('table')
  t.className = 'errors'
  t.innerHTML = '<thead><tr><th>row</th><th>column</th><th>value</th><th>error</th></tr></thead>'
  const tb = t.createTBody()
  for (const e of errors) {
    const tr = tb.insertRow()
    for (const v of [e.row, e.column, e.value === null ? 'null' : JSON.stringify(e.value), e.error]) {
      tr.insertCell().textContent = v
    }
  }
  $('schema-errors').appendChild(t)
})().catch((err) => {
  $('schema-rows').textContent = String(err)
})
