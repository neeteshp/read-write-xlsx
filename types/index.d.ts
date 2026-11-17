export type CellValue = string | number | boolean | Date | null

export type Row = CellValue[]

/** Anything the universal entry can read. */
export type XlsxInput = Uint8Array | ArrayBuffer | ArrayBufferView | Blob

// ---------------- Reading ----------------

export interface ReadOptions {
  /** 1-based sheet number or sheet name. Defaults to the first visible sheet. */
  sheet?: number | string
  /** Convert date-formatted numbers to Date (UTC). Default: true. */
  dates?: boolean
  /** Drop trailing empty rows and columns. Default: true. */
  trim?: boolean
}

export type SchemaType =
  | StringConstructor
  | NumberConstructor
  | BooleanConstructor
  | DateConstructor
  | 'Integer'
  | ((value: CellValue) => unknown)

export interface SchemaEntry {
  /** Property name on the output object. Supports dot paths ("address.city"). Defaults to the column title. */
  prop?: string
  type?: SchemaType
  required?: boolean
  oneOf?: unknown[]
  /** Throw an Error to reject the value; its message becomes the error text. */
  validate?: (value: any) => void
}

/** Keyed by the column title in the header row. A string value is shorthand for `{ prop }`. */
export type ReadSchema = Record<string, SchemaEntry | string>

export interface SchemaError {
  /** 1-based Excel row number. */
  row: number
  column: string
  value: unknown
  error: string
}

export interface SchemaResult<T = Record<string, any>> {
  rows: T[]
  errors: SchemaError[]
}

export function readXlsx(input: XlsxInput, options?: ReadOptions): Promise<Row[]>
export function readXlsx<T = Record<string, any>>(
  input: XlsxInput,
  options: ReadOptions & { schema: ReadSchema }
): Promise<SchemaResult<T>>

export function readSheetNames(input: XlsxInput): Promise<string[]>

export function readXlsxSheets(
  input: XlsxInput,
  options?: ReadOptions
): Promise<{ name: string; data: Row[] }[]>
export function readXlsxSheets<T = Record<string, any>>(
  input: XlsxInput,
  options: ReadOptions & { schema: ReadSchema }
): Promise<{ name: string; data: SchemaResult<T> }[]>

/** Schema type for whole numbers. */
export function Integer(value: CellValue): number

// ---------------- Writing ----------------

export type BorderStyle =
  | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair'
  | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'slantDashDot'

export interface CellStyle {
  /** Excel number format, e.g. "#,##0.00", "0%", "dd.mm.yyyy". */
  format?: string
  fontFamily?: string
  fontSize?: number
  fontWeight?: 'bold'
  fontStyle?: 'italic'
  textDecoration?: 'underline' | 'line-through'
  /** "#RRGGBB" */
  color?: string
  /** "#RRGGBB" */
  backgroundColor?: string
  align?: 'left' | 'center' | 'right' | 'fill' | 'justify'
  alignVertical?: 'top' | 'center' | 'bottom'
  wrap?: boolean
  indent?: number
  textRotation?: number
  borderStyle?: BorderStyle
  borderColor?: string
  borderTopStyle?: BorderStyle
  borderTopColor?: string
  borderBottomStyle?: BorderStyle
  borderBottomColor?: string
  borderLeftStyle?: BorderStyle
  borderLeftColor?: string
  borderRightStyle?: BorderStyle
  borderRightColor?: string
  /** Row height in points (applies to the whole row). */
  height?: number
}

export type WriteType =
  | StringConstructor
  | NumberConstructor
  | BooleanConstructor
  | DateConstructor
  | 'Integer'
  | 'Formula'

export interface CellObject extends CellStyle {
  value?: CellValue
  /** Inferred from the value when omitted. 'Formula' takes e.g. "SUM(A1:A3)". */
  type?: WriteType
  /** Merge this many columns to the right (put null in the covered cells). */
  columnSpan?: number
  /** Merge this many rows down. */
  rowSpan?: number
}

export type WriteCell = CellValue | undefined | CellObject

export interface WriteSchemaColumn<T = any> extends CellStyle {
  /** Header title. */
  column: string
  type?: WriteType
  /** Picks the value from an object. Defaults to `object[column]`. */
  value?: (object: T) => CellValue | undefined
  /** Column width in characters. */
  width?: number
  headerStyle?: CellStyle
  getCellStyle?: (object: T) => CellStyle | undefined
}

export interface SheetOptions {
  /** Column widths in characters, by column index. */
  columns?: ({ width?: number } | undefined)[]
  /** Estimate column widths from content (explicit `columns` win). */
  autoWidth?: boolean
  /** Freeze this many top rows (defaults to 1 when a schema is used). */
  stickyRowsCount?: number
  /** Freeze this many left columns. */
  stickyColumnsCount?: number
  showGridLines?: boolean
  rightToLeft?: boolean
  /** Style applied to schema header cells. Default: { fontWeight: 'bold' }. */
  headerStyle?: CellStyle
  /** Set to false to stop freezing the schema header row. */
  stickyHeader?: boolean
}

export interface WorkbookOptions extends SheetOptions {
  /** Default font. Default: "Calibri". */
  fontFamily?: string
  /** Default font size. Default: 11. */
  fontSize?: number
  /** Format for Date cells without their own `format`. Default: "yyyy-mm-dd". */
  dateFormat?: string
  title?: string
  author?: string
}

export interface WriteOptions extends WorkbookOptions {
  /** Sheet name. Default: "Sheet1". */
  sheet?: string
  schema?: WriteSchemaColumn[]
}

export interface SheetDefinition<T = any> extends SheetOptions {
  name?: string
  /** Rows of cells, or objects when `schema` is set. */
  data: WriteCell[][] | T[]
  schema?: WriteSchemaColumn<T>[]
}

export function writeXlsx(data: WriteCell[][], options?: WriteOptions): Promise<Uint8Array>
export function writeXlsx<T>(
  data: T[],
  options: WriteOptions & { schema: WriteSchemaColumn<T>[] }
): Promise<Uint8Array>

export function writeXlsxSheets(
  sheets: SheetDefinition[],
  options?: WorkbookOptions
): Promise<Uint8Array>

// ---------------- *File helpers (same names in Node and the browser) ----------------

/**
 * File paths, URLs and streams work in Node.js only. In the browser, pass a File, Blob,
 * ArrayBuffer or Uint8Array.
 */
export type FileInput = XlsxInput | string | URL | AsyncIterable<Uint8Array>

/** `filePath` works in Node.js only (it throws in the browser). */
export interface FilePathOption {
  filePath?: string
}

export function readXlsxFile(input: FileInput, options?: ReadOptions): Promise<Row[]>
export function readXlsxFile<T = Record<string, any>>(
  input: FileInput,
  options: ReadOptions & { schema: ReadSchema }
): Promise<SchemaResult<T>>

export function readSheetNamesFromFile(input: FileInput): Promise<string[]>

export function readXlsxFileSheets(
  input: FileInput,
  options?: ReadOptions
): Promise<{ name: string; data: Row[] }[]>
export function readXlsxFileSheets<T = Record<string, any>>(
  input: FileInput,
  options: ReadOptions & { schema: ReadSchema }
): Promise<{ name: string; data: SchemaResult<T> }[]>

/** Returns a Buffer in Node.js and a Uint8Array elsewhere. */
export function writeXlsxFile(
  data: WriteCell[][],
  options?: WriteOptions & FilePathOption
): Promise<Uint8Array>
export function writeXlsxFile<T>(
  data: T[],
  options: WriteOptions & { schema: WriteSchemaColumn<T>[] } & FilePathOption
): Promise<Uint8Array>

export function writeXlsxFileSheets(
  sheets: SheetDefinition[],
  options?: WorkbookOptions & FilePathOption
): Promise<Uint8Array>

// ---------------- Utilities ----------------

export function excelSerialToDate(serial: number, date1904?: boolean): Date
export function dateToExcelSerial(date: Date, date1904?: boolean): number

/** Swap the deflate implementation (the Node entry installs node:zlib automatically). */
export function setZipCodec(codec: {
  inflateRaw(bytes: Uint8Array): Uint8Array | Promise<Uint8Array>
  deflateRaw(bytes: Uint8Array): Uint8Array | Promise<Uint8Array>
}): void
