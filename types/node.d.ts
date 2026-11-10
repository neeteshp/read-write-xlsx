/// <reference types="node" />
import type { Readable } from 'node:stream'
import type {
  ReadOptions,
  ReadSchema,
  Row,
  SchemaResult,
  WriteCell,
  WriteOptions,
  WriteSchemaColumn,
  WorkbookOptions,
  SheetDefinition,
  XlsxInput
} from './index.js'

export * from './index.js'

export type NodeXlsxInput = XlsxInput | string | URL | Buffer | Readable

export function readXlsxFile(input: NodeXlsxInput, options?: ReadOptions): Promise<Row[]>
export function readXlsxFile<T = Record<string, any>>(
  input: NodeXlsxInput,
  options: ReadOptions & { schema: ReadSchema }
): Promise<SchemaResult<T>>

export function readSheetNamesFromFile(input: NodeXlsxInput): Promise<string[]>

export function readXlsxFileSheets(
  input: NodeXlsxInput,
  options?: ReadOptions
): Promise<{ name: string; data: Row[] }[]>

export function writeXlsxFile(
  data: WriteCell[][],
  options?: WriteOptions & { filePath?: string }
): Promise<Buffer>
export function writeXlsxFile<T>(
  data: T[],
  options: WriteOptions & { schema: WriteSchemaColumn<T>[]; filePath?: string }
): Promise<Buffer>

export function writeXlsxFileSheets(
  sheets: SheetDefinition[],
  options?: WorkbookOptions & { filePath?: string }
): Promise<Buffer>
