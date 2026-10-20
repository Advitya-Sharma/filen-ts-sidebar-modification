import { EXCEL_BUILTIN_FORMATS } from "../../sheet/numfmt/builtins"
import { XlsxReadError } from "../read/errors"
import type { XlsxBorder, XlsxCellStyle, XlsxDxf, XlsxFill, XlsxFont, XlsxStyleTable, XlsxXf } from "./types"

// The built-in number formats every file implies (ECMA-376 Part 1 §18.8.30), as Excel defines them (MS-OI29500
// §2.1.712 a: Excel shows id 14 as m/d/yyyy, so "mm-dd-yy" written as id 14 would show otherwise). Ids 5-8, 23-36,
// 41-44 and 50-81 depend on the language and are not listed.
export const BUILTIN_NUMBER_FORMATS: ReadonlyMap<number, string> = EXCEL_BUILTIN_FORMATS

const BUILTIN_IDS: ReadonlyMap<string, number> = new Map([...BUILTIN_NUMBER_FORMATS].map(([id, code]) => [code, id]))

// Custom number format ids start here (ECMA-376 Part 1 §18.8.30).
export const FIRST_CUSTOM_NUMBER_FORMAT = 164

// "Unique cell formats/cell styles" in Excel's specifications and limits.
export const MAX_CELL_XFS = 65_490

// A record's identity for reuse: keys sorted, so equal records written in any order compare alike.
export function recordKey(value: unknown): string {
	if (value === null || typeof value !== "object") return JSON.stringify(value) ?? "undefined"
	if (Array.isArray(value)) return `[${value.map(recordKey).join(",")}]`

	const entries = Object.entries(value as Record<string, unknown>)
		.filter(([, entry]) => entry !== undefined)
		.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))

	return `{${entries.map(([key, entry]) => `${JSON.stringify(key)}:${recordKey(entry)}`).join(",")}}`
}

export interface StyleSource {
	readonly numFmts: Map<number, string>
	readonly fonts: XlsxFont[]
	readonly fills: XlsxFill[]
	readonly borders: XlsxBorder[]
	readonly cellStyleXfs: XlsxXf[]
	readonly cellXfs: XlsxXf[]
	readonly cellStyles: XlsxCellStyle[]
	readonly dxfs: XlsxDxf[]
	readonly indexedColors: number[] | null
}

export class StyleTable implements XlsxStyleTable {
	readonly numFmts: Map<number, string>
	readonly fonts: XlsxFont[]
	readonly fills: XlsxFill[]
	readonly borders: XlsxBorder[]
	readonly cellStyleXfs: readonly XlsxXf[]
	readonly cellXfs: XlsxXf[]
	readonly cellStyles: readonly XlsxCellStyle[]
	readonly dxfs: XlsxDxf[]
	readonly indexedColors: readonly number[] | null
	readonly sourceCounts: {
		readonly numFmts: number
		readonly fonts: number
		readonly fills: number
		readonly borders: number
		readonly cellXfs: number
		readonly dxfs: number
	}
	// Appended records by key, so an equal append reuses one.
	private readonly appended = new Map<string, number>()
	// Custom format ids by code (first wins), built on first use.
	private codes: Map<string, number> | null = null

	constructor(source: StyleSource) {
		this.numFmts = source.numFmts
		this.fonts = source.fonts
		this.fills = source.fills
		this.borders = source.borders
		this.cellStyleXfs = source.cellStyleXfs
		this.cellXfs = source.cellXfs
		this.cellStyles = source.cellStyles
		this.dxfs = source.dxfs
		this.indexedColors = source.indexedColors
		this.sourceCounts = {
			numFmts: source.numFmts.size,
			fonts: source.fonts.length,
			fills: source.fills.length,
			borders: source.borders.length,
			cellXfs: source.cellXfs.length,
			dxfs: source.dxfs.length
		}
	}

	// A record's number format code: custom first (a file may redefine a built-in id), then the built-in one.
	numberFormat(id: number): string | undefined {
		return this.numFmts.get(id) ?? BUILTIN_NUMBER_FORMATS.get(id)
	}

	numFmtId(code: string): number {
		if (this.codes === null) {
			this.codes = new Map()

			for (const [id, custom] of this.numFmts) if (!this.codes.has(custom)) this.codes.set(custom, id)
		}

		const custom = this.codes.get(code)

		if (custom !== undefined) return custom

		const builtin = BUILTIN_IDS.get(code)

		if (builtin !== undefined && !this.numFmts.has(builtin)) return builtin

		let id = FIRST_CUSTOM_NUMBER_FORMAT

		for (const existing of this.numFmts.keys()) if (existing >= id) id = existing + 1

		this.numFmts.set(id, code)
		this.codes.set(code, id)

		return id
	}

	appendFont(font: XlsxFont): number {
		return this.append("font", this.fonts, font)
	}

	appendFill(fill: XlsxFill): number {
		return this.append("fill", this.fills, fill)
	}

	appendBorder(border: XlsxBorder): number {
		return this.append("border", this.borders, border)
	}

	appendXf(xf: XlsxXf): number {
		const key = `xf${recordKey(xf)}`
		const existing = this.appended.get(key)

		if (existing !== undefined) return existing
		if (this.cellXfs.length >= MAX_CELL_XFS)
			throw new XlsxReadError("styleLimit", "The workbook has as many cell formats as Excel allows")

		return this.append("xf", this.cellXfs, xf)
	}

	appendDxf(dxf: XlsxDxf): number {
		return this.append("dxf", this.dxfs, dxf)
	}

	private append<T>(kind: string, list: T[], record: T): number {
		const key = kind + recordKey(record)
		const existing = this.appended.get(key)

		if (existing !== undefined) return existing

		list.push(record)
		this.appended.set(key, list.length - 1)

		return list.length - 1
	}
}
