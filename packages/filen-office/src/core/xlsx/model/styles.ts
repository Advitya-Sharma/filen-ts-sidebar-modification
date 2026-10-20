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

// A border's identity: an edge without a line (<left/>, style "none") is no edge.
function borderKey(border: XlsxBorder): string {
	const lines: Record<string, unknown> = {}

	for (const [key, value] of Object.entries(border)) {
		if (typeof value === "object" && (value.style === undefined || (value.style as string) === "none")) continue

		lines[key] = value
	}

	return recordKey(lines)
}

function keyOf(kind: string, record: unknown): string {
	return kind + (kind === "border" ? borderKey(record as XlsxBorder) : recordKey(record))
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
	// Records by key, so an equal append reuses one: appended ones, and the source's (first of equals), indexed on
	// a kind's first append so a change back to a format the file had finds it.
	private readonly appended = new Map<string, number>()
	private readonly indexed = new Set<string>()
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

	// `copy`: a record of its own even where the source holds an equal one (an appended equal one is still reused).
	appendXf(xf: XlsxXf, copy = false): number {
		const key = `xf${recordKey(xf)}`

		if (!copy) this.indexSource("xf", this.cellXfs, this.sourceCounts.cellXfs)

		const existing = copy ? this.copies.get(key) : this.appended.get(key)

		if (existing !== undefined) return existing
		if (this.cellXfs.length >= MAX_CELL_XFS)
			throw new XlsxReadError("styleLimit", "The workbook has as many cell formats as Excel allows")

		this.cellXfs.push(xf)

		const index = this.cellXfs.length - 1

		this.copies.set(key, index)
		if (!this.appended.has(key)) this.appended.set(key, index)

		return index
	}

	appendDxf(dxf: XlsxDxf): number {
		return this.append("dxf", this.dxfs, dxf)
	}

	// Appended cell formats by key, for `copy` appends.
	private readonly copies = new Map<string, number>()

	private indexSource<T>(kind: string, list: readonly T[], count: number): void {
		if (this.indexed.has(kind)) return

		this.indexed.add(kind)

		for (let i = 0; i < count && i < list.length; i++) {
			const key = keyOf(kind, list[i])

			if (!this.appended.has(key)) this.appended.set(key, i)
		}
	}

	private append<T>(kind: "font" | "fill" | "border" | "dxf", list: T[], record: T): number {
		this.indexSource(
			kind,
			list,
			this.sourceCounts[kind === "font" ? "fonts" : kind === "fill" ? "fills" : kind === "border" ? "borders" : "dxfs"]
		)

		const key = keyOf(kind, record)
		const existing = this.appended.get(key)

		if (existing !== undefined) return existing

		list.push(record)
		this.appended.set(key, list.length - 1)

		return list.length - 1
	}
}
