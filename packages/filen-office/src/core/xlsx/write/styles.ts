import type { SliceClock } from "../../slicing"
import { decodeCharacterData, decodeUtf8, encodeUtf8 } from "../../xml/index"
import type { XlsxBorder, XlsxFill, XlsxFont, XlsxStyleTable, XlsxXf } from "../model/types"
import { Scanner } from "../read/scan"
import { ByteBuilder, attr, eachChild, findAttribute, qualify, rootTag, scanTag, spliceTag, type ScannedTag } from "./markup"
import { STYLESHEET } from "./order.generated"
import { borderMarkup, dxfMarkup, fillMarkup, fontMarkup, numFmtMarkup, sameBorder, sameFill, sameFont, sameXf, xfMarkup } from "./records"

// styles.xml with the records edits appended: each container gets its new records at its end and its count
// updated; nothing already there moves, so every index in the workbook keeps its meaning. The part can be large
// (megabytes of cell formats): it is walked in slices, never parsed.

export const STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"

interface Container {
	readonly name: string
	readonly total: number
	readonly appended: readonly string[]
}

// What a save writes of the style table: the source's records and the appended ones up to the last that a
// written cell, row or column refers to (an edit undone leaves records appended that nothing refers to), with
// the records those formats refer to. `numFmts`: the custom number formats the written appended formats use.
export interface StylesExtent {
	readonly fonts: number
	readonly fills: number
	readonly borders: number
	readonly cellXfs: number
	readonly dxfs: number
	readonly numFmts: ReadonlySet<number>
}

// `usedXfs`: one past the highest cell format a written cell, row or column refers to.
export function stylesExtent(styles: XlsxStyleTable, usedXfs: number): StylesExtent {
	const counts = styles.sourceCounts
	const cellXfs = Math.min(styles.cellXfs.length, Math.max(counts.cellXfs, usedXfs))
	let fonts = counts.fonts
	let fills = counts.fills
	let borders = counts.borders
	const numFmts = new Set<number>()

	for (let k = counts.cellXfs; k < cellXfs; k++) {
		const xf = styles.cellXfs[k]

		if (xf === undefined) continue

		fonts = Math.max(fonts, xf.fontId + 1)
		fills = Math.max(fills, xf.fillId + 1)
		borders = Math.max(borders, xf.borderId + 1)

		if (styles.numFmts.has(xf.numFmtId)) numFmts.add(xf.numFmtId)
	}

	// A part written for a workbook that had none starts with the records it implied (seedImpliedStyles).
	if (cellXfs > counts.cellXfs) {
		if (counts.fonts === 0) fonts = Math.max(fonts, 1)
		if (counts.fills === 0) fills = Math.max(fills, IMPLIED_FILLS.length)
		if (counts.borders === 0) borders = Math.max(borders, 1)
	}

	return {
		fonts: Math.min(styles.fonts.length, fonts),
		fills: Math.min(styles.fills.length, fills),
		borders: Math.min(styles.borders.length, borders),
		cellXfs,
		dxfs: styles.dxfs.length,
		numFmts
	}
}

// Whether the extent holds anything the source does not.
export function stylesGrow(styles: XlsxStyleTable, extent: StylesExtent): boolean {
	const counts = styles.sourceCounts

	return (
		extent.cellXfs > counts.cellXfs ||
		extent.fonts > counts.fonts ||
		extent.fills > counts.fills ||
		extent.borders > counts.borders ||
		extent.dxfs > counts.dxfs
	)
}

// `numFmts`: the custom number format ids the source part holds.
function containers(styles: XlsxStyleTable, extent: StylesExtent, prefix: string, numFmts: ReadonlySet<number>): Container[] {
	const counts = styles.sourceCounts
	const custom = [...styles.numFmts].filter(([id]) => !numFmts.has(id) && extent.numFmts.has(id)).sort((a, b) => a[0] - b[0])
	const appended = <T>(list: readonly T[], from: number, to: number, markup: (record: T) => string): string[] =>
		list.slice(from, to).map(markup)

	return [
		{ name: "numFmts", total: numFmts.size + custom.length, appended: custom.map(([id, code]) => numFmtMarkup(id, code, prefix)) },
		{
			name: "fonts",
			total: extent.fonts,
			appended: appended(styles.fonts, counts.fonts, extent.fonts, font => fontMarkup(font, prefix))
		},
		{
			name: "fills",
			total: extent.fills,
			appended: appended(styles.fills, counts.fills, extent.fills, fill => fillMarkup(fill, prefix))
		},
		{
			name: "borders",
			total: extent.borders,
			appended: appended(styles.borders, counts.borders, extent.borders, border => borderMarkup(border, prefix))
		},
		{
			name: "cellXfs",
			total: extent.cellXfs,
			appended: appended(styles.cellXfs, counts.cellXfs, extent.cellXfs, xf => xfMarkup(xf, prefix))
		},
		{ name: "dxfs", total: extent.dxfs, appended: appended(styles.dxfs, counts.dxfs, extent.dxfs, dxf => dxfMarkup(dxf, prefix)) }
	]
}

// A whole styles part from the model (a workbook that had none).
function newStyles(styles: XlsxStyleTable, extent: StylesExtent, main: string): Uint8Array {
	let markup = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<styleSheet xmlns="${main}">`
	const all = containers(styles, extent, "", new Set())
	const write = (name: string, records: readonly string[]): void => {
		if (records.length > 0) markup += `<${name} count="${records.length}">${records.join("")}</${name}>`
	}

	for (const name of STYLESHEET) {
		const container = all.find(candidate => candidate.name === name)

		if (container !== undefined) {
			write(name, container.appended)
		} else if (name === "cellStyleXfs") {
			write(
				name,
				(styles.cellStyleXfs.length > 0
					? styles.cellStyleXfs
					: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: -1, flags: 0 }]
				).map(xf => xfMarkup({ ...xf, xfId: -1 }, ""))
			)
		} else if (name === "cellStyles") {
			const cellStyles = styles.cellStyles.length > 0 ? styles.cellStyles : [{ name: "Normal", xfId: 0, builtinId: 0 }]

			write(
				name,
				cellStyles.map(
					style =>
						`<cellStyle${attr("name", style.name)}${attr("xfId", style.xfId)}${style.builtinId === undefined ? "" : attr("builtinId", style.builtinId)}/>`
				)
			)
		}
	}

	return encodeUtf8(`${markup}</styleSheet>`)
}

// A top-level child of the source's root: its local name and span.
interface Child {
	readonly name: string
	readonly start: number
	readonly end: number
}

function localName(bytes: Uint8Array, tag: ScannedTag): string {
	const name = decodeUtf8(bytes, tag.start + 1, tag.nameEnd)

	return name.slice(name.indexOf(":") + 1)
}

// `main`: the SpreadsheetML namespace of a new part. The source is never parsed: its root's children are walked
// (in slices of `clock`), the numFmts ids read from their tags, and the appended records and counts spliced in.
export async function writeStyles(
	styles: XlsxStyleTable,
	extent: StylesExtent,
	source: Uint8Array | null,
	main: string,
	part: string,
	clock?: SliceClock
): Promise<Uint8Array> {
	if (source === null) return newStyles(styles, extent, main)

	const { tag: root, prefix, name: rootName } = rootTag(source)
	const scanner = new Scanner(source, part)
	const children: Child[] = []
	const rootClose = root.selfClosing
		? root.end
		: await eachChild(scanner, root.start, root.end, clock, (start, end) => {
				const tag = scanTag(source, start)

				if (tag !== undefined) children.push({ name: localName(source, tag), start, end })
			})
	const numFmts = new Set<number>()
	const numFmtsChild = children.find(child => child.name === "numFmts")
	const numFmtsTag = numFmtsChild === undefined ? undefined : scanTag(source, numFmtsChild.start)

	if (numFmtsChild !== undefined && numFmtsTag !== undefined && !numFmtsTag.selfClosing) {
		await eachChild(scanner, numFmtsChild.start, numFmtsTag.end, clock, start => {
			const tag = scanTag(source, start)
			const at = tag === undefined ? -1 : findAttribute(source, tag, "numFmtId")

			if (tag !== undefined && at >= 0) {
				numFmts.add(Number(decodeCharacterData(source, tag.attributes[at + 2] ?? 0, tag.attributes[at + 3] ?? 0, true)))
			}
		})
	}

	// [from, to, replacement], ascending by `from` (stable: new containers at one place stay in schema order).
	const edits: [number, number, Uint8Array | string][] = []

	for (const container of containers(styles, extent, prefix, numFmts)) {
		if (container.appended.length === 0) continue

		const appended = container.appended.join("")
		const child = children.find(candidate => candidate.name === container.name)
		const tag = child === undefined ? undefined : scanTag(source, child.start)

		if (child !== undefined && tag !== undefined) {
			const head = spliceTag(source, tag, new Map([["count", String(container.total)]]), true)

			if (tag.selfClosing) {
				edits.push(
					[child.start, child.end, head],
					[child.end, child.end, `${appended}</${decodeUtf8(source, tag.start + 1, tag.nameEnd)}>`]
				)
			} else {
				const close = source.lastIndexOf(60, child.end - 1)

				edits.push([tag.start, tag.end, head], [close, close, appended])
			}

			continue
		}

		const name = qualify(prefix, container.name)
		const position = STYLESHEET.indexOf(container.name)
		const before = children.find(candidate => STYLESHEET.indexOf(candidate.name) > position)

		edits.push([before?.start ?? rootClose, before?.start ?? rootClose, `<${name} count="${container.total}">${appended}</${name}>`])
	}

	if (edits.length === 0) return source

	if (root.selfClosing) {
		edits.unshift([root.start, root.end, spliceTag(source, root, new Map(), true)])
		edits.push([root.end, root.end, `</${rootName}>`])
	}

	edits.sort((a, z) => a[0] - z[0])

	const out = new ByteBuilder()
	let at = 0

	for (const [from, to, replacement] of edits) {
		out.bytes(source, at, from)

		if (typeof replacement === "string") out.text(replacement)
		else out.bytes(replacement)

		at = to
	}

	out.bytes(source, at, source.length)

	return out.take()
}

// Whether cell format 0 is what an unformatted cell shows: the source's, or the defaults seedImpliedStyles put
// first in a workbook that had none (then a cell with format 0 needs no styles part written).
export function formatZeroImplied(styles: XlsxStyleTable): boolean {
	const counts = styles.sourceCounts

	return (
		counts.cellXfs > 0 ||
		styles.cellXfs.length === 0 ||
		(sameXf(styles.cellXfs[0], IMPLIED_XF) &&
			(counts.fonts > 0 || sameFont(styles.fonts[0], IMPLIED_FONT)) &&
			(counts.fills > 0 || IMPLIED_FILLS.every((fill, k) => sameFill(styles.fills[k], fill))) &&
			(counts.borders > 0 || sameBorder(styles.borders[0], IMPLIED_BORDER)))
	)
}

// The records a stylesheet implies when it lacks them (Excel's defaults: one font, the reserved "none" and
// "gray125" fills, an empty border, one cell format).
export const IMPLIED_FONT: XlsxFont = { name: "Calibri", size: 11, family: 2, scheme: "minor", color: { theme: 1 } }
export const IMPLIED_FILLS: readonly XlsxFill[] = [
	{ kind: "pattern", pattern: "none" },
	{ kind: "pattern", pattern: "gray125" }
]
export const IMPLIED_BORDER: XlsxBorder = {}
export const IMPLIED_XF: XlsxXf = { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, flags: 0 }

// A workbook without styles.xml has an empty table, where an edit's first record would take index 0, the
// format of every unformatted cell. Seeded at open, the defaults are appended records: a save writes them
// (first) only when an edit appends styles of its own.
export function seedImpliedStyles(styles: XlsxStyleTable): void {
	if (styles.fonts.length === 0) styles.appendFont(IMPLIED_FONT)
	if (styles.fills.length === 0) for (const fill of IMPLIED_FILLS) styles.appendFill(fill)
	if (styles.borders.length === 0) styles.appendBorder(IMPLIED_BORDER)
	if (styles.cellXfs.length === 0) styles.appendXf(IMPLIED_XF)
}
