import type { OdfStyleResolver, ResolvedOdfStyle } from "../../odf/resolve"
import type { OdfProperties, OdfStyleSheet } from "../../odf/styles"
import { parseAngle, parseColor, parseLength, pointsToPixels } from "../../odf/units"
import type { SheetBorderStyle, SheetHorizontalAlign, SheetVerticalAlign } from "../../sheet/protocol"
import { XfApply } from "../../sheet/style/xfApply"
import { StyleTable } from "../../xlsx/model/styles"
import type { XlsxAlignment, XlsxBorder, XlsxBorderEdge, XlsxCellProtection, XlsxFill, XlsxFont, XlsxXf } from "../../xlsx/model/types"
import { dataStyleCode } from "./dataStyle"
import type { OdsIssues } from "../issues"

// ODF cell styles (family table-cell, ODF 1.3 Part 3 §16.2 with §17.18 table-cell-properties, §17.6 paragraph and
// §16.27 text properties) as cell formats of the workbook's style table: each style name the body uses becomes
// one cellXfs record (equal ones shared). Column, row and table styles give sizes and sheet properties.

// LibreOffice's default cell font when no style names one.
const DEFAULT_FONT = "Liberation Sans"
const DEFAULT_SIZE = 10

export interface OdsColumnStyle {
	// Pixels, NaN for none.
	readonly widthPx: number
	readonly pageBreak: boolean
}

export interface OdsRowStyle {
	// Points, NaN for none.
	readonly heightPt: number
	// style:use-optimal-row-height="false": the height is the user's (customHeight).
	readonly custom: boolean
	readonly pageBreak: boolean
}

export interface OdsTableStyle {
	readonly display: boolean
	readonly rightToLeft: boolean
	// 0xRRGGBB, -1 for none.
	readonly tabColor: number
}

function argb(color: string | undefined): number | undefined {
	const rgb = parseColor(color)

	return rgb === undefined ? undefined : (0xff000000 | rgb) >>> 0
}

// fo:border-* (§20.183): "<width> <style> <color>" in any order.
function borderEdge(value: string | undefined, lineWidth: string | undefined): XlsxBorderEdge | undefined {
	if (value === undefined) return undefined

	const words = value.trim().split(/\s+/)

	if (words.includes("none") || words.includes("hidden")) return undefined

	let width = Number.NaN
	let kind = "solid"
	let color: number | undefined

	for (const word of words) {
		const length = parseLength(word)
		const rgb = argb(word)

		if (length !== undefined) width = length
		else if (rgb !== undefined) color = rgb
		else if (/^(thin|medium|thick)$/.test(word)) width = word === "thin" ? 0.75 : word === "medium" ? 1.75 : 2.5
		else kind = word
	}

	// A double line's own widths (style:border-line-width: inner, distance, outer) give its whole width.
	if (lineWidth !== undefined) {
		const parts = lineWidth.trim().split(/\s+/).map(parseLength)

		if (parts.every(part => part !== undefined)) width = parts.reduce<number>((sum, part) => sum + (part ?? 0), 0)
	}

	if (width === 0) return undefined

	const heavy = width > 1.25
	let style: SheetBorderStyle

	switch (kind) {
		case "double":
		case "double-thin":
			style = "double"
			break
		case "dotted":
			style = heavy ? "mediumDashed" : "dotted"
			break
		case "dashed":
		case "dash":
		case "long-dash":
		case "fine-dashed":
			style = heavy ? "mediumDashed" : "dashed"
			break
		case "dot-dash":
		case "dash-dot":
			style = heavy ? "mediumDashDot" : "dashDot"
			break
		case "dot-dot-dash":
		case "dash-dot-dot":
			style = heavy ? "mediumDashDotDot" : "dashDotDot"
			break
		default:
			style = Number.isNaN(width) || width <= 1.25 ? "thin" : width <= 2.25 ? "medium" : "thick"
	}

	return color === undefined ? { style } : { style, color: { argb: color } }
}

function horizontal(paragraph: OdfProperties, cell: OdfProperties): SheetHorizontalAlign | "general" | undefined {
	if (cell.attributes.get("style:text-align-source") === "value-type") return undefined

	switch (paragraph.attributes.get("fo:text-align")) {
		case "start":
		case "left":
			return "left"
		case "end":
		case "right":
			return "right"
		case "center":
			return "center"
		case "justify":
			return "justify"
		default:
			return cell.attributes.get("style:repeat-content") === "true" ? "fill" : undefined
	}
}

function vertical(cell: OdfProperties): SheetVerticalAlign | undefined {
	switch (cell.attributes.get("style:vertical-align")) {
		case "top":
			return "top"
		case "middle":
			return "center"
		case "bottom":
			return "bottom"
		default:
			return undefined
	}
}

function protection(value: string | undefined): XlsxCellProtection {
	// style:cell-protect (§20.189): "none", "hidden-and-protected", or "protected" and "formula-hidden" in a list.
	const words = new Set((value ?? "protected").trim().split(/\s+/))

	if (words.has("hidden-and-protected")) return { locked: true, hidden: true, hideAll: true }

	return { locked: words.has("protected"), hidden: words.has("formula-hidden") }
}

export class OdsCellStyles {
	readonly table: StyleTable
	private readonly sheet: OdfStyleSheet
	private readonly resolver: OdfStyleResolver
	private readonly issues: OdsIssues
	private readonly xfs = new Map<string, number>()
	private readonly columns = new Map<string, OdsColumnStyle>()
	private readonly rows = new Map<string, OdsRowStyle>()
	private readonly tables = new Map<string, OdsTableStyle>()
	private readonly textFonts = new Map<string, XlsxFont>()
	// The cell style a cell without one takes: "Default" when the file has it (LibreOffice writes it).
	private readonly fallback: string | undefined
	// The default cell style's style:decimal-places (§20.258): how many decimals General shows, -1 for no limit.
	readonly standardPrecision: number

	constructor(sheet: OdfStyleSheet, resolver: OdfStyleResolver, issues: OdsIssues) {
		this.sheet = sheet
		this.resolver = resolver
		this.issues = issues
		this.fallback = sheet.commonStyle("table-cell", "Default") === undefined ? undefined : "Default"

		const places = Number(this.resolver.resolve("table-cell", this.fallback, "content").get("table-cell", "style:decimal-places"))

		this.standardPrecision = Number.isInteger(places) && places >= 0 ? places : -1
		this.table = new StyleTable({
			numFmts: new Map(),
			fonts: [],
			fills: [
				{ kind: "pattern", pattern: "none" },
				{ kind: "pattern", pattern: "gray125" }
			],
			borders: [{}],
			cellStyleXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: -1, flags: 0 }],
			cellXfs: [],
			cellStyles: [{ name: "Normal", xfId: 0, builtinId: 0 }],
			dxfs: [],
			indexedColors: null
		})

		const base = this.build(this.resolver.resolve("table-cell", this.fallback, "content"))

		this.table.fonts.push(this.font(this.resolver.resolve("table-cell", this.fallback, "content")))
		this.table.cellXfs.push({ ...base, fontId: 0 })
		this.xfs.set("", 0)
	}

	// The cellXfs index of a cell style as the body names it (undefined: the default cell style).
	xf(name: string | undefined): number {
		const key = name ?? ""
		let index = this.xfs.get(key)

		if (index === undefined) {
			const resolved = this.resolver.resolve("table-cell", name ?? this.fallback, "content")

			index = this.table.appendXf(this.build(resolved))
			this.xfs.set(key, index)
		}

		return index
	}

	// A cell format for a dxf-like look: the cell style's font, fill and border alone (conditional formats).
	resolved(name: string): ResolvedOdfStyle {
		return this.resolver.resolve("table-cell", name, "content")
	}

	font(style: ResolvedOdfStyle): XlsxFont {
		return this.fontOf(style.properties("text"))
	}

	fontOf(text: OdfProperties): XlsxFont {
		const face = text.attributes.get("style:font-name")
		const family = (face === undefined ? undefined : this.sheet.fontFaces.get(face)?.family) ?? text.attributes.get("fo:font-family")
		const size = parseLength(text.attributes.get("fo:font-size"))
		const weight = text.attributes.get("fo:font-weight")
		const position = text.attributes.get("style:text-position")
		const underline = text.attributes.get("style:text-underline-style")
		const strike = text.attributes.get("style:text-line-through-style")
		const color = argb(text.attributes.get("fo:color"))
		const font: { -readonly [K in keyof XlsxFont]: XlsxFont[K] } = {
			name: family?.replace(/^['"]|['"]$/g, "") ?? DEFAULT_FONT,
			size: size ?? DEFAULT_SIZE
		}

		if (weight === "bold" || Number(weight) >= 600) font.bold = true
		if (text.attributes.get("fo:font-style") === "italic" || text.attributes.get("fo:font-style") === "oblique") font.italic = true
		if (underline !== undefined && underline !== "none") {
			font.underline = text.attributes.get("style:text-underline-type") === "double" ? "double" : "single"
		}
		if (strike !== undefined && strike !== "none") font.strike = true
		if (color !== undefined) font.color = { argb: color }
		if (text.attributes.get("style:text-outline") === "true") font.outline = true
		if (text.attributes.get("fo:text-shadow") !== undefined && text.attributes.get("fo:text-shadow") !== "none") font.shadow = true
		if (position !== undefined) {
			const first = position.trim().split(/\s+/)[0] ?? ""

			if (first === "super" || (first.endsWith("%") && Number(first.slice(0, -1)) > 0)) font.vertAlign = "superscript"
			else if (first === "sub" || (first.endsWith("%") && Number(first.slice(0, -1)) < 0)) font.vertAlign = "subscript"
		}

		return font
	}

	// A text style's own properties (family text, a span's) as a run font: only what it sets, the cell's font
	// giving the rest.
	textFont(name: string): XlsxFont {
		let font = this.textFonts.get(name)

		if (font === undefined) {
			const properties = this.resolver.resolve("text", name, "content").properties("text")
			const full = this.fontOf(properties)
			const own: { -readonly [K in keyof XlsxFont]: XlsxFont[K] } = { ...full }

			if (!properties.attributes.has("style:font-name") && !properties.attributes.has("fo:font-family")) delete own.name
			if (!properties.attributes.has("fo:font-size")) delete own.size

			font = own
			this.textFonts.set(name, font)
		}

		return font
	}

	fill(cell: OdfProperties): XlsxFill | undefined {
		const background = cell.attributes.get("fo:background-color")
		const color = argb(background)

		return color === undefined ? undefined : { kind: "pattern", pattern: "solid", fg: { argb: color }, bg: { indexed: 64 } }
	}

	border(cell: OdfProperties): XlsxBorder {
		const edge = (side: string): XlsxBorderEdge | undefined =>
			borderEdge(cell.attributes.get(`fo:border-${side}`), cell.attributes.get(`style:border-line-width-${side}`))
		const border: { -readonly [K in keyof XlsxBorder]: XlsxBorder[K] } = {}
		const top = edge("top")
		const bottom = edge("bottom")
		const left = edge("left")
		const right = edge("right")
		const up = borderEdge(cell.attributes.get("style:diagonal-bl-tr"), cell.attributes.get("style:diagonal-bl-tr-widths"))
		const down = borderEdge(cell.attributes.get("style:diagonal-tl-br"), cell.attributes.get("style:diagonal-tl-br-widths"))

		if (left !== undefined) border.left = left
		if (right !== undefined) border.right = right
		if (top !== undefined) border.top = top
		if (bottom !== undefined) border.bottom = bottom

		const diagonal = up ?? down

		if (diagonal !== undefined) {
			border.diagonal = diagonal
			border.diagonalUp = up !== undefined
			border.diagonalDown = down !== undefined
		}

		return border
	}

	alignment(paragraph: OdfProperties, cell: OdfProperties): XlsxAlignment | undefined {
		const out: { -readonly [K in keyof XlsxAlignment]: XlsxAlignment[K] } = {}
		const across = horizontal(paragraph, cell)
		const down = vertical(cell)
		const angle = parseAngle(cell.attributes.get("style:rotation-angle"))
		const indent = parseLength(paragraph.attributes.get("fo:margin-left"))

		if (across !== undefined) out.horizontal = across
		if (down !== undefined) out.vertical = down
		if (cell.attributes.get("fo:wrap-option") === "wrap") out.wrapText = true
		if (cell.attributes.get("style:shrink-to-fit") === "true") out.shrinkToFit = true
		if (cell.attributes.get("style:direction") === "ttb") out.textRotation = 255
		else if (angle !== undefined) {
			const degrees = ((Math.round(angle) % 360) + 360) % 360

			// SpreadsheetML: 0-90 counterclockwise, 91-180 for 1-90 clockwise (§18.8.1 textRotation).
			if (degrees > 0 && degrees <= 90) out.textRotation = degrees
			else if (degrees >= 270) out.textRotation = 90 + (360 - degrees)
			else if (degrees !== 0) this.issues.add("rotation", undefined, String(degrees))
		}

		// An indent level is three widths of the default font's digit (about 3 × 7 px).
		if (indent !== undefined && indent > 0) out.indent = Math.max(1, Math.round(pointsToPixels(indent) / 21))
		if (cell.attributes.get("style:writing-mode") === "rl-tb" || paragraph.attributes.get("style:writing-mode") === "rl-tb") {
			out.readingOrder = 2
		}

		return Object.keys(out).length === 0 ? undefined : out
	}

	numFmt(style: ResolvedOdfStyle): number {
		let name: string | undefined

		for (let k = style.chain.length - 1; k >= 0 && name === undefined; k--)
			name = style.chain[k]?.attributes.get("style:data-style-name")

		if (name === undefined || name === "") return 0

		const data = this.sheet.dataStyle(name, "content")

		if (data === undefined) return 0

		const converted = dataStyleCode(this.sheet, data, "content")

		for (const held of converted.held) this.issues.add("dataStyle", undefined, held)

		return converted.code === "General" ? 0 : this.table.numFmtId(converted.code)
	}

	private build(style: ResolvedOdfStyle): XlsxXf {
		// ODF 1.2's conditional cell styles (style:map on a cell style, §16.3) are not applied: named.
		for (const link of style.chain)
			for (const map of link.maps) this.issues.add("conditionalFormat", undefined, `style:map ${map.condition}`)

		const cell = style.properties("table-cell")
		const paragraph = style.properties("paragraph")
		const fill = this.fill(cell)
		const border = this.border(cell)
		const alignment = this.alignment(paragraph, cell)
		const fontId = this.table.fonts.length === 0 ? 0 : this.table.appendFont(this.font(style))

		return {
			numFmtId: this.numFmt(style),
			fontId,
			fillId: fill === undefined ? 0 : this.table.appendFill(fill),
			borderId: Object.keys(border).length === 0 ? 0 : this.table.appendBorder(border),
			xfId: 0,
			flags: XfApply.ALIGNMENT | XfApply.PROTECTION,
			...(alignment !== undefined && { alignment }),
			protection: protection(cell.attributes.get("style:cell-protect"))
		}
	}

	column(name: string | undefined): OdsColumnStyle {
		const key = name ?? ""
		let style = this.columns.get(key)

		if (style === undefined) {
			const properties = this.resolver.resolve("table-column", name, "content").properties("table-column")
			const width = parseLength(properties.attributes.get("style:column-width"))

			style = {
				widthPx: width === undefined ? Number.NaN : Math.round(pointsToPixels(width)),
				pageBreak: properties.attributes.get("fo:break-before") === "page"
			}
			this.columns.set(key, style)
		}

		return style
	}

	row(name: string | undefined): OdsRowStyle {
		const key = name ?? ""
		let style = this.rows.get(key)

		if (style === undefined) {
			const properties = this.resolver.resolve("table-row", name, "content").properties("table-row")
			const height = parseLength(properties.attributes.get("style:row-height") ?? properties.attributes.get("style:min-row-height"))

			style = {
				heightPt: height ?? Number.NaN,
				custom: properties.attributes.get("style:use-optimal-row-height") === "false",
				pageBreak: properties.attributes.get("fo:break-before") === "page"
			}
			this.rows.set(key, style)
		}

		return style
	}

	sheetStyle(name: string | undefined): OdsTableStyle {
		const key = name ?? ""
		let style = this.tables.get(key)

		if (style === undefined) {
			const properties = this.resolver.resolve("table", name, "content").properties("table")
			const tab = properties.attributes.get("table:tab-color") ?? properties.attributes.get("tableooo:tab-color")

			style = {
				display: properties.attributes.get("table:display") !== "false",
				rightToLeft: properties.attributes.get("style:writing-mode") === "rl-tb",
				tabColor: parseColor(tab) ?? -1
			}
			this.tables.set(key, style)
		}

		return style
	}
}
