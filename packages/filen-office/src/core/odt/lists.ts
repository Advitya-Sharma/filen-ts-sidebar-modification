import { parseLength, type OdfListLevel, type OdfListStyle, type OdfStyleSheet, type StyleScope } from "../odf/index"
import { fontFamily, type PropsSink } from "./props"
import { esc, twips } from "./wml"

// ODF list styles (Part 3 §16.30) as WordprocessingML abstract numbering, and list instances as w:num:
// text:list starts a new list unless it continues one (text:continue-numbering, text:continue-list), and a
// list item's text:start-value restarts the count through a level override.

const FORMATS: Readonly<Record<string, string>> = {
	"1": "decimal",
	a: "lowerLetter",
	A: "upperLetter",
	i: "lowerRoman",
	I: "upperRoman",
	"": "none",
	"1st": "ordinal",
	"١, ٢, ٣, ...": "arabicAlpha",
	"一, 二, 三, ...": "chineseCounting",
	"①, ②, ③, ...": "decimalEnclosedCircle",
	"Ⅰ, Ⅱ, Ⅲ, ...": "upperRoman",
	"ⅰ, ⅱ, ⅲ, ...": "lowerRoman"
}

// Where a level puts its label and text (§19.880 text:list-level-position-and-space-mode).
export interface LevelGeometry {
	readonly mode: "label-alignment" | "label-width-and-position"
	// label-alignment: absolute text indent and first-line offset (negative: hanging).
	readonly marginLeft: number
	readonly textIndent: number
	// label-width-and-position: offsets from the paragraph's own left margin.
	readonly spaceBefore: number
	readonly minLabelWidth: number
}

export function levelGeometry(level: OdfListLevel | undefined): LevelGeometry {
	const properties = level?.properties.attributes
	const mode =
		properties?.get("text:list-level-position-and-space-mode") === "label-alignment" ? "label-alignment" : "label-width-and-position"
	const alignment = level?.labelAlignment

	return {
		mode,
		marginLeft: parseLength(alignment?.get("fo:margin-left")) ?? 0,
		textIndent: parseLength(alignment?.get("fo:text-indent")) ?? 0,
		spaceBefore: parseLength(properties?.get("text:space-before")) ?? 0,
		minLabelWidth: parseLength(properties?.get("text:min-label-width")) ?? 0
	}
}

// A list paragraph's indents where its own do not win: absolute left and first-line offset.
export function listIndent(geometry: LevelGeometry, paragraphLeft: number): { readonly left: number; readonly firstLine: number } {
	if (geometry.mode === "label-alignment") return { left: geometry.marginLeft, firstLine: geometry.textIndent }

	return { left: paragraphLeft + geometry.spaceBefore + geometry.minLabelWidth, firstLine: -geometry.minLabelWidth }
}

interface ListInstance {
	readonly style: string
	readonly scope: StyleScope
	readonly abstract: number
	num: number
}

export class Lists {
	private readonly sheet: OdfStyleSheet
	private readonly abstracts: string[] = []
	private readonly nums: string[] = []
	private readonly abstractIds = new Map<string, number>()
	// The last instance of each list style (text:continue-numbering), and instances by xml:id.
	private readonly lastOf = new Map<string, ListInstance>()
	private readonly byId = new Map<string, ListInstance>()
	private outlineInstance: ListInstance | undefined

	constructor(sheet: OdfStyleSheet) {
		this.sheet = sheet
	}

	get empty(): boolean {
		return this.nums.length === 0
	}

	listStyle(name: string, scope: StyleScope): OdfListStyle | undefined {
		return this.sheet.listStyle(name, scope)
	}

	// A text:list starting at the top: a new list instance, or the one it continues.
	open(
		style: string,
		scope: StyleScope,
		continueNumbering: boolean,
		continueList: string | undefined,
		xmlId: string | undefined,
		sink: PropsSink
	): ListInstance | undefined {
		const list = this.sheet.listStyle(style, scope)

		if (list === undefined) return undefined

		const key = `${scope}\u0000${style}`
		let instance = continueList === undefined ? undefined : this.byId.get(continueList)

		if (instance === undefined && continueNumbering) instance = this.lastOf.get(key)
		if (instance === undefined) instance = this.newInstance(list, style, scope, sink)

		this.lastOf.set(key, instance)

		if (xmlId !== undefined) this.byId.set(xmlId, instance)

		return instance
	}

	// Headings numbered by the outline style (text:outline-style), when it numbers `level`.
	outline(level: number, sink: PropsSink): ListInstance | undefined {
		const outline = this.sheet.outlineStyle
		const levelStyle = outline?.levels.get(level)

		if (outline === undefined || levelStyle === undefined) return undefined

		const format = levelStyle.attributes.get("style:num-format") ?? ""
		const prefix = levelStyle.attributes.get("style:num-prefix") ?? ""
		const suffix = levelStyle.attributes.get("style:num-suffix") ?? ""

		if (format === "" && prefix === "" && suffix === "") return undefined

		this.outlineInstance ??= this.newInstance(outline, "\u0000outline", "styles", sink)

		return this.outlineInstance
	}

	// The instance restarted at `start` on `level` (a list item's text:start-value).
	restart(instance: ListInstance, ilvl: number, start: number): void {
		instance.num = this.addNum(instance.abstract, `<w:lvlOverride w:ilvl="${ilvl}"><w:startOverride w:val="${start}"/></w:lvlOverride>`)
	}

	level(instance: ListInstance, level: number): OdfListLevel | undefined {
		const list = instance.style === "\u0000outline" ? this.sheet.outlineStyle : this.sheet.listStyle(instance.style, instance.scope)

		return list?.levels.get(level)
	}

	numberingXml(): string {
		return `${this.abstracts.join("")}${this.nums.join("")}`
	}

	private newInstance(list: OdfListStyle, style: string, scope: StyleScope, sink: PropsSink): ListInstance {
		const key = `${scope}\u0000${style}`
		let abstract = this.abstractIds.get(key)

		if (abstract === undefined) {
			abstract = this.abstracts.length
			this.abstractIds.set(key, abstract)
			this.abstracts.push(this.abstractXml(list, abstract, sink))
		}

		return { style, scope, abstract, num: this.addNum(abstract, "") }
	}

	private addNum(abstract: number, overrides: string): number {
		const id = this.nums.length + 1

		this.nums.push(`<w:num w:numId="${id}"><w:abstractNumId w:val="${abstract}"/>${overrides}</w:num>`)

		return id
	}

	private abstractXml(list: OdfListStyle, id: number, sink: PropsSink): string {
		let out = `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="multilevel"/>`

		for (const level of list.levels.keys()) if (level > 9) sink.lossy("listLevel")

		for (let level = 1; level <= 9; level++) out += this.levelXml(list.levels.get(level), level, sink)

		return `${out}</w:abstractNum>`
	}

	private levelXml(level: OdfListLevel | undefined, number: number, sink: PropsSink): string {
		let format = "none"
		let text = ""
		let font = ""

		if (level?.kind === "bullet") {
			format = "bullet"
			text = level.attributes.get("text:bullet-char") ?? "•"

			const family = level.textProperties === undefined ? undefined : fontFamily(this.sheet, level.textProperties.attributes, "")

			if (family !== undefined) font = `<w:rFonts w:ascii="${esc(family)}" w:hAnsi="${esc(family)}" w:cs="${esc(family)}"/>`
		} else if (level?.kind === "image") {
			format = "bullet"
			text = "•"
		} else if (level !== undefined) {
			const numFormat = level.attributes.get("style:num-format") ?? "1"

			format = FORMATS[numFormat] ?? "decimal"

			if (FORMATS[numFormat] === undefined) sink.lossy("listLevel")

			const display = Math.max(1, Math.min(number, Number(level.attributes.get("text:display-levels") ?? "1") || 1))
			const parts: string[] = []

			// Levels shown before this one use their own formats (w:isLgl is not needed).
			for (let k = number - display + 1; k <= number; k++) parts.push(k === number && format === "none" ? "" : `%${k}`)

			text = `${level.attributes.get("style:num-prefix") ?? ""}${parts.filter(part => part !== "").join(".")}${level.attributes.get("style:num-suffix") ?? ""}`
		}

		const geometry = levelGeometry(level)
		let indent: string
		let suffix = ""
		let tab = ""

		if (geometry.mode === "label-alignment") {
			const follow = level?.labelAlignment?.get("text:label-followed-by") ?? "listtab"

			indent = `<w:ind w:left="${twips(geometry.marginLeft)}" ${geometry.textIndent < 0 ? `w:hanging="${twips(-geometry.textIndent)}"` : `w:firstLine="${twips(geometry.textIndent)}"`}/>`

			if (follow === "space") suffix = '<w:suff w:val="space"/>'
			else if (follow === "nothing" || follow === "newline") suffix = '<w:suff w:val="nothing"/>'

			if (follow === "newline") sink.lossy("listLevel")

			const stop = parseLength(level?.labelAlignment?.get("text:list-tab-stop-position"))

			if (follow === "listtab" && stop !== undefined) tab = `<w:tabs><w:tab w:val="num" w:pos="${twips(stop)}"/></w:tabs>`
		} else {
			indent = `<w:ind w:left="${twips(geometry.spaceBefore + geometry.minLabelWidth)}" w:hanging="${twips(geometry.minLabelWidth)}"/>`
		}

		const align = level?.properties.attributes.get("fo:text-align")
		const jc = align === "center" ? "center" : align === "end" || align === "right" ? "right" : "left"
		const start = Number(level?.attributes.get("text:start-value") ?? "1")
		const size = parseLength(level?.textProperties?.attributes.get("fo:font-size"))
		const rPr = `${font}${size === undefined ? "" : `<w:sz w:val="${Math.round(size * 2)}"/>`}`

		return `<w:lvl w:ilvl="${number - 1}"><w:start w:val="${Number.isFinite(start) && start >= 0 ? Math.floor(start) : 1}"/><w:numFmt w:val="${format}"/>${suffix}<w:lvlText w:val="${esc(text)}"/><w:lvlJc w:val="${jc}"/><w:pPr>${tab}${indent}</w:pPr>${rPr === "" ? "" : `<w:rPr>${rPr}</w:rPr>`}</w:lvl>`
	}
}
