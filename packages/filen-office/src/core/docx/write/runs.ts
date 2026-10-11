import type { Mark, MarkType } from "prosemirror-model"
import { docxSchema, type ColorAttrs } from "../edit/schema"
import type { RunBase } from "../edit/formatting"
import { partReader } from "../model/reader"
import { decodeUtf8 } from "../../xml/index"
import { Ns } from "../model/namespaces"
import { Names, Out, freshContainer, spliceChildren, startTagWith } from "./markup"
import { HIGHLIGHT, RUN_PROPERTIES, SHADING, THEME_COLOR, UNDERLINE } from "./order.generated"
import { color, integer, name, oneOf, optionalInteger, optionalOneOf } from "./values"

const marks = docxSchema.marks

// The w:rPr elements each formatting mark stands for.
const ELEMENTS: ReadonlyMap<MarkType, readonly string[]> = new Map<MarkType, readonly string[]>([
	[marks.char_style, ["rStyle"]],
	[marks.bold, ["b", "bCs"]],
	[marks.italic, ["i", "iCs"]],
	[marks.underline, ["u"]],
	[marks.strike, ["strike"]],
	[marks.double_strike, ["dstrike"]],
	[marks.caps, ["caps"]],
	[marks.small_caps, ["smallCaps"]],
	[marks.vert_align, ["vertAlign"]],
	[marks.color, ["color"]],
	[marks.highlight, ["highlight"]],
	[marks.shading, ["shd"]],
	[marks.font, ["rFonts"]],
	[marks.size, ["sz", "szCs"]],
	[marks.spacing, ["spacing"]],
	[marks.lang, ["lang"]]
])

export const PROPERTY_MARKS: readonly MarkType[] = [...ELEMENTS.keys()]

function hex2(value: number | null): string | null {
	return value === null ? null : value.toString(16).toUpperCase().padStart(2, "0")
}

function onOff(names: Names, local: string, value: unknown): string {
	return value === true ? names.empty(local) : value === false ? names.val(local, "0") : ""
}

function colorAttributes(color: ColorAttrs | null, prefix: "" | "fill"): (readonly [string, string | null])[] {
	if (color === null) return []

	if (prefix === "fill") {
		return [
			["fill", color.value ?? "auto"],
			["themeFill", color.themeColor],
			["themeFillTint", hex2(color.themeTint)],
			["themeFillShade", hex2(color.themeShade)]
		]
	}

	return [
		["color", color.value ?? "auto"],
		["themeColor", color.themeColor],
		["themeTint", hex2(color.themeTint)],
		["themeShade", hex2(color.themeShade)]
	]
}

// The markup of each element a mark stands for ("" for elements it leaves out); a missing mark leaves
// every element out.
export function elementsOf(names: Names, type: MarkType, mark: Mark | undefined): Map<string, string> {
	const out = new Map<string, string>()
	const locals = ELEMENTS.get(type) ?? []

	for (const local of locals) out.set(local, "")

	if (mark === undefined) return out

	const a = mark.attrs

	switch (type) {
		case marks.char_style:
			out.set("rStyle", names.val("rStyle", a["styleId"] as string))

			break
		case marks.bold:
		case marks.italic: {
			const [main, cs] = locals as [string, string]

			out.set(main, onOff(names, main, a["value"]))
			out.set(cs, onOff(names, cs, a["cs"]))

			break
		}
		case marks.underline: {
			const color = {
				value: a["value"],
				themeColor: a["themeColor"],
				themeTint: a["themeTint"],
				themeShade: a["themeShade"]
			} as ColorAttrs
			const colorPart = color.value === null && color.themeColor === null ? [] : colorAttributes(color, "")

			out.set("u", names.empty("u", [["val", a["style"] as string], ...colorPart]))

			break
		}
		case marks.strike:
		case marks.double_strike:
		case marks.caps:
		case marks.small_caps: {
			const local = locals[0] ?? ""

			out.set(local, onOff(names, local, a["value"]))

			break
		}
		case marks.vert_align:
			out.set("vertAlign", names.val("vertAlign", a["value"] as string))

			break
		case marks.color: {
			const [, ...theme] = colorAttributes(a as unknown as ColorAttrs, "")

			out.set("color", names.empty("color", [["val", (a["value"] as string | null) ?? "auto"], ...theme]))

			break
		}
		case marks.highlight:
			out.set("highlight", names.val("highlight", a["value"] as string))

			break
		case marks.shading:
			out.set(
				"shd",
				names.empty("shd", [
					["val", a["pattern"] as string],
					...colorAttributes(a["color"] as ColorAttrs | null, ""),
					...colorAttributes(a["fill"] as ColorAttrs | null, "fill")
				])
			)

			break
		case marks.font:
			out.set(
				"rFonts",
				names.empty("rFonts", [
					["ascii", a["ascii"] as string | null],
					["hAnsi", a["hAnsi"] as string | null],
					["eastAsia", a["eastAsia"] as string | null],
					["cs", a["cs"] as string | null],
					["asciiTheme", a["asciiTheme"] as string | null],
					["hAnsiTheme", a["hAnsiTheme"] as string | null],
					["eastAsiaTheme", a["eastAsiaTheme"] as string | null],
					["cstheme", a["csTheme"] as string | null],
					["hint", a["hint"] as string | null]
				])
			)

			break
		case marks.size:
			out.set("sz", a["value"] === null ? "" : names.val("sz", String(a["value"])))
			out.set("szCs", a["cs"] === null ? "" : names.val("szCs", String(a["cs"])))

			break
		case marks.spacing:
			out.set("spacing", names.val("spacing", String(a["value"])))

			break
		case marks.lang:
			out.set(
				"lang",
				names.empty("lang", [
					["val", a["latin"] as string | null],
					["eastAsia", a["eastAsia"] as string | null],
					["bidi", a["bidi"] as string | null]
				])
			)

			break
	}

	return out
}

// ST_VerticalAlignRun (shared-commonSimpleTypes.xsd).
const VERTICAL_ALIGN: ReadonlySet<string> = new Set(["baseline", "superscript", "subscript"])

// The values a mark's markup would carry, against their schema types. Font names, themes and hints only
// ever come from the source (Word writes w:hint="cs", which the 2016 schema lacks) and are not checked.
function checkMark(mark: Mark): void {
	const a = mark.attrs

	switch (mark.type) {
		case marks.char_style:
			name("w:rStyle", a["styleId"])

			break
		case marks.underline:
			optionalOneOf("w:u", UNDERLINE, a["style"])
			color(
				"w:u color",
				{ value: a["value"], themeColor: a["themeColor"], themeTint: a["themeTint"], themeShade: a["themeShade"] } as ColorAttrs,
				THEME_COLOR
			)

			break
		case marks.vert_align:
			oneOf("w:vertAlign", VERTICAL_ALIGN, a["value"])

			break
		case marks.color:
			color("w:color", a as unknown as ColorAttrs, THEME_COLOR)

			break
		case marks.highlight:
			oneOf("w:highlight", HIGHLIGHT, a["value"])

			break
		case marks.shading:
			oneOf("w:shd", SHADING, a["pattern"])
			color("w:shd color", a["color"] as ColorAttrs | null, THEME_COLOR)
			color("w:shd fill", a["fill"] as ColorAttrs | null, THEME_COLOR)

			break
		case marks.size:
			optionalInteger("w:sz", a["value"], 0)
			optionalInteger("w:szCs", a["cs"], 0)

			break
		case marks.spacing:
			integer("w:spacing", a["value"])

			break
	}
}

// Whether a node's formatting is its base's: then the base's w:rPr bytes stand as they are.
export function sameFormatting(nodeMarks: readonly Mark[], base: RunBase | undefined): boolean {
	if (marks.plain.isInSet(nodeMarks) !== undefined) return false

	for (const type of PROPERTY_MARKS) {
		const own = type.isInSet(nodeMarks)
		const from = base === undefined ? undefined : type.isInSet(base.marks)

		if (own === undefined ? from !== undefined : from === undefined || !own.eq(from)) return false
	}

	return true
}

// The w:rPr for content with `nodeMarks`: the base's bytes with only the differing properties spliced in;
// with the plain mark (or no base) new markup from the marks alone, keeping the base's w:rtl, w:cs and a
// tracked w:rPrChange. Nothing when there are no properties at all.
// `changeId`, when given, is a new w:id for the base's tracked formatting change (w:rPrChange): a source
// run written as several runs must not repeat a revision id.
export function writeRunProperties(out: Out, names: Names, nodeMarks: readonly Mark[], base: RunBase | undefined, changeId?: number): void {
	const plain = marks.plain.isInSet(nodeMarks) !== undefined
	const reader = base === undefined ? undefined : partReader(base.document)
	const change = base === undefined || changeId === undefined ? -1 : childOf(base, "rPrChange")
	const changeMarkup = (): string => {
		const sub = new Out()
		const document = (base as RunBase).document

		startTagWith(sub, document, change, names.attr("id"), String(changeId))
		sub.bytes(document.bytes, document.openEnd(change), document.end(change))

		if (document.isSelfClosing(change)) sub.text(`</${document.name(change)}>`)

		return decodeUtf8(sub.finish())
	}

	if (base !== undefined && reader !== undefined && !plain) {
		if (change < 0 && sameFormatting(nodeMarks, base)) {
			out.bytes(base.document.bytes, base.document.start(base.node), base.document.end(base.node))

			return
		}

		const changes = new Map<string, string>()

		if (change >= 0) changes.set("rPrChange", changeMarkup())

		for (const type of PROPERTY_MARKS) {
			const ownMark = type.isInSet(nodeMarks)
			const fromMark = type.isInSet(base.marks)

			if (ownMark === undefined ? fromMark === undefined : fromMark !== undefined && ownMark.eq(fromMark)) continue
			if (ownMark !== undefined) checkMark(ownMark)

			const own = elementsOf(names, type, ownMark)
			const from = elementsOf(names, type, fromMark)

			for (const [local, markup] of own) if (markup !== from.get(local)) changes.set(local, markup)
		}

		if (changes.size === 0) {
			out.bytes(base.document.bytes, base.document.start(base.node), base.document.end(base.node))

			return
		}

		spliceChildren(out, reader, base.node, changes, RUN_PROPERTIES)

		return
	}

	const children = new Map<string, string>()

	for (const type of PROPERTY_MARKS) {
		const mark = type.isInSet(nodeMarks)

		if (mark === undefined) continue

		checkMark(mark)

		for (const [local, markup] of elementsOf(names, type, mark)) if (markup !== "") children.set(local, markup)
	}

	if (base !== undefined) {
		// What clearing keeps from the source formatting: the run's direction and script, and a tracked change.
		const document = base.document
		const keep = new Set(["rtl", "cs", ...(marks.format_change.isInSet(nodeMarks) !== undefined ? ["rPrChange"] : [])])

		for (let child = document.firstChildElement(base.node); child >= 0; child = document.nextSiblingElement(child)) {
			const local = document.localName(child)

			if (keep.has(local) && reader?.group(child) === Ns.W) {
				children.set(
					local,
					child === change ? changeMarkup() : decodeUtf8(document.bytes, document.start(child), document.end(child))
				)
			}
		}
	}

	out.text(freshContainer(names, "rPr", children, RUN_PROPERTIES))
}

function childOf(base: RunBase, local: string): number {
	const document = base.document

	for (let child = document.firstChildElement(base.node); child >= 0; child = document.nextSiblingElement(child)) {
		if (document.localName(child) === local) return child
	}

	return -1
}
