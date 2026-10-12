import type { Mark, MarkType, Node } from "prosemirror-model"
import { NO_SOURCE, docxSchema } from "../edit/schema"
import type { DocxEditorSession } from "../edit/session"
import type { PagedEditor } from "../paged/editor"
import type { PagedGeometry } from "../paged/geometry"
import { numberSwitchFormat, numberText } from "../layout/fields"
import { FieldCharType, Item, type StoryPart } from "../model/index"
import { Ns } from "../model/namespaces"
import { bookmarkIndex, type DocxBookmark } from "./bookmarks"

// Field codes and what updating a field computes (ECMA-376 Part 1 §17.16): the instruction parsed into its
// type, arguments and switches (§17.16.1), the general formatting switches \* (§17.16.4.3), numeric pictures
// \# (§17.16.4.2) and date-time pictures \@ (§17.16.4.1), and the results of the fields an update recomputes
// from the document as laid out: PAGE, NUMPAGES, SECTIONPAGES, SECTION, DATE, TIME, REF, PAGEREF and SEQ.
// Every other field keeps the result Word stored (Word itself cannot compute most of them without data
// outside the file). Updates are explicit (Word's F9, Update Field, Update Table), never per keystroke.

const nodes = docxSchema.nodes

// ---- field codes ------------------------------------------------------------------------------------------

export interface FieldSwitch {
	// The switch character: "*", "#", "@", "!" or a letter as written.
	readonly name: string
	readonly arg: string | undefined
}

export interface FieldCode {
	// Upper case ("PAGEREF"); "" when the instruction is empty.
	readonly type: string
	readonly args: readonly string[]
	readonly switches: readonly FieldSwitch[]
}

// Field-specific switches that take an argument, per field type (the general ones always do).
const ARGUMENT_SWITCHES: Readonly<Record<string, string>> = {
	TOC: "abcdflnopst",
	SEQ: "rs",
	REF: "d",
	HYPERLINK: "lmot",
	STYLEREF: "",
	PAGEREF: "",
	NOTEREF: "",
	DATE: "",
	TIME: ""
}

interface Token {
	readonly value: string
	readonly quoted: boolean
	readonly switch: boolean
}

function tokenize(instruction: string): Token[] {
	const out: Token[] = []
	const length = instruction.length
	let at = 0

	while (at < length) {
		const char = instruction[at] ?? ""

		if (/\s/.test(char)) {
			at++

			continue
		}

		if (char === '"') {
			let value = ""

			at++

			while (at < length && instruction[at] !== '"') {
				const next = instruction[at + 1]

				// Inside quotes a backslash escapes a quote or a backslash (§17.16.1).
				if (instruction[at] === "\\" && (next === '"' || next === "\\")) {
					value += next
					at += 2
				} else {
					value += instruction[at] ?? ""
					at++
				}
			}

			at++
			out.push({ value, quoted: true, switch: false })

			continue
		}

		if (char === "\\" && at + 1 < length) {
			out.push({ value: instruction[at + 1] ?? "", quoted: false, switch: true })
			at += 2

			continue
		}

		let end = at

		while (end < length && !/\s/.test(instruction[end] ?? "")) end++

		out.push({ value: instruction.slice(at, end), quoted: false, switch: false })
		at = end
	}

	return out
}

const CODES = new Map<string, FieldCode>()

export function parseFieldCode(instruction: string): FieldCode {
	const cached = CODES.get(instruction)

	if (cached !== undefined) return cached

	const tokens = tokenize(instruction)
	const first = tokens[0]
	const type = first === undefined || first.switch ? "" : first.value.toUpperCase()
	const own = ARGUMENT_SWITCHES[type] ?? ""
	const args: string[] = []
	const switches: FieldSwitch[] = []

	for (let index = type === "" ? 0 : 1; index < tokens.length; index++) {
		const token = tokens[index] as Token

		if (!token.switch) {
			args.push(token.value)

			continue
		}

		const name = token.value
		const next = tokens[index + 1]
		const takes = name === "*" || name === "#" || name === "@" || own.includes(name.toLowerCase())

		if (takes && next !== undefined && !next.switch) {
			switches.push({ name, arg: next.value })
			index++
		} else {
			switches.push({ name, arg: undefined })
		}
	}

	const code: FieldCode = Object.freeze({ type, args: Object.freeze(args), switches: Object.freeze(switches) })

	if (CODES.size > 4096) CODES.clear()

	CODES.set(instruction, code)

	return code
}

// A switch by its character (letters match in either case).
export function switchOf(code: FieldCode, name: string): FieldSwitch | undefined {
	const lower = name.toLowerCase()

	return code.switches.find(item => item.name.toLowerCase() === lower)
}

export function hasSwitch(code: FieldCode, name: string): boolean {
	return switchOf(code, name) !== undefined
}

// The \* arguments, in order.
function generalSwitches(code: FieldCode): string[] {
	return code.switches.filter(item => item.name === "*" && item.arg !== undefined).map(item => item.arg as string)
}

// Whether a \* MERGEFORMAT switch asks for the previous result's formatting (§17.16.4.3).
export function mergeFormat(code: FieldCode): boolean {
	return generalSwitches(code).some(name => name.toUpperCase() === "MERGEFORMAT")
}

// ---- formatting -------------------------------------------------------------------------------------------

// \* Upper, Lower, FirstCap and Caps (§17.16.4.3), applied to a result in the order written.
function caseFormat(text: string, code: FieldCode): string {
	let out = text

	for (const name of generalSwitches(code)) {
		switch (name.toLowerCase()) {
			case "upper":
				out = out.toUpperCase()

				break
			case "lower":
				out = out.toLowerCase()

				break
			case "firstcap":
				out = out.replace(/\p{L}/u, letter => letter.toUpperCase())

				break
			case "caps":
				out = out.replace(/(^|[^\p{L}\p{N}'’])(\p{L})/gu, (_match, before: string, letter: string) => before + letter.toUpperCase())

				break
		}
	}

	return out
}

// The ST_NumberFormat of the last \* switch that formats numbers, if any.
function numberFormatOf(code: FieldCode): string | undefined {
	let format: string | undefined

	for (const name of generalSwitches(code)) format = numberSwitchFormat(name) ?? format

	return format
}

// A number under a \# numeric picture (§17.16.4.2): 0 a digit (zero when none), # a digit (nothing when
// none), x drops the digits to its left, "." the decimal point, "," digit grouping, "-" and "+" signs, text
// in single quotes and any other character as written; up to three sections (positive;negative;zero).
export function numericPicture(value: number, picture: string): string {
	const sections = splitUnquoted(picture, ";")
	const section =
		value < 0 && sections.length > 1
			? (sections[1] ?? "")
			: value === 0 && sections.length > 2
				? (sections[2] ?? "")
				: (sections[0] ?? "")
	const negativeSection = value < 0 && sections.length > 1
	const magnitude = Math.abs(value)
	// The digit placeholders before and after the decimal point.
	const plain = section.replace(/'[^']*'/g, "")
	const point = plain.indexOf(".")
	const integerPart = point < 0 ? plain : plain.slice(0, point)
	const fractionPart = point < 0 ? "" : plain.slice(point + 1)
	const decimals = (fractionPart.match(/[0#x]/g) ?? []).length
	const zeros = (integerPart.match(/0/g) ?? []).length
	const grouping = integerPart.includes(",")
	const truncate = integerPart.lastIndexOf("x")
	const [whole = "0", fraction = ""] = magnitude.toFixed(decimals).split(".")
	let digits = whole.replace(/^0+(?=\d)/, "")

	if (digits === "0" && zeros === 0 && decimals > 0) digits = ""

	digits = digits.padStart(zeros, "0")

	if (truncate >= 0) {
		const kept = (integerPart.slice(truncate + 1).match(/[0#]/g) ?? []).length

		digits = digits.slice(Math.max(0, digits.length - kept))
	}

	if (grouping) digits = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",")

	// Optional trailing fraction digits (#) are dropped when zero.
	let fractionDigits = fraction
	const optional = (fractionPart.match(/#/g) ?? []).length

	for (let k = 0; k < optional && fractionDigits.endsWith("0"); k++) fractionDigits = fractionDigits.slice(0, -1)

	let out = ""
	let integerDone = false
	let fractionDone = false
	let signDone = false

	for (let at = 0; at < section.length; at++) {
		const char = section[at] ?? ""

		if (char === "'") {
			const close = section.indexOf("'", at + 1)

			out += section.slice(at + 1, close < 0 ? section.length : close)
			at = close < 0 ? section.length : close

			continue
		}

		if ((char === "0" || char === "#" || char === "x" || char === ",") && (point < 0 || at < section.indexOf("."))) {
			if (!integerDone) {
				out += digits
				integerDone = true
			}

			continue
		}

		if (char === "." && point >= 0) {
			if (fractionDigits !== "") out += "."

			continue
		}

		if (char === "0" || char === "#" || char === "x") {
			if (!fractionDone) {
				out += fractionDigits
				fractionDone = true
			}

			continue
		}

		if (char === "-" || char === "+") {
			if (!signDone && !negativeSection) out += value < 0 ? "-" : char === "+" && value > 0 ? "+" : ""

			signDone = true

			continue
		}

		out += char
	}

	if (value < 0 && !signDone && !negativeSection) out = `-${out}`

	return out
}

function splitUnquoted(text: string, separator: string): string[] {
	const out: string[] = []
	let quoted = false
	let current = ""

	for (const char of text) {
		if (char === "'") quoted = !quoted

		if (char === separator && !quoted) {
			out.push(current)
			current = ""
		} else {
			current += char
		}
	}

	out.push(current)

	return out
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
// Word's defaults when no \@ is given: the en-US short date and time.
export const DEFAULT_DATE_PICTURE = "M/d/yyyy"
export const DEFAULT_TIME_PICTURE = "h:mm am/pm"

// A date under a \@ date-time picture (§17.16.4.1): M months, d days, y years, h (12-hour) and H (24-hour)
// hours, m minutes, s seconds, am/pm and AM/PM, text in single quotes and other characters as written.
export function datePicture(date: Date, picture: string): string {
	let out = ""
	let at = 0
	const run = (char: string): number => {
		let end = at

		while (end < picture.length && picture[end] === char) end++

		return end - at
	}

	while (at < picture.length) {
		const char = picture[at] ?? ""

		if (char === "'") {
			const close = picture.indexOf("'", at + 1)

			out += picture.slice(at + 1, close < 0 ? picture.length : close)
			at = close < 0 ? picture.length : close + 1

			continue
		}

		const ampm = /^(am\/pm|AM\/PM|a\/p|A\/P)/.exec(picture.slice(at))

		if (ampm !== null) {
			const pm = date.getHours() >= 12
			const token = ampm[1] as string
			const text = token.length === 3 ? (pm ? "p" : "a") : pm ? "pm" : "am"

			out += token === token.toUpperCase() ? text.toUpperCase() : text
			at += token.length

			continue
		}

		const count = run(char)

		switch (char) {
			case "M":
				out +=
					count >= 4
						? MONTHS[date.getMonth()]
						: count === 3
							? MONTHS[date.getMonth()]?.slice(0, 3)
							: String(date.getMonth() + 1).padStart(count, "0")

				break
			case "d":
			case "D":
				out +=
					count >= 4
						? DAYS[date.getDay()]
						: count === 3
							? DAYS[date.getDay()]?.slice(0, 3)
							: String(date.getDate()).padStart(count, "0")

				break
			case "y":
			case "Y":
				out += count <= 2 ? String(date.getFullYear() % 100).padStart(2, "0") : String(date.getFullYear())

				break
			case "h":
				out += String(date.getHours() % 12 === 0 ? 12 : date.getHours() % 12).padStart(Math.min(count, 2), "0")

				break
			case "H":
				out += String(date.getHours()).padStart(Math.min(count, 2), "0")

				break
			case "m":
				out += String(date.getMinutes()).padStart(Math.min(count, 2), "0")

				break
			case "s":
			case "S":
				out += String(date.getSeconds()).padStart(Math.min(count, 2), "0")

				break
			default:
				out += char.repeat(count)
		}

		at += count
	}

	return out
}

// A number result: \# picture, else a \* number format, else `fallback` (the page's own format for pages).
function numberResult(value: number, code: FieldCode, fallback?: string): string {
	const picture = switchOf(code, "#")?.arg
	const format = numberFormatOf(code)
	const text =
		picture !== undefined
			? numericPicture(value, picture)
			: format !== undefined
				? numberText(value, format)
				: (fallback ?? String(value))

	return caseFormat(text, code)
}

// ---- the document as laid out -----------------------------------------------------------------------------

export interface FieldPage {
	// The page number fields show (section restarts applied) and its text in the section's format.
	readonly number: number
	readonly numberText: string
	readonly section: number
}

// What an update reads from the laid-out pages (undefined where nothing is laid out: a headless session).
export interface FieldPages {
	readonly count: number
	page(index: number): FieldPage | undefined
	// The page holding position `pos` of `session`'s document.
	pageAt(session: DocxEditorSession, pos: number): number | undefined
	// The list number drawn before the paragraph holding `pos` ("1.2\t"), "" for none.
	markerAt(session: DocxEditorSession, pos: number): string
	sectionPages(section: number): number
}

// Where fields are updated: the body's session, every session of the document and its pages.
export interface FieldHost {
	readonly body: { readonly session: DocxEditorSession }
	readonly sessions: readonly DocxEditorSession[]
	readonly pages: FieldPages | undefined
}

// ---- headings --------------------------------------------------------------------------------------------

// A paragraph's outline level (0 for Heading 1, undefined for body text), its style's unless its own
// properties or a list may set one; most paragraphs are resolved by their style alone.
export function outlineLevels(session: DocxEditorSession): (node: Node) => number | undefined {
	const formats = session.formats
	const part = session.story.part
	const styleLevels = new Map<string | null, number | undefined>()

	return node => {
		if (node.type !== nodes.paragraph) return undefined

		const src = node.attrs["src"] as number
		const row = src === NO_SOURCE ? -1 : (part.aux[src] ?? -1)
		const pPr = row < 0 ? -1 : (part.paragraphs.pPr[row] ?? -1)

		if (node.attrs["numId"] !== null || (pPr >= 0 && part.reader.child(pPr, Ns.W, "outlineLvl") >= 0)) {
			const level = formats.paragraph(node).props.outlineLevel

			return level === undefined || level >= 9 ? undefined : level
		}

		const styleId = node.attrs["styleId"] as string | null

		if (!styleLevels.has(styleId)) {
			const style = formats.styleParagraph(styleId)

			styleLevels.set(styleId, style.props.numbering === undefined ? style.props.outlineLevel : -1)
		}

		const level = styleLevels.get(styleId)
		const resolved = level === -1 ? formats.paragraph(node).props.outlineLevel : level

		return resolved === undefined || resolved >= 9 ? undefined : resolved
	}
}

// ---- text of a range --------------------------------------------------------------------------------------

const DELETED: readonly MarkType[] = [docxSchema.marks.deletion, docxSchema.marks.move_from]

function deleted(marks: readonly Mark[]): boolean {
	return DELETED.some(type => type.isInSet(marks) !== undefined)
}

// What [from, to) of `doc` reads as in the final view: text, tabs, breaks and field results, paragraphs
// joined by line breaks (a field result is inline content).
export function rangeText(doc: Node, from: number, to: number): string {
	let out = ""
	let paragraphs = 0

	doc.nodesBetween(Math.max(0, from), Math.min(to, doc.content.size), (node, pos) => {
		if (node.isTextblock) {
			if (paragraphs++ > 0 && pos >= from) out += "\n"

			return true
		}

		if (!node.isInline) return !node.isAtom

		if (deleted(node.marks)) return false

		const start = Math.max(from, pos)
		const end = Math.min(to, pos + node.nodeSize)

		if (node.isText) out += (node.text ?? "").slice(start - pos, end - pos)
		else if (node.type === nodes.field) out += node.attrs["result"] as string
		else if (node.type === nodes.tab) out += "\t"
		else if (node.type === nodes.break) out += node.attrs["type"] === "textWrapping" ? "\n" : ""
		else if (node.type === nodes.symbol) out += String.fromCodePoint(node.attrs["char"] as number)
		else if (node.type === nodes.hyphen && node.attrs["soft"] !== true) out += "‑"

		return false
	})

	return out.replace(/\n+$/, "")
}

// ---- source fields ---------------------------------------------------------------------------------------

// A source field's place in its part: its row in the part's field table and the items of its characters.
export interface SourceField {
	readonly field: number
	readonly simple: boolean
	// fldChar begin, separate (-1: none) and end items; a simple field's are its w:fldSimple.
	readonly begin: number
	readonly separate: number
	readonly end: number
}

// Item kinds a rewritten result may drop (the rest of a result, a note reference or a picture, would be lost):
// runs and their text, wrappers that hold nothing else, nested fields; range markers and comment references
// are written again next to the new result.
const RESULT_KINDS: ReadonlySet<number> = new Set<number>([
	Item.Run,
	Item.Text,
	Item.DeletedText,
	Item.InstrText,
	Item.DeletedInstrText,
	Item.Tab,
	Item.Break,
	Item.CarriageReturn,
	Item.LastRenderedPageBreak,
	Item.Symbol,
	Item.SoftHyphen,
	Item.NoBreakHyphen,
	Item.FieldChar,
	Item.FieldSimple,
	Item.Insertion,
	Item.Deletion,
	Item.SmartTag,
	Item.CustomXml,
	Item.Bidi,
	Item.Hyperlink,
	Item.CommentReference,
	Item.BookmarkStart,
	Item.BookmarkEnd,
	Item.CommentRangeStart,
	Item.CommentRangeEnd,
	Item.MoveFromRangeStart,
	Item.MoveFromRangeEnd,
	Item.MoveToRangeStart,
	Item.MoveToRangeEnd,
	Item.PermissionStart,
	Item.PermissionEnd,
	Item.CustomXmlRangeMarker
])

const NESTED = new WeakMap<StoryPart, ReadonlySet<number>>()

// Fields of `part` that hold others in their instruction or result.
function nestingFields(part: StoryPart): ReadonlySet<number> {
	let set = NESTED.get(part)

	if (set === undefined) {
		const out = new Set<number>()
		const parents = part.fields.parent

		for (let field = 0; field < part.fields.count; field++) {
			const parent = parents[field] ?? -1

			if (parent >= 0) out.add(parent)
		}

		set = out
		NESTED.set(part, set)
	}

	return set
}

// The source field a field atom stands for, when a new result can be written in place of its stored one: its
// characters are where the atom's runs are (the separate in a run directly in the paragraph or wrapper that
// holds the field), it is not locked (w:fldLock), holds no other field and its result holds only what a
// rewrite may drop.
export function rewritableSource(part: StoryPart, node: Node): SourceField | undefined {
	const first = node.attrs["src"] as number
	const last = node.attrs["end"] as number

	if (first === NO_SOURCE || first < 0 || first >= part.count) return undefined

	const simple = part.kinds[first] === Item.FieldSimple
	const field = simple ? (part.aux[first] ?? -1) : part.fieldOf(first + 1)

	if (field < 0 || nestingFields(part).has(field)) return undefined

	if (simple) {
		const locked = part.reader.attr(part.node(first), Ns.W, "fldLock")

		if (locked === "1" || locked === "true" || locked === "on") return undefined

		for (let item = first + 1; item < part.end(first); item++) if (!RESULT_KINDS.has(part.kinds[item] ?? 0)) return undefined

		return { field, simple, begin: first, separate: -1, end: first }
	}

	const begin = part.fields.begin[field] ?? -1
	const separate = part.fields.separate[field] ?? -1
	const end = part.fields.end[field] ?? -1

	if (begin !== first + 1 || end < 0 || part.parent(end) !== last || part.fieldCharLocked(begin)) return undefined
	if (part.fieldCharType(begin) !== FieldCharType.Begin) return undefined

	if (separate >= 0) {
		const run = part.parent(separate)

		if (part.kinds[run] !== Item.Run || part.parent(run) !== part.parent(first)) return undefined
	}

	for (let item = (separate >= 0 ? separate : begin) + 1; item < end; item++)
		if (!RESULT_KINDS.has(part.kinds[item] ?? 0)) return undefined

	return { field, simple, begin, separate, end }
}

// ---- evaluation ------------------------------------------------------------------------------------------

// Field types an update recomputes.
export const UPDATABLE_FIELDS: ReadonlySet<string> = new Set([
	"PAGE",
	"NUMPAGES",
	"SECTIONPAGES",
	"SECTION",
	"DATE",
	"TIME",
	"REF",
	"PAGEREF",
	"SEQ"
])

// Fields whose result depends on page numbers.
const PAGE_FIELDS: ReadonlySet<string> = new Set(["PAGE", "NUMPAGES", "SECTIONPAGES", "SECTION", "PAGEREF"])

// What Word writes for a cross-reference to a bookmark the document lacks (en-US).
export const MISSING_BOOKMARK = "Error! Bookmark not defined."

export function needsPages(type: string): boolean {
	return PAGE_FIELDS.has(type)
}

// Whether field atom `node` (of `session`'s story) is one an update recomputes.
export function updatableField(session: DocxEditorSession, node: Node): boolean {
	if (node.type !== nodes.field || !UPDATABLE_FIELDS.has(parseFieldCode(node.attrs["instruction"] as string).type)) return false

	return (node.attrs["src"] as number) === NO_SOURCE || rewritableSource(session.story.part, node) !== undefined
}

// One update's view of a story: its bookmarks, sequence numbers and the clock, made once per update.
export class FieldEvaluation {
	readonly session: DocxEditorSession
	readonly host: FieldHost
	// The story's document the positions are in (the session's, or a transaction's about to apply).
	readonly doc: Node
	private readonly now: Date
	private sequences: Map<number, number> | undefined = undefined

	constructor(host: FieldHost, session: DocxEditorSession, now: number, doc: Node = session.state.doc) {
		this.host = host
		this.session = session
		this.doc = doc
		this.now = new Date(now)
	}

	private docOf(session: DocxEditorSession): Node {
		return session === this.session ? this.doc : session.state.doc
	}

	private bookmark(name: string | undefined): DocxBookmark | undefined {
		if (name === undefined) return undefined

		const key = name.toLowerCase()
		const own = bookmarkIndex(this.doc, this.session.story.part).byName.get(key)

		if (own !== undefined) return own

		const body = this.host.body.session

		return body === this.session ? undefined : bookmarkIndex(body.state.doc, body.story.part).byName.get(key)
	}

	// The session whose document holds `bookmark` (this story's, else the body's).
	private holder(bookmark: DocxBookmark): DocxEditorSession {
		return bookmarkIndex(this.doc, this.session.story.part).list.includes(bookmark) ? this.session : this.host.body.session
	}

	private pageOf(session: DocxEditorSession, pos: number): FieldPage | undefined {
		const pages = this.host.pages
		const index = pages?.pageAt(session, pos)

		return index === undefined ? undefined : pages?.page(index)
	}

	// The new result of the field atom `node` at `pos`, or undefined when it keeps the one it has.
	result(node: Node, pos: number): string | undefined {
		const code = parseFieldCode(node.attrs["instruction"] as string)
		const pages = this.host.pages

		switch (code.type) {
			case "PAGE": {
				const page = this.pageOf(this.session, pos)

				return page === undefined ? undefined : numberResult(page.number, code, page.numberText)
			}
			case "NUMPAGES":
				return pages === undefined || pages.count === 0 ? undefined : numberResult(pages.count, code)
			case "SECTIONPAGES": {
				const page = this.pageOf(this.session, pos)

				return page === undefined || pages === undefined ? undefined : numberResult(pages.sectionPages(page.section), code)
			}
			case "SECTION": {
				const page = this.pageOf(this.session, pos)

				return page === undefined ? undefined : numberResult(page.section + 1, code)
			}
			case "DATE":
			case "TIME":
				return caseFormat(
					datePicture(this.now, switchOf(code, "@")?.arg ?? (code.type === "DATE" ? DEFAULT_DATE_PICTURE : DEFAULT_TIME_PICTURE)),
					code
				)
			case "REF":
				return this.reference(code, pos)
			case "PAGEREF":
				return this.pageReference(code, pos)
			case "SEQ":
				return this.sequence(code, pos)
			default:
				return undefined
		}
	}

	// "above" or "below": where a bookmark is from `pos` (§17.16.5.51 \p).
	private relative(bookmark: DocxBookmark, pos: number): string {
		return this.holder(bookmark) !== this.session || bookmark.from <= pos ? "above" : "below"
	}

	// REF (§17.16.5.51): the bookmarked text, or with \n, \r or \w the list number of its paragraph as drawn
	// (the number in full whichever of the three: the layout draws it so), \d its separator, \p "above"/"below".
	private reference(code: FieldCode, pos: number): string | undefined {
		const bookmark = this.bookmark(code.args[0])

		if (bookmark === undefined) return MISSING_BOOKMARK

		const numbered = hasSwitch(code, "n") || hasSwitch(code, "r") || hasSwitch(code, "w")
		let text: string

		if (numbered) {
			const session = this.holder(bookmark)
			const marker = this.host.pages?.markerAt(session, bookmark.from) ?? ""

			text = marker.replace(/[\t\s]+$/, "")

			if (hasSwitch(code, "d")) text = text.replace(/[.]/g, switchOf(code, "d")?.arg ?? ".")
		} else {
			const holder = this.holder(bookmark)

			text = rangeText(this.docOf(holder), bookmark.from, bookmark.to)
		}

		if (hasSwitch(code, "p")) text = numbered ? `${text} ${this.relative(bookmark, pos)}` : this.relative(bookmark, pos)

		return caseFormat(text, code)
	}

	// PAGEREF (§17.16.5.43): the page number of the bookmark's start in its section's format (or a \* format);
	// \p "on page N", or "above"/"below" on the same page.
	private pageReference(code: FieldCode, pos: number): string | undefined {
		const bookmark = this.bookmark(code.args[0])

		if (bookmark === undefined) return MISSING_BOOKMARK

		const target = this.pageOf(this.holder(bookmark), bookmark.from)

		if (target === undefined) return undefined

		const text = numberResult(target.number, code, target.numberText)

		if (!hasSwitch(code, "p")) return text

		const here = this.host.pages?.pageAt(this.session, pos)
		const there = this.host.pages?.pageAt(this.holder(bookmark), bookmark.from)

		return here !== undefined && here === there ? this.relative(bookmark, pos) : `on page ${text}`
	}

	// SEQ (§17.16.5.56): the fields of the story in document order, each identifier counted on its own; \c
	// repeats the last number, \r n restarts at n, \s n restarts after each heading of level n or above,
	// \h hides the result.
	private sequence(code: FieldCode, pos: number): string | undefined {
		const value = this.sequenceValues().get(pos)

		if (value === undefined) return undefined
		if (hasSwitch(code, "h") && generalSwitches(code).length === 0) return ""

		return numberResult(value, code)
	}

	private sequenceValues(): Map<number, number> {
		if (this.sequences !== undefined) return this.sequences

		const values = new Map<number, number>()
		const counters = new Map<string, number>()
		const doc = this.doc
		const fields: { pos: number; code: FieldCode }[] = []
		const resets = new Map<string, number>()

		doc.descendants((node, pos) => {
			if (node.type === nodes.field) {
				const code = parseFieldCode(node.attrs["instruction"] as string)

				if (code.type === "SEQ" && !deleted(node.marks)) fields.push({ pos, code })

				return false
			}

			return !node.isLeaf
		})

		for (const { code } of fields) {
			const level = Number(switchOf(code, "s")?.arg)
			const id = (code.args[0] ?? "").toLowerCase()

			if (Number.isInteger(level) && level >= 1 && level <= 9) resets.set(id, Math.max(resets.get(id) ?? 0, level))
		}

		// Headings restarting a sequence, by position: only looked for when a \s asks.
		const headings: { pos: number; level: number }[] = []

		if (resets.size > 0) {
			const levelOf = outlineLevels(this.session)

			doc.descendants((node, pos) => {
				if (node.type !== nodes.paragraph) return node.type !== nodes.locked_block

				const level = levelOf(node)

				if (level !== undefined) headings.push({ pos, level })

				return false
			})
		}

		let heading = 0
		const lastHeading = new Map<string, number>()

		for (const { pos, code } of fields) {
			const id = (code.args[0] ?? "").toLowerCase()

			while (heading < headings.length && (headings[heading]?.pos ?? 0) < pos) {
				const level = headings[heading]?.level ?? 9

				for (const [seq, reset] of resets) if (level <= reset - 1) lastHeading.set(seq, (lastHeading.get(seq) ?? 0) + 1)

				heading++
			}

			const resetAt = `${id}\u0000${String(lastHeading.get(id) ?? 0)}`
			let counter = counters.get(resetAt) ?? 0
			const restart = Number(switchOf(code, "r")?.arg)

			if (Number.isInteger(restart) && switchOf(code, "r") !== undefined) counter = restart
			else if (!hasSwitch(code, "c")) counter++

			counters.set(resetAt, counter)
			values.set(pos, counter)
		}

		this.sequences = values

		return values
	}
}

// ---- the paged editor as a host ----------------------------------------------------------------------------

// What an update reads from a paged editor: the pages shown (call once layout settled), positions of the body
// and of the story being edited located on them.
export function pagedFieldHost(editor: PagedEditor): FieldHost {
	let sections: Map<number, number> | undefined
	const geometryOf = (session: DocxEditorSession): PagedGeometry | undefined =>
		session === editor.body.session ? editor.body.geometry : session === editor.session ? editor.geometry : undefined

	return {
		body: editor.body,
		sessions: editor.sessions,
		pages: {
			get count() {
				return editor.pageCount
			},
			page: index => editor.page(index),
			pageAt: (session, pos) => geometryOf(session)?.locate(pos, 1)?.page,
			markerAt: (session, pos) => geometryOf(session)?.markerText(pos) ?? "",
			sectionPages(section) {
				if (sections === undefined) {
					const counts = new Map<number, number>()

					for (let index = 0; index < editor.pageCount; index++) {
						const page = editor.page(index)

						if (page !== undefined) counts.set(page.section, (counts.get(page.section) ?? 0) + 1)
					}

					sections = counts
				}

				return sections.get(section) ?? 0
			}
		}
	}
}
