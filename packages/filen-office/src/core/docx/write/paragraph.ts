import type { Mark, Node } from "prosemirror-model"
import { decodeUtf8 } from "../../xml/index"
import type { RunBase } from "../edit/formatting"
import { NO_SOURCE, docxSchema, inlineLevel } from "../edit/schema"
import type { BorderProps, ParagraphProps, TabStop } from "../props/index"
import { paragraphAttrs, paragraphChange, type ParagraphAttrs, type ParagraphChange, type RevisionStamp } from "../review/track"
import { Item, RevisionFlag } from "../model/index"
import { Ns } from "../model/namespaces"
import { sectionEditOf } from "../stories/sections"
import { atomEdited, writeAtom, writeRunAtom } from "./atoms/index"
import {
	closeTag,
	copyNode,
	freshContainer,
	openTag,
	openTagWithout,
	Out,
	spliceChildren,
	startTagWith,
	textElement,
	type Names
} from "./markup"
import { newParagraphTag, paragraphId, sameParagraphId, writeParagraphIdTag } from "./newStory"
import { JUSTIFICATION, NUMBERING_PROPERTIES, PARAGRAPH_MARK_PROPERTIES, PARAGRAPH_PROPERTIES } from "./order.generated"
import { baseWithoutChange, sameFormatting, writeRunProperties, type RunChange } from "./runs"
import { sectionEditMarkup } from "./sections"
import { sectionWritten, src } from "./source"
import { changedControl } from "./sdt"
import type { StoryWriting } from "./story"
import {
	DocxWriteError,
	indent as checkIndent,
	integer,
	name as checkName,
	oneOf,
	optionalOneOf,
	paragraphBorders as checkBorders,
	spacing as checkSpacing
} from "./values"
import { isRevisionWrapper, paragraphOf, sourceWrappers, transition, wrappersOf } from "./wrappers"

const nodes = docxSchema.nodes
const marks = docxSchema.marks

const MARK_REVISIONS = new Map([
	["ins", ""],
	["del", ""],
	["moveFrom", ""],
	["moveTo", ""],
	["rPrChange", ""]
])

// A run of a paragraph's inline content written as one unit: consecutive content of one source run (or of
// no run) under the same wrappers, or one paragraph-level atom.
interface Unit {
	readonly nodes: Node[]
	readonly run: Mark | undefined
	readonly wrappers: readonly Mark[]
	readonly atom: boolean
	copy: boolean
}

export function writeParagraph(w: StoryWriting, out: Out, node: Node): void {
	const part = w.part
	const document = part.document
	const own = src(node)
	const source = own !== NO_SOURCE && part.kinds[own] === Item.Paragraph ? own : NO_SOURCE
	const owner = source !== NO_SOURCE && w.owns(node)

	const from = out.size
	const id = paragraphId(node)

	for (const mark of node.marks)
		if (
			mark.type !== marks.para_id &&
			!(mark.type === marks.paragraph_change && writableChange(mark.attrs["value"])) &&
			!(mark.type === marks.section_change && sectionEditOf(node) !== undefined)
		)
			throw new DocxWriteError("missingSource", `A ${mark.type.name} mark on a ${node.type.name}`)

	// Word's w14:textId hashes the paragraph's text: an edited paragraph drops it (Word writes a new one);
	// the paragraph keeps its w14:paraId unless a feature gave it a new one.
	if (owner && id !== undefined && !sameParagraphId(w, source, id)) writeParagraphIdTag(w, out, source, id)
	else if (owner && w.force === undefined && w.textId !== undefined) openTagWithout(out, document, part.node(source), w.textId)
	else if (owner) openTag(out, document, part.node(source))
	else out.text(newParagraphTag(w, source, id))

	writeParagraphProperties(w, out, node, source, owner)
	writeInline(w, out, node, owner ? source : NO_SOURCE)

	if (owner) closeTag(out, document, part.node(source))
	else out.text(`</${w.names.el("p")}>`)

	w.paragraphs.push({ node, from, to: out.size })
}

const CHANGE_KEYS: ReadonlySet<string> = new Set(["props", "inserted", "deleted", "source"])

// Whether a paragraph_change value is what tracked changes record (see ParagraphChange): anything else on it
// has no writer.
function writableChange(value: unknown): boolean {
	if (typeof value !== "object" || value === null) return false

	const change = value as ParagraphChange

	return (
		Object.keys(value).every(key => CHANGE_KEYS.has(key)) &&
		(change.source === undefined || change.source === "accept" || change.source === "reject") &&
		(change.props === undefined || (typeof change.props.author === "string" && typeof change.props.before === "object"))
	)
}

function markup(write: (out: Out) => void): string {
	const out = new Out()

	write(out)

	return decodeUtf8(out.finish())
}

// The pPr children the paragraph's attributes stand for (style, numbering, alignment, indent), where `attrs`
// differ from `props` (the properties of the container they are spliced into).
function attributeChanges(
	w: StoryWriting,
	attrs: ParagraphAttrs,
	props: ParagraphProps,
	owner: boolean,
	sourceNumPr: number
): Map<string, string> {
	const names = w.names
	const reader = w.part.reader
	const changes = new Map<string, string>()
	const { styleId, numId, ilvl, align, indent, spacing, borders } = attrs

	if (styleId !== (props.styleId ?? null)) {
		if (styleId !== null) checkName("w:pStyle", styleId)

		changes.set("pStyle", styleId === null ? "" : names.val("pStyle", styleId))
	}

	const numberingChanged = numId !== (props.numbering?.numId ?? null) || ilvl !== (props.numbering?.ilvl ?? null)

	if (numberingChanged || (!owner && sourceNumPr >= 0)) {
		const fileNumId = numId === null ? null : numId < 0 ? w.env.numId(numId) : numId

		if (fileNumId !== null) integer("w:numId", fileNumId, 0)
		if (ilvl !== null) integer("w:ilvl", ilvl, 0, 8)
		const numberingChanges = new Map<string, string>([
			["ilvl", ilvl === null ? "" : names.val("ilvl", String(ilvl))],
			["numId", fileNumId === null ? "" : names.val("numId", String(fileNumId))]
		])

		if (!owner) {
			numberingChanges.set("numberingChange", "")
			numberingChanges.set("ins", "")
		}

		if (numId === null && ilvl === null) changes.set("numPr", "")
		else if (sourceNumPr >= 0)
			changes.set(
				"numPr",
				markup(sub => spliceChildren(sub, reader, sourceNumPr, numberingChanges, NUMBERING_PROPERTIES))
			)
		else changes.set("numPr", freshContainer(names, "numPr", numberingChanges, NUMBERING_PROPERTIES))
	}

	if (align !== (props.alignment ?? null)) {
		if (align !== null) oneOf("w:jc", JUSTIFICATION, align)

		changes.set("jc", align === null ? "" : names.val("jc", align))
	}

	if (spacing !== (props.spacing ?? null)) {
		checkSpacing(spacing)
		changes.set(
			"spacing",
			spacing === null
				? ""
				: names.empty("spacing", [
						["before", spacing.before?.toString()],
						["beforeLines", spacing.beforeLines?.toString()],
						["beforeAutospacing", onOff(spacing.beforeAutospacing)],
						["after", spacing.after?.toString()],
						["afterLines", spacing.afterLines?.toString()],
						["afterAutospacing", onOff(spacing.afterAutospacing)],
						["line", spacing.line?.toString()],
						["lineRule", spacing.lineRule]
					])
		)
	}

	if (borders !== (props.borders ?? null)) {
		checkBorders(borders)

		const sides = BORDER_ELEMENTS.flatMap(([side, local]) => {
			const border = borders?.[side]

			return border === undefined ? [] : [borderElement(names, local, border)]
		})

		changes.set("pBdr", sides.length === 0 ? "" : `<${names.el("pBdr")}>${sides.join("")}</${names.el("pBdr")}>`)
	}

	if (indent !== (props.indent ?? null)) {
		checkIndent(indent)
		changes.set(
			"ind",
			indent === null
				? ""
				: names.empty("ind", [
						["left", indent.start?.toString()],
						["leftChars", indent.startChars?.toString()],
						["right", indent.end?.toString()],
						["rightChars", indent.endChars?.toString()],
						["hanging", indent.hanging?.toString()],
						["hangingChars", indent.hangingChars?.toString()],
						["firstLine", indent.firstLine?.toString()],
						["firstLineChars", indent.firstLineChars?.toString()]
					])
		)
	}

	return changes
}

// w:tabs from the editor's tab stops (ST_TabJc, ST_TabTlc, twips), checked.
function tabsElement(names: Names, tabs: readonly TabStop[]): string {
	const out = tabs.map(tab => {
		oneOf("w:tab/@w:val", TAB_ALIGNMENTS, tab.alignment)
		optionalOneOf("w:tab/@w:leader", TAB_LEADERS, tab.leader)
		integer("w:tab/@w:pos", tab.position, -31680, 31680)

		return names.empty("tab", [
			["val", tab.alignment],
			["leader", tab.leader],
			["pos", String(tab.position)]
		])
	})

	return out.length === 0 ? "" : `<${names.el("tabs")}>${out.join("")}</${names.el("tabs")}>`
}

const TAB_ALIGNMENTS: ReadonlySet<string> = new Set(["clear", "start", "left", "center", "end", "right", "decimal", "bar", "num"])
const TAB_LEADERS: ReadonlySet<string> = new Set(["none", "dot", "hyphen", "underscore", "heavy", "middleDot"])

function onOff(value: boolean | undefined): string | undefined {
	return value === undefined ? undefined : value ? "1" : "0"
}

// CT_PBdr's sides in schema order, with their Transitional element names (start/end as left/right, like w:ind).
const BORDER_ELEMENTS = [
	["top", "top"],
	["start", "left"],
	["bottom", "bottom"],
	["end", "right"],
	["between", "between"],
	["bar", "bar"]
] as const

function hex2(value: number | undefined): string | undefined {
	return value === undefined ? undefined : value.toString(16).toUpperCase().padStart(2, "0")
}

function borderElement(names: Names, local: string, border: BorderProps): string {
	return names.empty(local, [
		["val", border.style],
		["sz", border.size?.toString()],
		["space", border.space?.toString()],
		["color", border.val ?? (border.themeColor === undefined ? undefined : "auto")],
		["themeColor", border.themeColor],
		["themeTint", hex2(border.themeTint)],
		["themeShade", hex2(border.themeShade)],
		["shadow", onOff(border.shadow)],
		["frame", onOff(border.frame)]
	])
}

// A revision element (w:ins, w:del, w:pPrChange...) the editor recorded: a new w:id, its author and date.
export function revisionTag(w: StoryWriting, local: string, stamp: RevisionStamp, close: boolean): string {
	const names = w.names
	const attributes = names.empty(local, [
		["id", String(w.env.nextRevisionId())],
		["author", stamp.author],
		["date", stamp.date === "" ? undefined : stamp.date]
	])

	return close ? attributes : attributes.replace(/\/>$/, ">")
}

// The child of a WordprocessingML container named `local`, or -1.
function childNamed(w: StoryWriting, container: number, local: string): number {
	const document = w.part.document

	if (container < 0) return -1

	for (let child = document.firstChildElement(container); child >= 0; child = document.nextSiblingElement(child)) {
		if (w.part.reader.group(child) === Ns.W && document.localName(child) === local) return child
	}

	return -1
}

// A source element with a new w:id (written where the paragraph owning it is not).
function withNewId(w: StoryWriting, element: number): string {
	const document = w.part.document

	return markup(sub => {
		startTagWith(sub, document, element, w.names.attr("id"), String(w.env.nextRevisionId()), true)
		sub.bytes(document.bytes, document.openEnd(element), document.end(element))
	})
}

const MARK_KINDS: readonly (readonly [string, number])[] = [
	["ins", RevisionFlag.Inserted],
	["del", RevisionFlag.Deleted],
	["moveFrom", RevisionFlag.MovedFrom],
	["moveTo", RevisionFlag.MovedTo]
]

// The paragraph mark's run properties where its revisions differ from the source's, as markup; undefined when
// the source's stand. What is written follows the paragraph's flags, whichever occurrence of a repeated source
// paragraph owns its bytes: a flag with a stamp is a new element; one without, the source's element (a new id
// where this paragraph does not own it); a source element without the flag goes.
function markProperties(w: StoryWriting, node: Node, change: ParagraphChange | undefined, row: number, owner: boolean): string | undefined {
	const part = w.part
	const markRPr = row < 0 ? -1 : (part.paragraphs.markRPr[row] ?? -1)
	const source = row >= 0 ? (part.paragraphs.markRevision[row] ?? 0) : 0
	const flags = node.attrs["revision"] as number
	const changes = new Map<string, string>()

	if (!owner) for (const local of MARK_REVISIONS.keys()) changes.set(local, "")

	for (const [local, flag] of MARK_KINDS) {
		const stamp = flag === RevisionFlag.Inserted ? change?.inserted : flag === RevisionFlag.Deleted ? change?.deleted : undefined
		const wanted = (flags & flag) !== 0
		const had = (source & flag) !== 0

		if (stamp !== undefined && wanted) {
			// The source's own element moved here with the mark, with a new id: a copy of the paragraph (a paste)
			// may hold the same element where it was.
			const element = stamp.element ?? -1

			changes.set(local, element >= 0 ? withNewId(w, element) : revisionTag(w, local, stamp, true))
		} else if (had && !wanted) {
			changes.set(local, "")
		} else if (had && wanted && !owner) {
			const element = childNamed(w, markRPr, local)

			if (element >= 0) changes.set(local, withNewId(w, element))
		}
	}

	if (changes.size === 0 || (owner && [...changes.values()].every(value => value === "") && markRPr < 0)) return undefined
	if (markRPr >= 0) {
		const spliced = markup(sub => spliceChildren(sub, part.reader, markRPr, changes, PARAGRAPH_MARK_PROPERTIES))

		// A mark whose last property went (a revision resolved) has no w:rPr left to write.
		return /^<[^\s>]+>\s*<\/[^>]+>$/.test(spliced) ? "" : spliced
	}

	const fresh = freshContainer(w.names, "rPr", changes, PARAGRAPH_MARK_PROPERTIES)

	return fresh === "" ? undefined : fresh
}

function writeParagraphProperties(w: StoryWriting, out: Out, node: Node, source: number, owner: boolean): void {
	const names = w.names
	const part = w.part
	const reader = part.reader
	const document = part.document
	const row = source === NO_SOURCE ? -1 : (part.aux[source] ?? -1)
	const pPr = row < 0 ? -1 : (part.paragraphs.pPr[row] ?? -1)
	const sectPr = node.attrs["sectPr"] as number
	const sourceNumPr = row < 0 ? -1 : (part.paragraphs.numPr[row] ?? -1)
	const change = paragraphChange(node)
	// The source's own w:pPrChange: kept where the paragraph is (a new id where it does not own the bytes)
	// unless accepted or rejected since; a copy pasted in had it taken back (see review/track.ts).
	const sourceChange = pPr >= 0 ? childNamed(w, pPr, "pPrChange") : -1
	// A rejected source change writes the properties it recorded, under the paragraph's mark and section break.
	const rejected = sourceChange >= 0 && change?.source === "reject" ? childNamed(w, sourceChange, "pPr") : -1
	const container = change?.source === "reject" && sourceChange >= 0 ? rejected : pPr
	const props = w.env.context.store.paragraph(document, container)
	const changes = attributeChanges(w, paragraphAttrs(node), props, owner, sourceNumPr)
	// Tab stops a command set (a table of contents' right tab; absent from schemas without the attribute).
	const tabs = node.attrs["tabs"] as readonly TabStop[] | null | undefined

	if (tabs !== null && tabs !== undefined) changes.set("tabs", tabsElement(names, tabs))

	const hasSection = row >= 0 && (part.paragraphs.sectPr[row] ?? -1) >= 0
	const sectionEdit = sectionEditOf(node)

	if (!sectionWritten(part, sectPr, w.cells > 0)) {
		if (hasSection) changes.set("sectPr", "")
	} else if (sectionEdit !== undefined) {
		changes.set("sectPr", sectionEditMarkup(w, sectionEdit))
	} else if (sectPr !== source || !hasSection || container !== pPr) {
		const from = part.paragraphs.sectPr[part.aux[sectPr] ?? -1] ?? -1

		if (from < 0) throw new DocxWriteError("missingSource", "A section break without its properties")

		changes.set("sectPr", decodeUtf8(w.bytes, document.start(from), document.end(from)))
	}

	if (!owner && pPr >= 0) changes.set("pPrChange", sourceChange >= 0 && change?.source === undefined ? withNewId(w, sourceChange) : "")

	const mark = markProperties(w, node, change, row, owner)

	if (mark !== undefined) {
		changes.set("rPr", mark)
	} else if (container !== pPr && row >= 0) {
		const markRPr = part.paragraphs.markRPr[row] ?? -1

		if (markRPr >= 0) changes.set("rPr", decodeUtf8(w.bytes, document.start(markRPr), document.end(markRPr)))
	}

	if (sourceChange >= 0 && change?.source !== undefined) changes.set("pPrChange", "")

	if (change?.props !== undefined && !(sourceChange >= 0 && change.source === undefined)) {
		// What the editor recorded: the properties before, as this paragraph's pPr would have been.
		const before = attributeChanges(w, change.props.before, props, owner, sourceNumPr)

		for (const local of ["rPr", "sectPr", "pPrChange"]) before.set(local, "")

		const inner =
			container >= 0
				? markup(sub => spliceChildren(sub, reader, container, before, PARAGRAPH_PROPERTIES))
				: freshContainer(names, "pPr", before, PARAGRAPH_PROPERTIES) || names.empty("pPr")

		changes.set("pPrChange", `${revisionTag(w, "pPrChange", change.props, false)}${inner}</${names.el("pPrChange")}>`)
	}

	if (container >= 0) {
		if (changes.size === 0) {
			copyNode(out, document, container)

			return
		}

		// Properties that lost their last child are left out rather than written empty.
		const spliced = markup(sub => spliceChildren(sub, reader, container, changes, PARAGRAPH_PROPERTIES))

		if (!/^<[^>]*>\s*<\/[^>]*>$/.test(spliced)) out.text(spliced)
	} else {
		out.text(freshContainer(names, "pPr", changes, PARAGRAPH_PROPERTIES))
	}
}

// ---- inline content ----------------------------------------------------------------------------------------

function units(paragraph: Node): Unit[] {
	const units: Unit[] = []

	paragraph.forEach(child => {
		const run = marks.run.isInSet(child.marks)
		const wrappers = wrappersOf(child)
		const inRun = inlineLevel(child) === "run" || (child.type === nodes.marker && run !== undefined)
		const last = units[units.length - 1]

		if (inRun && last !== undefined && !last.atom && sameMark(last.run, run) && sameMarks(last.wrappers, wrappers)) {
			last.nodes.push(child)
		} else {
			units.push({ nodes: [child], run, wrappers, atom: !inRun, copy: false })
		}
	})

	return units
}

// Content tokens of a source run (text and atoms by item), as its unit would show them.
function sourceTokens(w: StoryWriting, run: number): string {
	const part = w.part
	let out = ""

	for (let item = run + 1; item < part.end(run); item = part.end(item)) {
		const kind = part.kinds[item]

		if (kind === Item.Text || kind === Item.DeletedText) out += part.text(item)
		else out += `\u0000${item}\u0000`
	}

	return out
}

function unitTokens(unit: Unit): string {
	let out = ""

	for (const node of unit.nodes) out += node.isText ? (node.text ?? "") : `\u0000${src(node)}\u0000`

	return out
}

function base(w: StoryWriting, run: Mark | undefined): RunBase | undefined {
	const id = run?.attrs["base"] as number | undefined

	return id === undefined || id < 0 ? undefined : w.env.context.bases.get(id)
}

// What becomes of a node's tracked formatting change (see RunChange). `owned`: the source paragraph this
// paragraph owns (its runs' changes keep their ids there), or NO_SOURCE.
function runChange(w: StoryWriting, node: Node, run: number, runBase: RunBase | undefined, owned: number): RunChange {
	const change = marks.format_change.isInSet(node.marks)
	const recorded = change?.attrs["base"] as number | null | undefined
	const sourceChange = runBase?.props.change !== undefined

	if (change !== undefined && recorded !== null && recorded !== undefined) {
		const before = baseWithoutChange(w.names, recorded < 0 ? undefined : w.env.context.bases.get(recorded))
		const stamp = { author: String(change.attrs["author"] ?? ""), date: String(change.attrs["date"] ?? "") }

		return { kind: "new", markup: `${revisionTag(w, "rPrChange", stamp, false)}${before}</${w.names.el("rPrChange")}>` }
	}

	if (change === undefined) return sourceChange ? { kind: "strip" } : { kind: "keep" }

	// A formatting change read from the file lives in the run's base.
	if (!sourceChange) throw new DocxWriteError("missingSource", "A formatting change without its source")

	// Written a second time, or away from its own paragraph (a copy): a new id.
	if ((runBase !== undefined && w.writtenChanges.has(runBase.id)) || run === NO_SOURCE || paragraphOf(w, run) !== owned)
		return { kind: "keep", id: w.env.nextRevisionId() }

	return { kind: "keep" }
}

function canCopyUnit(w: StoryWriting, unit: Unit, forced: boolean, owned: number): boolean {
	if (forced || unit.atom || unit.run === undefined) return false

	const run = src(unit.run)

	if (run === NO_SOURCE || w.copiedRuns.has(run) || w.part.kinds[run] !== Item.Run) return false

	const runBase = base(w, unit.run)
	const sourceChange = runBase?.props.change !== undefined
	// Away from its own paragraph (a copy, or a joined paragraph's), a run whose bytes hold revision ids is
	// rebuilt with new ones.
	const away = paragraphOf(w, run) !== owned

	if (away && (sourceChange || unit.wrappers.some(isRevisionWrapper))) return false

	for (const node of unit.nodes) {
		if (node.type !== nodes.marker && !sameFormatting(node.marks, runBase)) return false
		if (node.isText ? newFormatChange(node.marks) : atomEdited(node)) return false
		if (node.type !== nodes.marker && (marks.format_change.isInSet(node.marks) !== undefined) !== sourceChange) return false
	}

	const wrappers = sourceWrappers(w, run)

	if (wrappers.length !== unit.wrappers.length || unit.wrappers.some((mark, k) => src(mark) !== wrappers[k] || changedControl(mark)))
		return false

	return unitTokens(unit) === sourceTokens(w, run)
}

function writeInline(w: StoryWriting, out: Out, paragraph: Node, owned: number): void {
	const forced = paragraph === w.force
	const list = units(paragraph)

	for (const unit of list) {
		unit.copy = canCopyUnit(w, unit, forced, owned)

		if (unit.copy && unit.run !== undefined) {
			w.copiedRuns.add(src(unit.run))

			const runBase = base(w, unit.run)

			if (runBase?.props.change !== undefined) w.writtenChanges.add(runBase.id)
		}
	}

	const stack: Mark[] = []
	// Rendered page breaks are Word's layout cache: rebuilt runs leave them out (and an otherwise empty
	// wrapper with them).
	const dropped = (unit: Unit): boolean => !forced && !unit.copy && unit.nodes.every(child => child.type === nodes.rendered_break)
	// A proofing mark is no item of the model: a hyperlink holding nothing else (its text deleted) would read
	// back as an empty hyperlink, kept as its bytes. Those marks go, and the hyperlink with them.
	const holdsContent = (link: Mark): boolean =>
		list.some(unit => unit.nodes.length > 0 && unit.nodes[0]?.type !== nodes.proof_err && !dropped(unit) && link.isInSet(unit.wrappers))

	list.forEach((unit, index) => {
		const node = unit.nodes[0]

		if (node === undefined || dropped(unit)) return

		if (node.type === nodes.proof_err && unit.wrappers.some(mark => mark.type === marks.link && !holdsContent(mark))) return

		// Bookkeeping marks stay only between content written as it was.
		if (node.type === nodes.proof_err) {
			const previous = list[index - 1]
			const next = list[index + 1]

			// Gate 1 keeps them: it proves the run rebuilding, and with them most paragraphs come out as
			// their very bytes.
			if (
				!forced &&
				((previous !== undefined && !previous.atom && !previous.copy) || (next !== undefined && !next.atom && !next.copy))
			) {
				return
			}
		}

		transition(w, out, stack, unit.wrappers, owned)

		if (unit.atom) {
			writeAtom(w, out, node)

			return
		}

		if (unit.copy && unit.run !== undefined) {
			w.copyNode(out, w.part.node(src(unit.run)))

			return
		}

		writeRuns(w, out, unit, forced, owned)
	})

	transition(w, out, stack, [], owned)
}

// A unit rebuilt: consecutive nodes with the same marks become one w:r with the base's properties and
// their changes; markers Word had put inside the run go between the runs.
function writeRuns(w: StoryWriting, out: Out, unit: Unit, forced: boolean, owned: number): void {
	const names = w.names
	const part = w.part
	const document = part.document
	const run = unit.run === undefined ? NO_SOURCE : src(unit.run)
	const runBase = base(w, unit.run)
	let open: readonly Mark[] | null = null
	let raw = false

	const close = (): void => {
		if (open === null) return

		out.text(raw ? `</${document.name(part.node(run))}>` : `</${names.el("r")}>`)
		open = null
	}

	for (const node of unit.nodes) {
		if (node.type === nodes.marker) {
			close()
			writeAtom(w, out, node)

			continue
		}

		// Word's record of its last page break goes stale with the edit (gate 1 keeps it, see above).
		if (node.type === nodes.rendered_break && !forced) continue

		if (open === null || !sameMarks(open, node.marks)) {
			close()
			raw = run !== NO_SOURCE && part.kinds[run] === Item.Run && !w.startedRuns.has(run)

			if (raw) {
				w.startedRuns.add(run)
				openTag(out, document, part.node(run))
			} else {
				out.text(`<${names.el("r")}>`)
			}

			const change = runChange(w, node, run, runBase, owned)

			writeRunProperties(out, names, node.marks, runBase, change)

			if (change.kind === "keep" && runBase?.props.change !== undefined) w.writtenChanges.add(runBase.id)
			open = node.marks
		}

		if (node.isText) {
			// Deleted text is w:delText; moved-away text stays w:t, as Word writes it.
			out.text(textElement(names, node.text ?? "", marks.deletion.isInSet(node.marks) !== undefined))
		} else {
			writeRunAtom(w, out, node)
		}
	}

	close()
}

// A tracked formatting change a command recorded (its `base`), not one read from the source.
function newFormatChange(nodeMarks: readonly Mark[]): boolean {
	return (marks.format_change.isInSet(nodeMarks)?.attrs["base"] ?? null) !== null
}

function sameMark(a: Mark | undefined, b: Mark | undefined): boolean {
	return a === undefined ? b === undefined : b !== undefined && a.eq(b)
}

function sameMarks(a: readonly Mark[], b: readonly Mark[]): boolean {
	if (a === b) return true
	if (a.length !== b.length) return false

	for (let k = 0; k < a.length; k++) if (!(a[k] as Mark).eq(b[k] as Mark)) return false

	return true
}
