import type { Mark, Node } from "prosemirror-model"
import { decodeUtf8 } from "../../xml/index"
import type { RunBase } from "../edit/formatting"
import { NO_SOURCE, docxSchema, inlineLevel, type IndentAttrs } from "../edit/schema"
import { Item } from "../model/index"
import { atomEdited, writeAtom, writeRunAtom } from "./atoms/index"
import { closeTag, copyNode, freshContainer, openTag, openTagWithout, Out, spliceChildren, textElement } from "./markup"
import { newParagraphTag, paragraphId, sameParagraphId, writeParagraphIdTag } from "./newStory"
import { JUSTIFICATION, NUMBERING_PROPERTIES, PARAGRAPH_MARK_PROPERTIES, PARAGRAPH_PROPERTIES } from "./order.generated"
import { sameFormatting, writeRunProperties } from "./runs"
import { refuseBlockMarks, sectionWritten, src } from "./source"
import { changedControl } from "./sdt"
import type { StoryWriting } from "./story"
import { DocxWriteError, indent as checkIndent, integer, name as checkName, oneOf } from "./values"
import { sourceWrappers, transition, wrappersOf } from "./wrappers"

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

	refuseBlockMarks(node, marks.para_id)

	// Word's w14:textId hashes the paragraph's text: an edited paragraph drops it (Word writes a new one);
	// the paragraph keeps its w14:paraId unless a feature gave it a new one.
	if (owner && id !== undefined && !sameParagraphId(w, source, id)) writeParagraphIdTag(w, out, source, id)
	else if (owner && w.force === undefined && w.textId !== undefined) openTagWithout(out, document, part.node(source), w.textId)
	else if (owner) openTag(out, document, part.node(source))
	else out.text(newParagraphTag(w, source, id))

	writeParagraphProperties(w, out, node, source, owner)
	writeInline(w, out, node)

	if (owner) closeTag(out, document, part.node(source))
	else out.text(`</${w.names.el("p")}>`)

	w.paragraphs.push({ node, from, to: out.size })
}

function markup(write: (out: Out) => void): string {
	const out = new Out()

	write(out)

	return decodeUtf8(out.finish())
}

function writeParagraphProperties(w: StoryWriting, out: Out, node: Node, source: number, owner: boolean): void {
	const names = w.names
	const part = w.part
	const reader = part.reader
	const document = part.document
	const row = source === NO_SOURCE ? -1 : (part.aux[source] ?? -1)
	const pPr = row < 0 ? -1 : (part.paragraphs.pPr[row] ?? -1)
	const props = w.env.context.store.paragraph(document, pPr)
	const changes = new Map<string, string>()
	const styleId = node.attrs["styleId"] as string | null
	const numId = node.attrs["numId"] as number | null
	const ilvl = node.attrs["ilvl"] as number | null
	const align = node.attrs["align"] as string | null
	const indent = node.attrs["indent"] as IndentAttrs | null
	const sectPr = node.attrs["sectPr"] as number
	const sourceNumPr = row < 0 ? -1 : (part.paragraphs.numPr[row] ?? -1)

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

	const hasSection = row >= 0 && (part.paragraphs.sectPr[row] ?? -1) >= 0

	if (!sectionWritten(part, sectPr, w.cells > 0)) {
		if (hasSection) changes.set("sectPr", "")
	} else if (sectPr !== source || !hasSection) {
		const from = part.paragraphs.sectPr[part.aux[sectPr] ?? -1] ?? -1

		if (from < 0) throw new DocxWriteError("missingSource", "A section break without its properties")

		changes.set("sectPr", decodeUtf8(w.bytes, document.start(from), document.end(from)))
	}

	if (!owner && pPr >= 0) {
		// A copy of the paragraph's properties is a new paragraph mark: tracked changes stay with the original.
		changes.set("pPrChange", "")

		const markRPr = part.paragraphs.markRPr[row] ?? -1

		if (markRPr >= 0)
			changes.set(
				"rPr",
				markup(sub => spliceChildren(sub, reader, markRPr, MARK_REVISIONS, PARAGRAPH_MARK_PROPERTIES))
			)
	}

	if (pPr >= 0) {
		if (changes.size === 0) {
			copyNode(out, document, pPr)

			return
		}

		// Properties that lost their last child are left out rather than written empty.
		const spliced = markup(sub => spliceChildren(sub, reader, pPr, changes, PARAGRAPH_PROPERTIES))

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

function canCopyUnit(w: StoryWriting, unit: Unit, forced: boolean): boolean {
	if (forced || unit.atom || unit.run === undefined) return false

	const run = src(unit.run)

	if (run === NO_SOURCE || w.copiedRuns.has(run) || w.part.kinds[run] !== Item.Run) return false

	const runBase = base(w, unit.run)

	for (const node of unit.nodes) {
		if (node.type !== nodes.marker && !sameFormatting(node.marks, runBase)) return false
		if (node.isText ? newFormatChange(node.marks) : atomEdited(node)) return false
	}

	const wrappers = sourceWrappers(w, run)

	if (wrappers.length !== unit.wrappers.length || unit.wrappers.some((mark, k) => src(mark) !== wrappers[k] || changedControl(mark)))
		return false

	return unitTokens(unit) === sourceTokens(w, run)
}

function writeInline(w: StoryWriting, out: Out, paragraph: Node): void {
	const forced = paragraph === w.force
	const list = units(paragraph)

	for (const unit of list) {
		unit.copy = canCopyUnit(w, unit, forced)

		if (unit.copy && unit.run !== undefined) {
			w.copiedRuns.add(src(unit.run))

			const runBase = base(w, unit.run)

			if (runBase?.props.change !== undefined) w.writtenChanges.add(runBase.id)
		}
	}

	const stack: Mark[] = []

	list.forEach((unit, index) => {
		const node = unit.nodes[0]

		if (node === undefined) return

		// Rendered page breaks are Word's layout cache: rebuilt runs leave them out (and an otherwise
		// empty wrapper with them).
		if (!forced && !unit.copy && unit.nodes.every(child => child.type === nodes.rendered_break)) return

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

		transition(w, out, stack, unit.wrappers)

		if (unit.atom) {
			writeAtom(w, out, node)

			return
		}

		if (unit.copy && unit.run !== undefined) {
			w.copyNode(out, w.part.node(src(unit.run)))

			return
		}

		writeRuns(w, out, unit, forced)
	})

	transition(w, out, stack, [])
}

// A unit rebuilt: consecutive nodes with the same marks become one w:r with the base's properties and
// their changes; markers Word had put inside the run go between the runs.
function writeRuns(w: StoryWriting, out: Out, unit: Unit, forced: boolean): void {
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
			// A formatting change recorded by a command is not written yet.
			if (newFormatChange(node.marks)) throw new DocxWriteError("missingSource", "A new formatting change")

			close()
			raw = run !== NO_SOURCE && part.kinds[run] === Item.Run && !w.startedRuns.has(run)

			if (raw) {
				w.startedRuns.add(run)
				openTag(out, document, part.node(run))
			} else {
				out.text(`<${names.el("r")}>`)
			}

			// A tracked formatting change written a second time needs an id of its own.
			const changed = runBase?.props.change !== undefined

			writeRunProperties(
				out,
				names,
				node.marks,
				runBase,
				changed && w.writtenChanges.has(runBase.id) ? w.env.nextRevisionId() : undefined
			)

			if (changed) w.writtenChanges.add(runBase.id)
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
