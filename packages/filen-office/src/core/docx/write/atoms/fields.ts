import type { Node } from "prosemirror-model"
import { escapeText } from "../../../xml/index"
import { NO_SOURCE, docxSchema } from "../../edit/schema"
import { fieldUpdateOf, mergeFormat, parseFieldCode, rewritableSource, type SourceField } from "../../fields/evaluate"
import { storedResult, type FieldUpdate } from "../../fields/result"
import { Item, type StoryPart } from "../../model/index"
import { MARKER_ITEMS } from "../markers"
import { spliceChildren, textElement, type Names, type Out } from "../markup"
import { legacyFieldCopyData, legacyFieldEdited, writeLegacyFormField } from "../sdt"
import { RUN_PROPERTIES } from "../order.generated"
import { src } from "../source"
import type { StoryWriting } from "../story"
import { DocxWriteError } from "../values"
import type { AtomWriter } from "./index"

// Fields. A source field whose result an update changed keeps its bytes up to its separator (instruction,
// w:dirty, w:fldLock, revisions), gets the new result in runs formatted as \* MERGEFORMAT asks (the previous
// result's first run, else the field code's: ECMA-376 Part 1 §17.16.4.3), keeps the range markers and comment
// references its old result held, and ends with its own end character. An update made while tracking changes
// is written as Word records it: the result before in w:del, the new one in w:ins. A new field is complex (begin,
// instruction, separate, result, end) as Word writes them, or a w:fldSimple when it is simple. A field_char is
// one character of a field spanning paragraphs (a table of contents).

const nodes = docxSchema.nodes
const marks = docxSchema.marks

function deleted(node: Node): boolean {
	return marks.deletion.isInSet(node.marks) !== undefined || marks.move_from.isInSet(node.marks) !== undefined
}

// A result's markup inside a run: text, tabs and line breaks.
function resultContent(names: Names, text: string, removed: boolean): string {
	let out = ""

	for (const piece of text.split(/(\t|\n)/)) {
		if (piece === "\t") out += names.empty("tab")
		else if (piece === "\n") out += names.empty("br")
		else if (piece !== "") out += textElement(names, piece, removed)
	}

	return out
}

function instructionElement(names: Names, instruction: string, removed: boolean): string {
	const local = names.el(removed ? "delInstrText" : "instrText")

	return `<${local} xml:space="preserve">${escapeText(instruction)}</${local}>`
}

function fieldChar(names: Names, type: "begin" | "separate" | "end", dirty = false): string {
	return names.empty("fldChar", [
		["fldCharType", type],
		["dirty", dirty ? "true" : null]
	])
}

const DELETED_KINDS: ReadonlySet<number> = new Set([Item.DeletedText, Item.DeletedInstrText])

// Whether a tracked change inside `container` (a field's) deletes source item `item`.
function deletedWithin(part: StoryPart, item: number, container: number): boolean {
	for (let at = part.parent(item); at > container; at = part.parent(at))
		if (part.kinds[at] === Item.Deletion || part.kinds[at] === Item.MoveFrom) return true

	return false
}

// Source items `first` to `last` (siblings, of `container`) as their bytes up to byte `to`, their text and
// field codes as deleted text (w:delText, w:delInstrText) where the field is deleted (`removed`) or a tracked
// change inside it deletes them, else as text (a deletion of the field rejected).
function copyItems(w: StoryWriting, out: Out, first: number, last: number, to: number, container: number, removed: boolean): void {
	const part = w.part
	let at = part.start(first)

	for (let item = first; item < part.end(last) && part.start(item) < to; item++) {
		const kind = part.kinds[item] ?? 0
		const code = kind === Item.InstrText || kind === Item.DeletedInstrText

		if (!code && kind !== Item.Text && kind !== Item.DeletedText) continue

		const gone = removed || deletedWithin(part, item, container)

		if (gone === DELETED_KINDS.has(kind)) continue

		w.copySource(out, at, part.start(item), false)
		out.text(code ? instructionElement(w.names, part.text(item), gone) : textElement(w.names, part.text(item), gone))
		at = part.byteEnd(item)
	}

	w.copySource(out, at, to, false)
}

// A source run's properties without its tracked formatting change (whose id is the source run's).
function runProperties(w: StoryWriting, out: Out, run: number): void {
	const part = w.part
	const rPr = run < 0 || part.kinds[run] !== Item.Run ? -1 : (part.runs.rPr[part.aux[run] ?? -1] ?? -1)

	if (rPr >= 0) spliceChildren(out, part.reader, rPr, new Map([["rPrChange", ""]]), RUN_PROPERTIES)
}

// The run whose formatting a new result takes: the previous result's first run under \* MERGEFORMAT, else the
// run of the field code's first character.
function formattingRun(part: StoryPart, field: SourceField, node: Node): number {
	if (mergeFormat(parseFieldCode(node.attrs["instruction"] as string))) {
		const from = field.simple ? field.begin : field.separate
		const to = field.simple ? part.end(field.begin) : field.end

		if (from >= 0) for (let item = from + 1; item < to; item++) if (part.kinds[item] === Item.Run) return item
	}

	if (field.simple) return -1

	const instruction = part.fields.instruction[field.field] ?? -1

	return instruction >= 0 ? part.parent(instruction) : src(node)
}

// The range markers and comment references in source items [from, to), written again next to the result.
function keepMarkers(w: StoryWriting, out: Out, from: number, to: number): void {
	const part = w.part

	for (let item = from; item < to; item++) {
		const kind = part.kinds[item] ?? 0

		if (MARKER_ITEMS.has(kind)) {
			w.copyNode(out, part.node(item))
		} else if (kind === Item.CommentReference) {
			out.text(`<${w.names.el("r")}>`)
			w.copyNode(out, part.node(item))
			out.text(`</${w.names.el("r")}>`)
		}
	}
}

function resultRun(w: StoryWriting, out: Out, run: number, text: string, removed: boolean): void {
	if (text === "") return

	out.text(`<${w.names.el("r")}>`)
	runProperties(w, out, run)
	out.text(`${resultContent(w.names, text, removed)}</${w.names.el("r")}>`)
}

// A w:del or w:ins start tag for a field's tracked update (a new w:id: the save's own, as every revision's).
function revisionStart(w: StoryWriting, local: "del" | "ins", update: FieldUpdate): string {
	return w.names
		.empty(local, [
			["id", String(w.env.nextRevisionId())],
			["author", update.author],
			["date", update.date === "" ? undefined : update.date]
		])
		.replace(/\/>$/, ">")
}

// A field's result in runs formatted as `run`'s: with its tracked update, the result before deleted and the
// new one inserted. A field a tracked change deleted has its result as deleted text (an update inside it too:
// inserted by one change, deleted by another, as Word nests them).
function resultRuns(w: StoryWriting, out: Out, run: number, node: Node): void {
	const result = node.attrs["result"] as string
	const update = fieldUpdateOf(node)
	const removed = deleted(node)

	if (update === null) {
		resultRun(w, out, run, result, removed)

		return
	}

	if (update.result !== "") {
		out.text(revisionStart(w, "del", update))
		resultRun(w, out, run, update.result, true)
		out.text(`</${w.names.el("del")}>`)
	}

	if (result !== "") {
		out.text(revisionStart(w, "ins", update))
		resultRun(w, out, run, result, removed)
		out.text(`</${w.names.el("ins")}>`)
	}
}

function sameUpdate(a: FieldUpdate | null, b: FieldUpdate | null): boolean {
	return a === b || (a !== null && b !== null && a.result === b.result && a.author === b.author && a.date === b.date && a.id === b.id)
}

// A source field with a new result (see the top).
function rewrite(w: StoryWriting, out: Out, node: Node): void {
	const part = w.part
	const document = part.document
	const field = rewritableSource(part, node)

	if (field === undefined) throw new DocxWriteError("unsupported", "A field whose result cannot be written anew")

	const names = w.names
	const run = formattingRun(part, field, node)

	if (field.simple) {
		const xml = part.node(field.begin)
		const start = document.start(xml)

		// The start tag as an opening tag, its attributes (w:dirty, w:fldLock) as written.
		if (document.isSelfClosing(xml)) {
			let close = document.openEnd(xml) - 2

			while (close > start && /\s/.test(String.fromCharCode(document.bytes[close - 1] ?? 0))) close--

			w.copySource(out, start, close, false)
			out.text(">")
		} else {
			w.copySource(out, start, document.openEnd(xml), false)
		}

		keepMarkers(w, out, field.begin + 1, part.end(field.begin))
		resultRuns(w, out, run, node)
		out.text(`</${document.name(xml)}>`)

		return
	}

	const beginRun = src(node)
	const endRun = node.attrs["end"] as number

	if (field.separate >= 0) {
		// Up to the separator, then its run closed.
		copyItems(w, out, beginRun, endRun, part.byteEnd(field.separate), part.parent(beginRun), deleted(node))
		out.text(`</${document.name(part.node(part.parent(field.separate)))}>`)
		keepMarkers(w, out, field.separate + 1, field.end)
	} else {
		// No result yet: the code up to the end character, its run closed, and a separator.
		copyItems(w, out, beginRun, endRun, part.start(field.end), part.parent(beginRun), deleted(node))
		out.text(`</${document.name(part.node(endRun))}><${names.el("r")}>`)
		runProperties(w, out, endRun)
		out.text(`${fieldChar(names, "separate")}</${names.el("r")}>`)
	}

	resultRuns(w, out, run, node)
	out.text(`<${names.el("r")}>`)
	runProperties(w, out, endRun)
	w.copyNode(out, part.node(field.end))
	out.text(`</${names.el("r")}>`)
}

// A field the editor made (`ffData`: a copy of a legacy form field's, in its begin character).
function writeNew(w: StoryWriting, out: Out, node: Node, ffData: string): void {
	const names = w.names
	const instruction = node.attrs["instruction"] as string
	const result = node.attrs["result"] as string
	const removed = deleted(node)
	const run = (inner: string): string => `<${names.el("r")}>${inner}</${names.el("r")}>`
	const tracked = fieldUpdateOf(node) !== null

	if (node.attrs["simple"] === true) {
		const tag = names.empty("fldSimple", [["instr", instruction]])

		if (result === "" && !tracked) {
			out.text(tag)

			return
		}

		out.text(`${tag.slice(0, -2)}>`)
		resultRuns(w, out, NO_SOURCE, node)
		out.text(`</${names.el("fldSimple")}>`)

		return
	}

	const begin = fieldChar(names, "begin")

	out.text(
		run(ffData === "" ? begin : `${begin.slice(0, -2)}>${ffData}</${names.el("fldChar")}>`) +
			run(instructionElement(names, instruction, removed)) +
			run(fieldChar(names, "separate"))
	)
	resultRuns(w, out, NO_SOURCE, node)
	out.text(run(fieldChar(names, "end")))
}

const FIELD: AtomWriter = {
	type: nodes.field,
	// Decided in write: a source field whose result is still the stored one is written as its bytes there.
	edited: node => src(node) !== NO_SOURCE,
	write(w, out, node) {
		const own = src(node)

		if (own === NO_SOURCE) {
			writeNew(w, out, node, legacyFieldCopyData(w.names, node))

			return
		}

		// A legacy form field a form command filled in is the forms writer's.
		if (legacyFieldEdited(node)) {
			writeLegacyFormField(w, out, node)

			return
		}

		const part = w.part
		const end = node.attrs["end"] as number
		const stored = storedResult(part, own, end, part.kinds[own] === Item.FieldSimple)

		if ((node.attrs["result"] as string) === stored.result && sameUpdate(fieldUpdateOf(node), stored.update)) {
			const simple = part.kinds[own] === Item.FieldSimple

			copyItems(w, out, own, end, part.byteEnd(end), simple ? own : part.parent(own), deleted(node))

			return
		}

		rewrite(w, out, node)
	}
}

// A character of a field spanning paragraphs (begin, separate, end) or an instruction text of one
// ("instruction", "deletedInstruction" for a w:delInstrText of the file), inside the run opened for it. One of
// the file's is its bytes, an instruction text written anew when a tracked change deleted it or a deletion of
// it was rejected.
const FIELD_CHAR: AtomWriter = {
	type: nodes.field_char,
	// An instruction text deleted, or one the file has deleted no longer (a deletion rejected).
	edited: node =>
		(node.attrs["kind"] === "instruction" && deleted(node)) || (node.attrs["kind"] === "deletedInstruction" && !deleted(node)),
	write(w, out, node) {
		const kind = node.attrs["kind"] as string

		if (kind === "instruction" || kind === "deletedInstruction") {
			out.text(instructionElement(w.names, node.attrs["instruction"] as string, deleted(node)))

			return
		}

		if (kind !== "begin" && kind !== "separate" && kind !== "end")
			throw new DocxWriteError("unsupported", `A field character "${kind}"`)

		out.text(fieldChar(w.names, kind, kind === "begin" && node.attrs["dirty"] === true))
	}
}

export const fieldAtoms: readonly AtomWriter[] = [FIELD, FIELD_CHAR]
