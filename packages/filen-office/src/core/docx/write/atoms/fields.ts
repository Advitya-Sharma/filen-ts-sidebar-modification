import type { Node } from "prosemirror-model"
import { escapeText } from "../../../xml/index"
import { NO_SOURCE, docxSchema } from "../../edit/schema"
import { mergeFormat, parseFieldCode, rewritableSource, type SourceField } from "../../fields/evaluate"
import { FieldCharType, Item, type StoryPart } from "../../model/index"
import { MARKER_ITEMS } from "../markers"
import { spliceChildren, textElement, type Names, type Out } from "../markup"
import { legacyFieldEdited, writeLegacyFormField } from "../sdt"
import { RUN_PROPERTIES } from "../order.generated"
import { src } from "../source"
import type { StoryWriting } from "../story"
import { DocxWriteError } from "../values"
import type { AtomWriter } from "./index"

// Fields. A source field whose result an update changed keeps its bytes up to its separator (instruction,
// w:dirty, w:fldLock, revisions), gets the new result in runs formatted as \* MERGEFORMAT asks (the previous
// result's first run, else the field code's: ECMA-376 Part 1 §17.16.4.3), keeps the range markers and comment
// references its old result held, and ends with its own end character. A new field is complex (begin,
// instruction, separate, result, end) as Word writes them, or a w:fldSimple when it is simple. A field_char is
// one character of a field spanning paragraphs (a table of contents the editor made).

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

// The cached result of a source field as the editor reads it (what follows the separator, nested field codes
// left out): the document builder's reading, so a result no update changed is told apart.
function storedResult(part: StoryPart, first: number, last: number, simple: boolean): string {
	const field = simple ? (part.aux[first] ?? -1) : part.fieldOf(first + 1)
	const separate = simple || field < 0 ? first : (part.fields.separate[field] ?? -1)

	if (separate < 0) return ""

	const end = part.end(last)
	const codes: boolean[] = []
	let out = ""
	let depth = 0
	let code = false

	for (let item = separate + 1; item < end; item++) {
		switch (part.kinds[item]) {
			case Item.FieldChar: {
				const type = part.fieldCharType(item)

				if (type === FieldCharType.Begin) {
					codes.push(code)
					depth++
					code = true
				} else if (type === FieldCharType.Separate && depth > 0) {
					code = codes[codes.length - 1] ?? false
				} else if (type === FieldCharType.End && depth > 0) {
					depth--
					code = codes.pop() ?? false
				}

				break
			}
			case Item.Text:
			case Item.DeletedText:
				if (!code) out += part.text(item)

				break
			case Item.Tab:
				if (!code) out += "\t"

				break
			case Item.Break:
			case Item.CarriageReturn:
				if (!code) out += "\n"

				break
		}
	}

	return out
}

// Source items `first` to `last` (siblings) as their bytes, their text and field codes as deleted text
// (w:delText, w:delInstrText): a field a tracked change deleted.
function copyDeleted(w: StoryWriting, out: Out, first: number, last: number): void {
	const part = w.part
	let at = part.start(first)

	for (let item = first; item < part.end(last); item++) {
		const kind = part.kinds[item]

		if (kind !== Item.Text && kind !== Item.InstrText) continue

		w.copySource(out, at, part.start(item), false)
		out.text(kind === Item.Text ? textElement(w.names, part.text(item), true) : instructionElement(w.names, part.text(item), true))
		at = part.byteEnd(item)
	}

	w.copySource(out, at, part.byteEnd(last), false)
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

// A source field with a new result (see the top).
function rewrite(w: StoryWriting, out: Out, node: Node): void {
	const part = w.part
	const document = part.document
	const field = rewritableSource(part, node)

	if (field === undefined) throw new DocxWriteError("unsupported", "A field whose result cannot be written anew")

	const names = w.names
	const result = node.attrs["result"] as string
	const removed = deleted(node)
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
		resultRun(w, out, run, result, removed)
		out.text(`</${document.name(xml)}>`)

		return
	}

	const beginRun = src(node)
	const endRun = node.attrs["end"] as number

	if (field.separate >= 0) {
		// Up to the separator, then its run closed.
		w.copySource(out, part.start(beginRun), part.byteEnd(field.separate), false)
		out.text(`</${document.name(part.node(part.parent(field.separate)))}>`)
		keepMarkers(w, out, field.separate + 1, field.end)
	} else {
		// No result yet: the code up to the end character, its run closed, and a separator.
		w.copySource(out, part.start(beginRun), part.start(field.end), false)
		out.text(`</${document.name(part.node(endRun))}><${names.el("r")}>`)
		runProperties(w, out, endRun)
		out.text(`${fieldChar(names, "separate")}</${names.el("r")}>`)
	}

	resultRun(w, out, run, result, removed)
	out.text(`<${names.el("r")}>`)
	runProperties(w, out, endRun)
	w.copyNode(out, part.node(field.end))
	out.text(`</${names.el("r")}>`)
}

// A field the editor made.
function writeNew(w: StoryWriting, out: Out, node: Node): void {
	const names = w.names
	const instruction = node.attrs["instruction"] as string
	const result = node.attrs["result"] as string
	const removed = deleted(node)
	const run = (inner: string): string => `<${names.el("r")}>${inner}</${names.el("r")}>`

	if (node.attrs["simple"] === true) {
		const tag = names.empty("fldSimple", [["instr", instruction]])

		out.text(result === "" ? tag : `${tag.slice(0, -2)}>${run(resultContent(names, result, removed))}</${names.el("fldSimple")}>`)

		return
	}

	out.text(
		run(fieldChar(names, "begin")) +
			run(instructionElement(names, instruction, removed)) +
			run(fieldChar(names, "separate")) +
			(result === "" ? "" : run(resultContent(names, result, removed))) +
			run(fieldChar(names, "end"))
	)
}

const FIELD: AtomWriter = {
	type: nodes.field,
	// Decided in write: a source field whose result is still the stored one is written as its bytes there.
	edited: node => src(node) !== NO_SOURCE,
	write(w, out, node) {
		const own = src(node)

		if (own === NO_SOURCE) {
			writeNew(w, out, node)

			return
		}

		// A legacy form field a form command filled in is the forms writer's.
		if (legacyFieldEdited(node)) {
			writeLegacyFormField(w, out, node)

			return
		}

		const part = w.part
		const end = node.attrs["end"] as number

		if ((node.attrs["result"] as string) === storedResult(part, own, end, part.kinds[own] === Item.FieldSimple)) {
			if (deleted(node)) copyDeleted(w, out, own, end)
			else w.copySource(out, part.start(own), part.byteEnd(end), false)

			return
		}

		rewrite(w, out, node)
	}
}

// A character of a field spanning paragraphs, inside the run opened for it; the begin holds the instruction.
const FIELD_CHAR: AtomWriter = {
	type: nodes.field_char,
	write(w, out, node) {
		const kind = node.attrs["kind"] as string

		if (kind !== "begin" && kind !== "separate" && kind !== "end")
			throw new DocxWriteError("unsupported", `A field character "${kind}"`)

		out.text(fieldChar(w.names, kind, kind === "begin" && node.attrs["dirty"] === true))

		if (kind === "begin") out.text(instructionElement(w.names, node.attrs["instruction"] as string, deleted(node)))
	}
}

export const fieldAtoms: readonly AtomWriter[] = [FIELD, FIELD_CHAR]
