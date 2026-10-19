import type { Mark, Node } from "prosemirror-model"
import { decodeUtf8, escapeText } from "../../../xml/index"
import type { RunBase } from "../../edit/formatting"
import { NO_SOURCE, docxSchema } from "../../edit/schema"
import { fieldUpdateOf, rewritableSource, simpleFieldRewritable, type SourceField } from "../../fields/evaluate"
import { fieldFormatRun, fieldFormatState, fieldFormatting, resultFormatRun, resultRunOf } from "../../fields/format"
import { storedResult, type FieldUpdate } from "../../fields/result"
import { Item, type StoryPart } from "../../model/index"
import { Ns } from "../../model/namespaces"
import { MARKER_ITEMS } from "../markers"
import { Out, spliceChildren, textElement, type Names } from "../markup"
import { PROPERTY_MARKS, writeRunProperties, type RunChange } from "../runs"
import { legacyFieldCopyData, legacyFieldEdited, writeLegacyFormField } from "../sdt"
import { RUN_PROPERTIES } from "../order.generated"
import { src } from "../source"
import type { StoryWriting } from "../story"
import { DocxWriteError } from "../values"
import { isRevisionWrapper } from "../wrappers"
import type { AtomWriter } from "./index"

// Fields. A source field whose result an update changed keeps its bytes up to its separator (instruction,
// w:dirty, w:fldLock, revisions), gets the new result in runs formatted as \* MERGEFORMAT asks (the previous
// result's first run, a w:fldSimple's too, else the field code's: ECMA-376 Part 1 §17.16.4.3), keeps the range markers and comment
// references its old result held, and ends with its own end character. Formatting a command gave the field
// (its atom's marks against those it was read with, fields/format.ts) goes into every run of it, each run's
// other properties kept, recorded per run in w:rPrChange when the change was tracked. An update made while
// tracking changes is written as Word records it: the result before in w:del, the new one in w:ins. A new
// field is complex (begin, instruction, separate, result, end) as Word writes them, or a w:fldSimple when it
// is simple; its runs take its formatting marks. A field_char is one character of a field spanning paragraphs
// (a table of contents).

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
const PROPERTY: ReadonlySet<unknown> = new Set(PROPERTY_MARKS)

// The w:rPr markup of a run of a field written with its formatting: source run `run` (NO_SOURCE for a new one),
// `fresh` when a new run takes its properties (its own tracked change is not the new run's).
type Restyle = (run: number, fresh: boolean) => string

// `marks` (a run's) with what changed from `from` to `to` (a field's formatting) applied.
function restyled(marksOf: readonly Mark[], from: readonly Mark[], to: readonly Mark[]): readonly Mark[] {
	let out: readonly Mark[] = marksOf.filter(mark => PROPERTY.has(mark.type))

	for (const type of PROPERTY_MARKS) {
		const was = type.isInSet(from)
		const now = type.isInSet(to)

		if (was === undefined ? now === undefined : now !== undefined && was.eq(now)) continue

		out = now === undefined ? type.removeFromSet(out) : now.addToSet(out)
	}

	return out
}

function sameProperties(a: readonly Mark[], b: readonly Mark[]): boolean {
	return PROPERTY_MARKS.every(type => {
		const one = type.isInSet(a)
		const other = type.isInSet(b)

		return one === undefined ? other === undefined : other !== undefined && one.eq(other)
	})
}

const EMPTY_PROPERTIES = /^<([\w.-]+:)?rPr\s*(\/>|><\/([\w.-]+:)?rPr>)$/

// How the runs of field `node` are formatted when it is written: undefined while its formatting (and tracked
// change of it) is the one it was read with (its runs stay as they are). Each run keeps what the field's marks
// did not change. A tracked change the field was read with (every run recording one: fields/format.ts) stays
// in each run, or, accepted or rejected, resolves in each run by its own: accepted, the run keeps its
// properties; rejected, it gets back those its own change recorded. A new tracked change is recorded in every
// run (a run already recording one from the file keeps what that one recorded before).
function restyleOf(w: StoryWriting, node: Node): Restyle | undefined {
	const own = src(node)
	const part = w.part
	const context = w.env.context
	const plain = marks.plain.isInSet(node.marks)
	const change = marks.format_change.isInSet(node.marks)
	const read =
		own === NO_SOURCE ? [] : fieldFormatting(context, part, own, node.attrs["end"] as number, part.kinds[own] === Item.FieldSimple)
	const readChange = marks.format_change.isInSet(read)
	// The change read with the field, still there: each run keeps its own.
	const kept = change !== undefined && readChange !== undefined && change.eq(readChange)

	if (
		own !== NO_SOURCE &&
		plain === undefined &&
		(change === undefined ? readChange === undefined : kept) &&
		sameProperties(read, node.marks)
	)
		return undefined

	const marksOfBase = (id: unknown): readonly Mark[] => (typeof id === "number" && id >= 0 ? (context.bases.get(id)?.marks ?? []) : [])
	const baseOf = (run: number): RunBase | undefined => {
		const rPr = run === NO_SOURCE ? -1 : (part.runs.rPr[part.aux[run] ?? -1] ?? -1)

		return rPr < 0 ? undefined : context.bases.of(part.document, rPr)
	}
	// What a run's own tracked change recorded before it (the run's properties when it records none).
	const ownBefore = (run: number, base: RunBase | undefined): { readonly marks: readonly Mark[]; readonly own: boolean } => {
		const rPr = run === NO_SOURCE ? -1 : (part.runs.rPr[part.aux[run] ?? -1] ?? -1)
		const recorded = rPr < 0 ? -1 : part.reader.child(rPr, Ns.W, "rPrChange")

		if (recorded < 0) return { marks: base?.marks ?? [], own: false }

		const properties = part.reader.child(recorded, Ns.W, "rPr")

		return { marks: properties < 0 ? [] : (context.bases.of(part.document, properties)?.marks ?? []), own: true }
	}
	const fresh = change === undefined || kept ? undefined : change
	const readBefore = readChange === undefined ? [] : marksOfBase(readChange.attrs["base"])
	// The read change accepted (the runs as they are) or rejected (as their own changes recorded them before):
	// as the editor saw it resolved; else (a state the editor did not follow, its formatting as resolved) by the
	// formatting, accepted when that cannot tell.
	const state = fieldFormatState(node)
	const settled = fresh === undefined ? node.marks : marksOfBase(fresh.attrs["base"])
	const rejected =
		readChange !== undefined &&
		!kept &&
		(state?.resolved !== undefined
			? state.resolved === "reject"
			: !sameProperties(settled, read) && sameProperties(settled, readBefore))
	// Tracked formatting over the read change: recorded in the runs that have no change of their own.
	const tracked = kept ? state?.tracked : undefined

	const names = w.names
	const anchorRead = rejected ? readBefore : read
	const properties = (marksOf: readonly Mark[], change: RunChange, base: RunBase | undefined): string => {
		const out = new Out()

		writeRunProperties(out, names, marksOf, base, change)

		const markup = decodeUtf8(out.finish())

		return EMPTY_PROPERTIES.test(markup) ? "" : markup
	}

	return (run, newRun) => {
		const base = baseOf(run)
		const before = ownBefore(run, base)
		const anchor = rejected ? before.marks : (base?.marks ?? [])
		const target =
			plain === undefined
				? restyled(anchor, anchorRead, node.marks)
				: plain.addToSet(node.marks.filter(mark => PROPERTY.has(mark.type)))
		// Kept: the field's change read with it, or the runs' own when the field records none.
		let runChange: RunChange =
			!newRun && (kept || (readChange === undefined && fresh === undefined)) ? { kind: "keep" } : { kind: "strip" }

		const record = (stamp: { readonly author: unknown; readonly date: unknown }, was: readonly Mark[]): RunChange => {
			const open = names
				.empty("rPrChange", [
					["id", String(w.env.nextRevisionId())],
					["author", (stamp.author as string | null) ?? ""],
					["date", (stamp.date as string | null) ?? undefined]
				])
				.replace(/\/>$/, ">")
			const recorded = properties(was, { kind: "strip" }, base)

			return { kind: "new", markup: `${open}${recorded === "" ? names.empty("rPr") : recorded}</${names.el("rPrChange")}>` }
		}

		// Before a change by a run recording one of its own from the file: what that one recorded.
		if (fresh !== undefined)
			runChange = record(
				{ author: fresh.attrs["author"], date: fresh.attrs["date"] },
				readChange === undefined && before.own && !newRun
					? before.marks
					: restyled(anchor, anchorRead, marksOfBase(fresh.attrs["base"]))
			)
		else if (tracked !== undefined && !before.own && !newRun) runChange = record(tracked, restyled(anchor, anchorRead, tracked.from))

		return properties(target, runChange, base)
	}
}

// Whether a tracked change inside `container` (a field's) deletes source item `item`.
function deletedWithin(part: StoryPart, item: number, container: number): boolean {
	for (let at = part.parent(item); at > container; at = part.parent(at))
		if (part.kinds[at] === Item.Deletion || part.kinds[at] === Item.MoveFrom) return true

	return false
}

// Source items `first` to `last` (siblings, of `container`) as their bytes up to byte `to`, their text and
// field codes as deleted text (w:delText, w:delInstrText) where the field is deleted (`removed`) or a tracked
// change inside it deletes them, else as text (a deletion of the field rejected).
function copyItems(
	w: StoryWriting,
	out: Out,
	first: number,
	last: number,
	to: number,
	container: number,
	removed: boolean,
	restyle?: Restyle
): void {
	const part = w.part
	const document = part.document
	let at = part.start(first)

	for (let item = first; item < part.end(last) && part.start(item) < to; item++) {
		const kind = part.kinds[item] ?? 0
		const code = kind === Item.InstrText || kind === Item.DeletedInstrText

		// A run of the field given its new properties in place of its w:rPr.
		if (restyle !== undefined && kind === Item.Run && !document.isSelfClosing(part.node(item))) {
			const rPr = part.runs.rPr[part.aux[item] ?? -1] ?? -1
			const openEnd = document.openEnd(part.node(item))

			w.copySource(out, at, openEnd, false)
			out.text(restyle(item, false))
			at = rPr >= 0 ? document.end(rPr) : openEnd

			continue
		}

		if (!code && kind !== Item.Text && kind !== Item.DeletedText) continue

		const gone = removed || deletedWithin(part, item, container)

		if (gone === DELETED_KINDS.has(kind)) continue

		w.copySource(out, at, part.start(item), false)
		out.text(code ? instructionElement(w.names, part.text(item), gone) : textElement(w.names, part.text(item), gone))
		at = part.byteEnd(item)
	}

	w.copySource(out, at, to, false)
}

// Whether item `item` is the first content of run `run` (only its properties, if any, before it).
function startsRun(part: StoryPart, run: number, item: number): boolean {
	const document = part.document
	const rPr = part.runs.rPr[part.aux[run] ?? -1] ?? -1
	const bytes = document.bytes

	for (let at = rPr >= 0 ? document.end(rPr) : document.openEnd(part.node(run)); at < part.start(item); at++) {
		const byte = bytes[at] ?? 0

		if (byte !== 0x20 && byte !== 0x09 && byte !== 0x0a && byte !== 0x0d) return false
	}

	return true
}

// A source run's properties without its tracked formatting change (whose id is the source run's), or as the
// field's formatting gives them.
function runProperties(w: StoryWriting, out: Out, run: number, restyle?: Restyle): void {
	if (restyle !== undefined) {
		out.text(restyle(run < 0 || w.part.kinds[run] !== Item.Run ? NO_SOURCE : run, true))

		return
	}

	const part = w.part
	const rPr = run < 0 || part.kinds[run] !== Item.Run ? -1 : (part.runs.rPr[part.aux[run] ?? -1] ?? -1)

	if (rPr >= 0) spliceChildren(out, part.reader, rPr, new Map([["rPrChange", ""]]), RUN_PROPERTIES)
}

// The run whose formatting a new result takes (fields/format.ts).
function formattingRun(part: StoryPart, field: SourceField, node: Node): number {
	return resultFormatRun(part, src(node), node.attrs["end"] as number, field.simple)
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

function resultRun(w: StoryWriting, out: Out, run: number, text: string, removed: boolean, restyle?: Restyle): void {
	if (text === "") return

	out.text(`<${w.names.el("r")}>`)
	runProperties(w, out, run, restyle)
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
// `deletedRun`: the run the result before a tracked update takes its properties from (else `run`'s).
function resultRuns(w: StoryWriting, out: Out, run: number, node: Node, restyle?: Restyle, deletedRun = run): void {
	const result = node.attrs["result"] as string
	const update = fieldUpdateOf(node)
	const removed = deleted(node)

	// (An update read from tracked changes in the code alone leaves the result as it is.)
	if (update === null || (update.code !== undefined && update.result === result)) {
		resultRun(w, out, run, result, removed, restyle)

		return
	}

	if (update.result !== "") {
		out.text(revisionStart(w, "del", update))
		resultRun(w, out, deletedRun, update.result, true, restyle)
		out.text(`</${w.names.el("del")}>`)
	}

	if (result !== "") {
		out.text(revisionStart(w, "ins", update))
		resultRun(w, out, run, result, removed, restyle)
		out.text(`</${w.names.el("ins")}>`)
	}
}

function sameUpdate(a: FieldUpdate | null, b: FieldUpdate | null): boolean {
	return (
		a === b ||
		(a !== null &&
			b !== null &&
			a.result === b.result &&
			a.author === b.author &&
			a.date === b.date &&
			a.id === b.id &&
			a.code?.before === b.code?.before &&
			a.code?.after === b.code?.after)
	)
}

// A source field with a new result (see the top).
function rewrite(w: StoryWriting, out: Out, node: Node, restyle: Restyle | undefined): void {
	const part = w.part
	const document = part.document
	const field = rewritableSource(part, node)

	if (field === undefined) throw new DocxWriteError("unsupported", "A field whose result cannot be written anew")

	const names = w.names
	// A tracked update accepted or rejected keeps the runs of the result it keeps; a result before a tracked
	// update keeps its own (the deleted runs of one read from the file).
	const end = node.attrs["end"] as number
	const stored = storedResult(part, src(node), end, field.simple)
	const resolved = stored.update !== null && fieldUpdateOf(node) === null
	// (An update read from tracked changes in the code alone has no result before.)
	const codeOnly = stored.update?.code !== undefined && stored.update.result === stored.result
	const kept =
		resolved && !codeOnly && (node.attrs["result"] as string) === stored.update?.result
			? resultRunOf(part, src(node), end, field.simple, true)
			: resolved && (node.attrs["result"] as string) === stored.result
				? resultRunOf(part, src(node), end, field.simple, false)
				: -1
	const run = kept >= 0 ? kept : formattingRun(part, field, node)
	const deletedRun = resultRunOf(part, src(node), end, field.simple, stored.update !== null)

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
		resultRuns(w, out, run, node, restyle, deletedRun >= 0 ? deletedRun : run)
		out.text(`</${document.name(xml)}>`)

		return
	}

	const beginRun = src(node)
	const endRun = node.attrs["end"] as number

	// Tracked changes in the code accepted or rejected: the code as the atom holds it, one instruction text.
	if (stored.update?.code !== undefined && fieldUpdateOf(node)?.code === undefined) {
		const inRun = (inner: () => void, properties = beginRun): void => {
			out.text(`<${names.el("r")}>`)
			runProperties(w, out, properties, restyle)
			inner()
			out.text(`</${names.el("r")}>`)
		}
		const separate = field.separate
		// The instruction in the properties of the code's first run (those the atom's are read from: see
		// fields/format.ts resultFormatRun).
		const code = part.fields.instruction[field.field] ?? -1

		inRun(() => w.copyNode(out, part.node(field.begin)))
		inRun(
			() => out.text(instructionElement(names, node.attrs["instruction"] as string, deleted(node))),
			code >= 0 ? part.parent(code) : beginRun
		)
		keepMarkers(w, out, field.begin + 1, separate >= 0 ? separate : field.end)

		if (separate >= 0) {
			inRun(() => w.copyNode(out, part.node(separate)), part.parent(separate))
			keepMarkers(w, out, separate + 1, field.end)
		} else inRun(() => out.text(fieldChar(names, "separate")), endRun)
	} else if (field.separate >= 0) {
		// Up to the separator, then its run closed.
		copyItems(w, out, beginRun, endRun, part.byteEnd(field.separate), part.parent(beginRun), deleted(node), restyle)
		out.text(`</${document.name(part.node(part.parent(field.separate)))}>`)
		keepMarkers(w, out, field.separate + 1, field.end)
	} else {
		// No result yet: the code up to the end character, its run closed (left out when it would hold no more
		// than its properties: the end character opens a run of its own), and a separator.
		const endAlone = endRun !== beginRun && startsRun(part, endRun, field.end)

		copyItems(
			w,
			out,
			beginRun,
			endRun,
			endAlone ? part.start(endRun) : part.start(field.end),
			part.parent(beginRun),
			deleted(node),
			restyle
		)

		if (!endAlone) out.text(`</${document.name(part.node(endRun))}>`)

		out.text(`<${names.el("r")}>`)
		runProperties(w, out, endRun, restyle)
		out.text(`${fieldChar(names, "separate")}</${names.el("r")}>`)
	}

	resultRuns(w, out, run, node, restyle, deletedRun >= 0 ? deletedRun : run)
	out.text(`<${names.el("r")}>`)
	runProperties(w, out, endRun, restyle)
	w.copyNode(out, part.node(field.end))
	out.text(`</${names.el("r")}>`)
}

// A tracked change holds runs only (CT_RunTrackChange, ECMA-376 Part 1 §17.13.5): a source w:fldSimple that a
// new w:ins or w:del wraps is written as the complex field it stands for, its instruction in a run of its own
// (deleted with the field) and w:dirty / w:fldLock on its begin character.
function writeSimpleAsComplex(w: StoryWriting, out: Out, node: Node, restyle: Restyle | undefined): void {
	const part = w.part
	const field = { begin: src(node) }
	const names = w.names
	const xml = part.node(field.begin)
	const on = (name: string): boolean => /^(1|true|on)$/.test(part.document.attribute(xml, names.attr(name)) ?? "")
	// Every run carries the properties the field is read with (its first result run's, fields/format.ts): a
	// complex field's are read from its code, and the result looks as it did.
	const run = fieldFormatRun(part, field.begin, node.attrs["end"] as number, true)
	const inRun = (inner: string): void => {
		out.text(`<${names.el("r")}>`)
		runProperties(w, out, run, restyle)
		out.text(`${inner}</${names.el("r")}>`)
	}

	inRun(
		names.empty("fldChar", [
			["fldCharType", "begin"],
			["fldLock", on("fldLock") ? "true" : null],
			["dirty", on("dirty") ? "true" : null]
		])
	)
	inRun(instructionElement(names, node.attrs["instruction"] as string, deleted(node)))
	inRun(fieldChar(names, "separate"))
	keepMarkers(w, out, field.begin + 1, part.end(field.begin))
	resultRuns(w, out, run, node, restyle)
	inRun(fieldChar(names, "end"))
}

// A field the editor made (`ffData`: a copy of a legacy form field's, in its begin character).
function writeNew(w: StoryWriting, out: Out, node: Node, ffData: string): void {
	const names = w.names
	const instruction = node.attrs["instruction"] as string
	const result = node.attrs["result"] as string
	const removed = deleted(node)
	const restyle = restyleOf(w, node) as Restyle
	const run = (inner: string): string => `<${names.el("r")}>${restyle(NO_SOURCE, true)}${inner}</${names.el("r")}>`
	const tracked = fieldUpdateOf(node) !== null
	// A tracked change holds runs only (CT_RunTrackChange, ECMA-376 Part 1 §17.13.5): a simple field in one (a copy
	// of a source w:fldSimple pasted or deleted while tracking) is written complex.
	const inChange = removed || marks.insertion.isInSet(node.marks) !== undefined || marks.move_to.isInSet(node.marks) !== undefined

	if (node.attrs["simple"] === true && !inChange) {
		const tag = names.empty("fldSimple", [["instr", instruction]])

		if (result === "" && !tracked) {
			out.text(tag)

			return
		}

		out.text(`${tag.slice(0, -2)}>`)
		resultRuns(w, out, NO_SOURCE, node, restyle)
		out.text(`</${names.el("fldSimple")}>`)

		return
	}

	const begin = fieldChar(names, "begin")

	out.text(
		run(ffData === "" ? begin : `${begin.slice(0, -2)}>${ffData}</${names.el("fldChar")}>`) +
			run(instructionElement(names, instruction, removed)) +
			run(fieldChar(names, "separate"))
	)
	resultRuns(w, out, NO_SOURCE, node, restyle)
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
			writeLegacyFormField(w, out, node, restyleOf(w, node))

			return
		}

		const part = w.part
		const end = node.attrs["end"] as number
		const stored = storedResult(part, own, end, part.kinds[own] === Item.FieldSimple)

		const restyle = restyleOf(w, node)

		if (part.kinds[own] === Item.FieldSimple && node.marks.some(mark => isRevisionWrapper(mark) && src(mark) === NO_SOURCE)) {
			// One holding a drawing or a field of its own has no complex form: tracked changes leave it out
			// (review/track.ts), and a save refuses one that a change wraps anyway.
			if (!simpleFieldRewritable(part, own))
				throw new DocxWriteError("unsupported", "A tracked change around a simple field that cannot be written inside one")

			writeSimpleAsComplex(w, out, node, restyle)

			return
		}

		if ((node.attrs["result"] as string) === stored.result && sameUpdate(fieldUpdateOf(node), stored.update)) {
			const simple = part.kinds[own] === Item.FieldSimple

			copyItems(w, out, own, end, part.byteEnd(end), simple ? own : part.parent(own), deleted(node), restyle)

			return
		}

		rewrite(w, out, node, restyle)
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
