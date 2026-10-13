import type { Mark, Node } from "prosemirror-model"
import type { DocxEditContext } from "../edit/context"
import { docxSchema } from "../edit/schema"
import { Item, RevisionView, type StoryPart } from "../model/index"
import { Ns } from "../model/namespaces"
import { mergeFormat, parseFieldCode } from "./evaluate"

// A source field's formatting as its atom carries it (ECMA-376 Part 1 §17.16.4.3): the direct properties of the
// run a new result takes them from (the previous result's first run under \* MERGEFORMAT, else the run of the
// field code's first character; a w:fldSimple's first result run, which an update without \* MERGEFORMAT writes
// anew without properties), so an update keeps what the atom shows; with that run's tracked formatting change (w:rPrChange), the properties before it as a run base,
// as a change the editor recorded carries them, when every run of the field records one change (Word's, and
// the writer's, for a field formatted at once; a field only some of whose runs record one shows none). A command formatting the field changes those marks; the writer gives every run of the
// field what changed. Read the same by the builder and the writer, so formatting no command changed is told
// apart.

const formatChange = docxSchema.marks.format_change

// The run a new result of a source field takes its properties from (§17.16.4.3): the previous result's first
// run under \* MERGEFORMAT, else the run of the code's first character (a w:fldSimple has none: -1).
export function resultFormatRun(part: StoryPart, first: number, last: number, simple: boolean): number {
	const field = simple ? (part.aux[first] ?? -1) : part.fieldOf(first + 1)

	if (field < 0) return -1

	// The code as the atom holds it (a field deleted whole has its code deleted, one inserted and deleted both).
	const instruction =
		part.fieldInstruction(field) ||
		part.fieldInstruction(field, RevisionView.Original) ||
		part.fieldInstruction(field, RevisionView.Markup)

	if (mergeFormat(parseFieldCode(instruction))) {
		const run = firstResultRun(part, first, last, simple, field)

		if (run >= 0) return run
	}

	if (simple) return -1

	const code = part.fields.instruction[field] ?? -1

	return code >= 0 ? part.parent(code) : first
}

// The run whose properties stand for the field's (see the top), or -1: the one a new result takes them from,
// a w:fldSimple's first result run (its result shows them until an update drops them).
export function fieldFormatRun(part: StoryPart, first: number, last: number, simple: boolean): number {
	if (!simple) return resultFormatRun(part, first, last, simple)

	return firstResultRun(part, first, last, simple, part.aux[first] ?? -1)
}

function firstResultRun(part: StoryPart, first: number, last: number, simple: boolean, field: number): number {
	const from = simple ? first : field < 0 ? -1 : (part.fields.separate[field] ?? -1)
	const to = simple ? part.end(first) : Math.min(part.fields.end[field] ?? -1, part.end(last))

	if (from >= 0) for (let item = from + 1; item < to; item++) if (part.kinds[item] === Item.Run) return item

	return -1
}

// The formatting marks a source field's atom is read with (see the top).
export function fieldFormatting(context: DocxEditContext, part: StoryPart, first: number, last: number, simple: boolean): readonly Mark[] {
	const run = fieldFormatRun(part, first, last, simple)
	const rPr = run < 0 ? -1 : (part.runs.rPr[part.aux[run] ?? -1] ?? -1)
	const base = rPr < 0 ? undefined : context.bases.of(part.document, rPr)
	const out = base === undefined ? [] : base.marks.filter(mark => mark.type !== formatChange)
	const reader = part.reader
	// The oldest tracked formatting change among the field's runs is the field's (one formatting the field at
	// once records the same one in each run; a run some change left alone has none of its own).
	let oldest: { author: string | undefined; date: string | undefined; id: string | undefined } | undefined

	for (const each of fieldRuns(part, first, last)) {
		const properties = part.runs.rPr[part.aux[each] ?? -1] ?? -1
		const change = properties < 0 ? -1 : reader.child(properties, Ns.W, "rPrChange")

		if (change < 0) continue

		const date = reader.attr(change, Ns.W, "date") ?? undefined

		if (oldest === undefined || (date ?? "") < (oldest.date ?? ""))
			oldest = { author: reader.attr(change, Ns.W, "author") ?? undefined, date, id: reader.attr(change, Ns.W, "id") ?? undefined }
	}

	if (oldest === undefined) return out

	// What rejecting it gives the field: what its own run's change recorded, or the run as it is.
	const own = rPr < 0 ? -1 : reader.child(rPr, Ns.W, "rPrChange")
	const properties = own < 0 ? -1 : reader.child(own, Ns.W, "rPr")
	const before = own < 0 ? base : properties < 0 ? undefined : context.bases.of(part.document, properties)
	const id = Number(oldest.id)

	return formatChange
		.create({
			id: Number.isInteger(id) ? id : null,
			author: oldest.author ?? null,
			date: oldest.date ?? null,
			base: before === undefined ? -1 : before.id,
			run: null
		})
		.addToSet(out)
}

// How a field's tracked formatting change read from the file was resolved in the editor (accepted: its runs
// keep their properties; rejected: each gets back what its own change recorded), and a tracked formatting the
// editor recorded over it while it stayed (who, and the field's formatting before): kept with the field's atom
// as its node changes (edit/invariants/fields.ts), read by the writer.
export interface FieldFormatState {
	readonly resolved?: "accept" | "reject"
	readonly tracked?: { readonly author: string; readonly date: string; readonly from: readonly Mark[] }
}

const STATES = new WeakMap<Node, FieldFormatState>()

export function fieldFormatState(node: Node): FieldFormatState | undefined {
	return STATES.get(node)
}

export function setFieldFormatState(node: Node, state: FieldFormatState | undefined): void {
	if (state === undefined) STATES.delete(node)
	else STATES.set(node, state)
}

// The runs of a source field (its own and those of fields nested in it).
export function fieldRuns(part: StoryPart, first: number, last: number): number[] {
	const out: number[] = []

	for (let item = first, end = part.end(last); item < end; item++) if (part.kinds[item] === Item.Run) out.push(item)

	return out
}

// Whether an update writes a field's new result without its properties (a w:fldSimple without \* MERGEFORMAT:
// §17.16.4.3, it has no code to take them from): its atom then keeps no formatting.
export function resultDropsFormatting(node: Node): boolean {
	return node.attrs["simple"] === true && !mergeFormat(parseFieldCode(node.attrs["instruction"] as string))
}

// Whether a document or block holds a field that cannot take formatting (see holdsFormatting), by node: known
// from where it was built, then carried along the edits and kept for blocks looked at (edit/invariants/fields.ts).
export const unformattableFields = new WeakMap<Node, boolean>()

// Whether a field can take formatting a command gives: all but a w:fldSimple without a result (the file's or a
// copy of one), which has no run to hold it (commands leave it out; edit/invariants/fields.ts takes formatting
// given to it off).
export function holdsFormatting(node: Node): boolean {
	return node.attrs["simple"] !== true || node.attrs["result"] !== ""
}

// The first run of a source field's result showing text that a tracked change inside the field deletes
// (`deleted`) or does not, or -1: the runs of a result before and after a tracked update.
export function resultRunOf(part: StoryPart, first: number, last: number, simple: boolean, deleted: boolean): number {
	const field = simple ? (part.aux[first] ?? -1) : part.fieldOf(first + 1)
	const from = simple ? first : field < 0 ? -1 : (part.fields.separate[field] ?? -1)
	const container = simple ? first : part.parent(first)

	if (from < 0) return -1

	for (let item = from + 1, end = part.end(last); item < end; item++) {
		const kind = part.kinds[item]

		if (kind !== Item.Text && kind !== Item.DeletedText) continue

		const run = part.parent(item)
		let gone = false

		for (let at = part.parent(run); at > container; at = part.parent(at))
			if (part.kinds[at] === Item.Deletion || part.kinds[at] === Item.MoveFrom) gone = true

		if (part.kinds[run] === Item.Run && gone === deleted) return run
	}

	return -1
}
