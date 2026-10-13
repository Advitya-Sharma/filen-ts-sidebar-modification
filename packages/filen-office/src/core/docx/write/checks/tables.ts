import { runSliced } from "../../../slicing"
import { NodeKind, type XmlDocument } from "../../../xml/index"
import { Ns } from "../../model/namespaces"
import { partReader, type PartReader } from "../../model/reader"
import { DocxVerifyError } from "../failure"
import type { CheckedPart, CheckInput, DocxCheck } from "./index"

// The tables a save wrote keep their grid: no row (w:gridBefore, its cells' spans, w:gridAfter) wider than
// the w:tblGrid, every w:vMerge "continue" under a cell of the same columns that takes part in a merge, every
// cell ending with a paragraph, gridCol widths that are twips, and tracked row and cell changes with an id
// and an author: insertions and deletions, merges, and property changes holding the properties before them
// (w:tblGridChange, CT_Markup, has an id only). In a rewritten story only the tables with new markup are read, the source copies inside them
// (placeholders in the checked skeleton) from the source part. A table as broken as its source was is let
// through: the save refuses only more of a kind of fault than the stories it rewrote had.

type Fault = "wide" | "orphan" | "unended" | "grid" | "record"

const FAULTS: readonly Fault[] = ["wide", "orphan", "unended", "grid", "record"]
const MESSAGES: Readonly<Record<Fault, string>> = {
	wide: "a table row is wider than its grid",
	orphan: "a vertical merge continues under no merge",
	unended: "a table cell does not end with a paragraph",
	grid: "a grid column has no width in twips",
	record: "a tracked table, row or cell change has no id, author or previous properties"
}
const TWIPS = /^\d+$/

// A node of the skeleton (new markup) or of the source part (what a placeholder stands for).
interface View {
	readonly doc: XmlDocument
	readonly reader: PartReader
	readonly node: number
}

class Resolver {
	private readonly skeleton: XmlDocument | undefined
	private readonly source: XmlDocument | undefined
	private readonly sourceReader: PartReader | undefined
	// Placeholder nodes of the skeleton in order, and the source byte ranges they stand for (two per span).
	private readonly holders: readonly number[]
	private readonly spans: readonly number[]
	// A placeholder whose source nodes could not be found: its table is not judged.
	unknown = false

	constructor(skeleton: XmlDocument | undefined, source: XmlDocument | undefined, holders: readonly number[], spans: readonly number[]) {
		this.skeleton = skeleton
		this.source = source
		this.sourceReader = source === undefined ? undefined : partReader(source)
		this.holders = holders
		this.spans = spans
	}

	// The element children of `view`, placeholders replaced by the source nodes they stand for.
	children(view: View): View[] {
		const out: View[] = []
		const doc = view.doc

		for (let child = doc.firstChild(view.node); child >= 0; child = doc.nextSibling(child)) {
			if (doc.kinds[child] !== NodeKind.Element) continue

			if (doc === this.skeleton && doc.name(child) === "_") this.expand(child, out)
			else out.push({ doc, reader: view.reader, node: child })
		}

		return out
	}

	private expand(holder: number, out: View[]): void {
		const source = this.source
		const reader = this.sourceReader
		const ordinal = binarySearch(this.holders, holder)
		const from = this.spans[ordinal * 2]
		const to = this.spans[ordinal * 2 + 1]

		if (source === undefined || reader === undefined || ordinal < 0 || from === undefined || to === undefined) {
			this.unknown = true

			return
		}

		let node = binarySearch(source.starts, from, source.count)

		if (node < 0) {
			this.unknown = true

			return
		}

		for (; node >= 0 && source.start(node) < to; node = source.nextSibling(node)) {
			if (source.kinds[node] === NodeKind.Element) out.push({ doc: source, reader, node })
		}
	}
}

// The index of `value` in ascending `list` (its first `count` entries), or -1.
function binarySearch(list: ArrayLike<number>, value: number, count: number = list.length): number {
	let low = 0
	let high = count - 1

	while (low <= high) {
		const middle = (low + high) >> 1
		const at = list[middle] ?? 0

		if (at === value) return middle
		if (at < value) low = middle + 1
		else high = middle - 1
	}

	return -1
}

function isW(view: View, local: string): boolean {
	return view.reader.group(view.node) === Ns.W && view.doc.localName(view.node) === local
}

function first(views: readonly View[], local: string): View | undefined {
	return views.find(view => isW(view, local))
}

function intVal(view: View | undefined, fallback: number): number {
	if (view === undefined) return fallback

	const value = Number(view.reader.attr(view.node, Ns.W, "val"))

	return Number.isInteger(value) && value >= 0 ? value : fallback
}

type Counts = Record<Fault, number>

function noCounts(): Counts {
	return { wide: 0, orphan: 0, unended: 0, grid: 0, record: 0 }
}

// A tracked change element of new markup with what Word needs of it: an id, and an author (`author`).
function recorded(view: View, author = true): boolean {
	const id = view.reader.attr(view.node, Ns.W, "id")
	const name = view.reader.attr(view.node, Ns.W, "author")

	return id !== undefined && TWIPS.test(id) && (!author || (name !== undefined && name !== ""))
}

// Property change elements (w:tblPrChange...) by the local name of the element before them they hold.
const PROPERTY_CHANGES: Readonly<Record<string, string>> = {
	tblPrChange: "tblPr",
	tblGridChange: "tblGrid",
	trPrChange: "trPr",
	tblPrExChange: "tblPrEx",
	tcPrChange: "tcPr"
}
const MARKS = new Set(["ins", "del", "cellIns", "cellDel", "cellMerge"])

// Faults of the tracked change records among the children of a property element (new markup only).
function recordFaults(children: readonly View[], resolver: Resolver, skeleton: XmlDocument | undefined): number {
	let faults = 0

	for (const child of children) {
		if (skeleton !== undefined && child.doc !== skeleton) continue
		if (child.reader.group(child.node) !== Ns.W) continue

		const local = child.doc.localName(child.node)
		const inner = PROPERTY_CHANGES[local]

		if (MARKS.has(local)) {
			if (!recorded(child)) faults++
		} else if (inner !== undefined) {
			if (!recorded(child, local !== "tblGridChange") || first(resolver.children(child), inner) === undefined) faults++
		}
	}

	return faults
}

// Faults of one table, some rows at a time (a long table of copied rows is long to read); nested tables are
// found on their own.
function* tableFaults(table: View, resolver: Resolver, skeleton: XmlDocument | undefined, counts: Counts): Generator<void, void> {
	const children = resolver.children(table)
	const tblPr = first(children, "tblPr")
	const tblGrid = first(children, "tblGrid")
	let columns = 0

	if (tblPr !== undefined) counts.record += recordFaults(resolver.children(tblPr), resolver, skeleton)

	if (tblGrid !== undefined) {
		const cols = resolver.children(tblGrid)

		counts.record += recordFaults(cols, resolver, skeleton)

		for (const col of cols) {
			if (!isW(col, "gridCol")) continue

			columns++

			if ((skeleton === undefined || col.doc === skeleton) && !TWIPS.test(col.reader.attr(col.node, Ns.W, "w") ?? "")) counts.grid++
		}
	}

	// Per row: the cells' first columns and spans, and whether each takes part in a merge.
	let above: Cells | undefined
	let rows = 0

	for (const row of children) {
		if (!isW(row, "tr")) continue

		above = rowFaults(row, above, tblGrid === undefined ? undefined : columns, resolver, skeleton, counts)

		if (++rows % ROW_STEP === 0) yield
	}
}

// Rows of a table read between yields.
const ROW_STEP = 32

// A row's cells by first column: their spans, and whether each takes part in a merge.
type Cells = Map<number, { span: number; merged: boolean }>

// Faults of one row under the row `above` (when judged), against a grid of `columns` (when it has one); the
// row's cells, when it is judged.
function rowFaults(
	row: View,
	above: Cells | undefined,
	columns: number | undefined,
	resolver: Resolver,
	skeleton: XmlDocument | undefined,
	counts: Counts
): Cells | undefined {
	const parts = resolver.children(row)
	const trPr = first(parts, "trPr")
	const trChildren = trPr === undefined ? [] : resolver.children(trPr)
	let column = intVal(first(trChildren, "gridBefore"), 0)
	const here: Cells = new Map()
	let judged = true

	counts.record += recordFaults(trChildren, resolver, skeleton)

	const tblPrEx = first(parts, "tblPrEx")

	if (tblPrEx !== undefined) counts.record += recordFaults(resolver.children(tblPrEx), resolver, skeleton)

	for (const part of parts) {
		if (isW(part, "trPr") || isW(part, "tblPrEx")) continue
		// Cells inside content controls or custom XML: not judged.
		if (!isW(part, "tc")) {
			if (isW(part, "sdt") || isW(part, "customXml")) judged = false

			continue
		}

		const content = resolver.children(part)
		const tcPr = first(content, "tcPr")
		const props = tcPr === undefined ? [] : resolver.children(tcPr)
		const span = Math.max(1, intVal(first(props, "gridSpan"), 1))
		const vMerge = first(props, "vMerge")
		const value = vMerge === undefined ? null : (vMerge.reader.attr(vMerge.node, Ns.W, "val") ?? "continue")

		if (value === "continue") {
			const cell = above?.get(column)

			if (cell === undefined || cell.span !== span || !cell.merged) counts.orphan++
		}

		counts.record += recordFaults(props, resolver, skeleton)

		const blocks = content.filter(view => !isW(view, "tcPr"))
		const last = blocks.at(-1)

		if (last === undefined || isW(last, "tbl")) counts.unended++

		here.set(column, { span, merged: value !== null })
		column += span
	}

	column += intVal(first(trChildren, "gridAfter"), 0)

	if (judged && columns !== undefined && column > columns) counts.wide++

	return judged ? here : undefined
}

// Faults of every table in `doc` under `root` (nested ones included), a step at a time.
function* faultsUnder(
	doc: XmlDocument,
	root: number,
	resolver: Resolver,
	skeleton: XmlDocument | undefined,
	counts: Counts
): Generator<void, void> {
	const reader = partReader(doc)
	const end = root < 0 ? doc.count : nodeEnd(doc, root)

	for (let from = Math.max(0, root); from < end;) {
		const to = Math.min(end, from + NODE_STEP)
		const node = nextTable(doc, reader, from, to)

		if (node < 0) {
			from = to

			yield

			continue
		}

		resolver.unknown = false

		const table = noCounts()

		yield* tableFaults({ doc, reader, node }, resolver, skeleton, table)

		if (!resolver.unknown) for (const fault of FAULTS) counts[fault] += table[fault]

		from = node + 1
	}
}

// Nodes looked at between yields.
const NODE_STEP = 8192

// The first w:tbl among nodes [from, to), or -1.
function nextTable(doc: XmlDocument, reader: PartReader, from: number, to: number): number {
	for (let node = from; node < to; node++)
		if (doc.kinds[node] === NodeKind.Element && reader.group(node) === Ns.W && doc.localName(node) === "tbl") return node

	return -1
}

// The node after `node`'s last descendant.
function nodeEnd(doc: XmlDocument, node: number): number {
	for (let at = node; at >= 0; at = doc.parent(at)) {
		const next = doc.nextSibling(at)

		if (next >= 0) return next
	}

	return doc.count
}

// Faults of the source tables in the stories the save rewrote on this part (the source part when the save
// replaced it whole).
async function sourceFaults(input: CheckInput, part: CheckedPart): Promise<Counts> {
	const counts = noCounts()
	const story = part.part

	if (story !== undefined && part.skeleton) {
		const resolver = new Resolver(undefined, undefined, [], [])

		for (const change of input.changes.stories) {
			if (change.part !== story) continue

			const item = change.session.story.story.item

			if (item >= 0) await runSliced(faultsUnder(story.document, story.node(item), resolver, undefined, counts), input.clock)
		}

		return counts
	}

	if (!input.original.has(part.name)) return counts

	const original = story?.document ?? (await input.original.xml(part.name))

	await runSliced(faultsUnder(original, -1, new Resolver(undefined, undefined, [], []), undefined, counts), input.clock)

	return counts
}

async function checkTables(input: CheckInput, part: CheckedPart): Promise<void> {
	const doc = part.document
	const reader = partReader(doc)
	const holders: number[] = []
	let tables = false

	for (let node = 0; node < doc.count; node++) {
		if (doc.kinds[node] !== NodeKind.Element) continue

		if (part.skeleton && doc.name(node) === "_") holders.push(node)
		else if (!tables && reader.group(node) === Ns.W && doc.localName(node) === "tbl") tables = true
	}

	if (!tables) return

	const spans: number[] = []

	if (part.skeleton) {
		for (const change of input.changes.stories) {
			if (change.part !== part.part) continue

			const trusted = change.output.trusted

			for (let k = 0; k < trusted.length; k += 4) spans.push(trusted[k + 2] ?? 0, trusted[k + 3] ?? 0)
		}
	}

	const skeleton = part.skeleton ? doc : undefined
	const resolver = new Resolver(skeleton, part.part?.document, holders, spans)
	const counts = noCounts()

	await runSliced(faultsUnder(doc, -1, resolver, skeleton, counts), input.clock)

	if (FAULTS.every(fault => counts[fault] === 0)) return

	const before = await sourceFaults(input, part)

	for (const fault of FAULTS) {
		if (counts[fault] > before[fault]) throw new DocxVerifyError("content", `${part.name}: ${MESSAGES[fault]}`)
	}
}

export const tableChecks: readonly DocxCheck[] = [
	{
		async part(input, part) {
			if (part.part !== undefined) await checkTables(input, part)
		}
	}
]
