import { XmlEditor, XmlError, parseXml, type Markup, type XmlDocument } from "../../xml/index"
import type { SheetRange } from "../../sheet/protocol"
import { preimageCell, shiftIndex, type AxisShift } from "./axis"
import { refuse, type RefConstruct } from "./catalogue"
import type { FormulaShift } from "./formulas"
import { formatSqref, parseRef, parseSqref, shiftRangeList, shiftRanges, topLeft } from "./sqref"

// What one sheet's (or part's) rewrite needs: the edit as it applies here, and where to report.
export interface RefJob {
	// The sheet the markup belongs to (unqualified references point into it); null for a workbook-level part.
	readonly own: string | null
	// Set when positions on this sheet move.
	readonly shift: AxisShift | null
	readonly formulas: FormulaShift
	touch(construct: RefConstruct): void
}

// Edits recorded first and applied at serialize: a rule may drop an element after another rule edited inside it
// (a sort state losing its last condition), and an edit inside a removed element is then simply not made.
export class Edits {
	private readonly document: XmlDocument
	private readonly removed = new Set<number>()
	private readonly ops: { readonly node: number; readonly apply: (editor: XmlEditor) => void }[] = []

	constructor(document: XmlDocument) {
		this.document = document
	}

	get changed(): boolean {
		return this.ops.length > 0 || this.removed.size > 0
	}

	setAttribute(node: number, name: string, value: string): void {
		this.ops.push({ node, apply: editor => editor.setAttribute(node, name, value) })
	}

	setText(node: number, text: string): void {
		this.ops.push({ node, apply: editor => editor.setText(node, text) })
	}

	insertAfter(node: number, markup: Markup): void {
		this.ops.push({ node, apply: editor => editor.insertAfter(node, markup) })
	}

	prepend(node: number, markup: Markup): void {
		this.ops.push({ node, apply: editor => editor.prepend(node, markup) })
	}

	remove(node: number): void {
		this.removed.add(node)
	}

	private inside(node: number, self: boolean): boolean {
		for (let at = self ? node : this.document.parent(node); at >= 0; at = this.document.parent(at))
			if (this.removed.has(at)) return true

		return false
	}

	serialize(): Uint8Array {
		const editor = new XmlEditor(this.document)

		for (const node of this.removed) if (!this.inside(node, false)) editor.remove(node)
		for (const op of this.ops) if (!this.inside(op.node, true)) op.apply(editor)

		return editor.serialize()
	}
}

export interface Edited {
	readonly document: XmlDocument
	readonly editor: Edits
}

// Parses markup for an edit; markup that does not read refuses the edit (`what` names it).
export function editMarkup(bytes: Uint8Array, what: string): Edited {
	try {
		const document = parseXml(bytes)

		return { document, editor: new Edits(document) }
	} catch (error) {
		if (error instanceof XmlError) refuse("structureLocked", `${what}: malformed`)

		throw error
	}
}

export function childElements(document: XmlDocument, node: number): number[] {
	const out: number[] = []

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) out.push(child)

	return out
}

export function childByLocal(document: XmlDocument, node: number, local: string): number {
	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child))
		if (document.localName(child) === local) return child

	return -1
}

// Every element below `node` (itself included), in document order.
export function* descendants(document: XmlDocument, node: number): Generator<number> {
	yield node

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child))
		yield* descendants(document, child)
}

// Formulas written relative to a range list's top-left cell (the rectangle around its ranges: conditional formats,
// validations) are written again for the cell that is the top-left after the edit: `rows`/`cols` from the old
// top-left to where that cell was.
export interface RangeAnchor {
	readonly rows: number
	readonly cols: number
}

const STILL: RangeAnchor = { rows: 0, cols: 0 }

// A range list attribute (sqref) moved. Returns the moved ranges (empty: all deleted) and the anchor move, or null
// when the attribute is absent. Unreadable lists refuse. `mode`: how the list follows cells shifted in a band
// (shiftRangeList).
export function shiftSqrefAttribute(
	edited: Edited,
	node: number,
	name: string,
	job: RefJob,
	construct: RefConstruct,
	mode: "list" | "cells" = "list"
): { ranges: readonly SheetRange[]; anchor: RangeAnchor } | null {
	const text = edited.document.attribute(node, name)

	if (text === undefined) return null

	return shiftSqrefText(text, job, construct, next => edited.editor.setAttribute(node, name, next), mode)
}

export function shiftSqrefText(
	text: string,
	job: RefJob,
	construct: RefConstruct,
	write: (next: string) => void,
	mode: "list" | "cells" = "list"
): { ranges: readonly SheetRange[]; anchor: RangeAnchor } {
	const ranges = parseSqref(text)

	if (ranges === null) refuse("structureLocked", `${construct}: range "${text}"`)

	const first = ranges[0]

	if (job.shift === null || first === undefined) return { ranges, anchor: STILL }

	const next = shiftRangeList(ranges, job.shift, mode)
	const head = next[0]

	if (next !== ranges) {
		job.touch(construct)

		if (next.length > 0) write(formatSqref(next))
	}

	if (head === undefined) return { ranges: next, anchor: STILL }

	// The cell now at the new top-left was at its preimage before.
	const before = topLeft(ranges)
	const after = topLeft(next)
	const was = preimageCell(after.row, after.col, job.shift)

	return { ranges: next, anchor: { rows: was.row - before.row, cols: was.col - before.col } }
}

// A single range attribute (ref) moved: false when the range was deleted (the caller drops the element), true
// otherwise (also when absent).
export function shiftRefAttribute(edited: Edited, node: number, name: string, job: RefJob, construct: RefConstruct): boolean {
	const text = edited.document.attribute(node, name)

	if (text === undefined || job.shift === null) return true

	const range = parseRef(text)

	if (range === null) refuse("structureLocked", `${construct}: range "${text}"`)

	const next = shiftRanges([range], job.shift)

	if (next.length === 0) return false

	if (next[0] !== range) {
		job.touch(construct)
		edited.editor.setAttribute(node, name, formatSqref(next))
	}

	return true
}

// A cell attribute moved: false when the cell was deleted.
export function shiftCellAttribute(edited: Edited, node: number, name: string, job: RefJob, construct: RefConstruct): boolean {
	return shiftRefAttribute(edited, node, name, job, construct)
}

// A formula held as an element's text or an attribute, anchored at `anchor` (relative references are relative
// to it) or plain.
export function shiftFormulaText(edited: Edited, node: number, job: RefJob, construct: RefConstruct, anchor: RangeAnchor | null): void {
	const text = edited.document.textContent(node)
	const next = anchoredShift(text, job, anchor)

	if (next !== text) {
		job.touch(construct)
		edited.editor.setText(node, next)
	}
}

export function shiftFormulaAttribute(
	edited: Edited,
	node: number,
	name: string,
	job: RefJob,
	construct: RefConstruct,
	anchor: RangeAnchor | null
): void {
	const text = edited.document.attribute(node, name)

	if (text === undefined || text === "") return

	const next = anchoredShift(text, job, anchor)

	if (next !== text) {
		job.touch(construct)
		edited.editor.setAttribute(node, name, next)
	}
}

function anchoredShift(text: string, job: RefJob, anchor: RangeAnchor | null): string {
	if (anchor === null || (anchor.rows === 0 && anchor.cols === 0)) return job.formulas.shift(text, job.own)

	return job.formulas.shiftAnchored(text, job.own, anchor.rows, anchor.cols)
}

// A zero-based index held as text or attribute on the shifted axis: -1 when deleted.
export function shiftedIndex(text: string | undefined, shift: AxisShift): number | null {
	if (text === undefined) return null

	const value = Number(text)

	if (!Number.isInteger(value) || value < 0) return null

	return shiftIndex(value, shift)
}
