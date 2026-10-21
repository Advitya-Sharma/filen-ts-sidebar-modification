import type { SheetRange } from "../../sheet/protocol"
import { moveRelativeReferences } from "./sharedFormula"
import { FormulaKind, type XlsxFormula, type XlsxFormulaTable } from "./types"

// A sheet's formulas by id, column by column: kinds and flags as bytes, shared indexes and cell metadata as
// integers, texts as strings (a shared child's is "" as written, and costs nothing), the rare ranges and
// data-table inputs in maps.
export class FormulaTable implements XlsxFormulaTable {
	private count = 0
	private kinds = new Uint8Array(64)
	private flags = new Uint8Array(64)
	private shared = new Int32Array(64)
	private texts: string[] = []
	private readonly refs = new Map<number, SheetRange>()
	private readonly metadata = new Map<number, number>()
	private readonly inputs = new Map<number, { readonly r1?: string; readonly r2?: string }>()
	// Shared index -> the master's id.
	private readonly masters = new Map<number, number>()
	// A master's own cell from its ref's top-left, by id, where it is not that cell (XlsxFormula.masterOffset).
	private readonly offsets = new Map<number, { readonly row: number; readonly col: number }>()
	private chars = 0

	get length(): number {
		return this.count
	}

	kind(id: number): FormulaKind {
		return (this.kinds[id] ?? 0) as FormulaKind
	}

	text(id: number): string {
		return this.texts[id] ?? ""
	}

	sharedIndex(id: number): number {
		return id < this.count ? (this.shared[id] ?? -1) : -1
	}

	ref(id: number): SheetRange | null {
		return this.refs.get(id) ?? null
	}

	// The master of a shared index (its formula id), or -1.
	master(sharedIndex: number): number {
		return this.masters.get(sharedIndex) ?? -1
	}

	formula(id: number): XlsxFormula {
		const inputs = this.inputs.get(id)
		const offset = this.offsets.get(id)

		return {
			kind: this.kind(id),
			text: this.text(id),
			sharedIndex: this.sharedIndex(id),
			ref: this.ref(id),
			flags: this.flags[id] ?? 0,
			cellMetadata: this.metadata.get(id) ?? -1,
			...(inputs?.r1 !== undefined && { r1: inputs.r1 }),
			...(inputs?.r2 !== undefined && { r2: inputs.r2 }),
			...(offset !== undefined && { masterOffset: offset })
		}
	}

	expanded(id: number, row: number, col: number): string {
		if (this.kind(id) !== FormulaKind.SHARED_CHILD) return this.text(id)

		const master = this.master(this.sharedIndex(id))
		const anchor = master < 0 ? null : this.ref(master)

		if (master < 0 || anchor === null) return this.text(id)

		// From the master's own cell, its ref's top-left unless the file put it elsewhere.
		const offset = this.offsets.get(master)

		return moveRelativeReferences(
			this.text(master),
			row - anchor.startRow - (offset?.row ?? 0),
			col - anchor.startCol - (offset?.col ?? 0)
		)
	}

	// The reader's: where a master it added lies, kept when that is not its ref's top-left.
	placeMaster(id: number, row: number, col: number): void {
		const ref = this.kind(id) === FormulaKind.SHARED_MASTER ? this.ref(id) : null

		if (ref !== null && (row !== ref.startRow || col !== ref.startCol))
			this.offsets.set(id, { row: row - ref.startRow, col: col - ref.startCol })
	}

	add(formula: XlsxFormula): number {
		const id = this.count

		if (id === this.kinds.length) this.grow()

		this.count++
		this.texts.push(formula.text)
		this.write(id, formula)

		return id
	}

	set(id: number, formula: XlsxFormula): void {
		if (id < 0 || id >= this.count) throw new RangeError(`No formula ${id}`)

		const previousShared = this.sharedIndex(id)

		if (this.kind(id) === FormulaKind.SHARED_MASTER && this.masters.get(previousShared) === id) this.masters.delete(previousShared)

		this.chars -= this.texts[id]?.length ?? 0
		this.texts[id] = formula.text
		this.write(id, formula)
	}

	// Only a formula's text: no object per formula (structure edits rewrite many).
	setText(id: number, text: string): void {
		if (id < 0 || id >= this.count) throw new RangeError(`No formula ${id}`)

		this.chars += text.length - (this.texts[id]?.length ?? 0)
		this.texts[id] = text
	}

	// The reader's add: no object per formula.
	addRead(
		kind: FormulaKind,
		text: string,
		sharedIndex: number,
		ref: SheetRange | null,
		flags: number,
		cellMetadata: number,
		r1?: string,
		r2?: string
	): number {
		const id = this.count

		if (id === this.kinds.length) this.grow()

		this.count++
		this.texts.push(text)
		this.kinds[id] = kind
		this.flags[id] = flags
		this.shared[id] = sharedIndex
		this.chars += text.length

		if (ref !== null) this.refs.set(id, ref)
		if (cellMetadata >= 0) this.metadata.set(id, cellMetadata)
		if (r1 !== undefined || r2 !== undefined) this.inputs.set(id, { ...(r1 !== undefined && { r1 }), ...(r2 !== undefined && { r2 }) })
		if (kind === FormulaKind.SHARED_MASTER && sharedIndex >= 0 && !this.masters.has(sharedIndex)) this.masters.set(sharedIndex, id)

		return id
	}

	private write(id: number, formula: XlsxFormula): void {
		this.kinds[id] = formula.kind
		this.flags[id] = formula.flags
		this.shared[id] = formula.sharedIndex
		this.chars += formula.text.length

		if (formula.ref === null) this.refs.delete(id)
		else this.refs.set(id, formula.ref)

		if (formula.cellMetadata < 0) this.metadata.delete(id)
		else this.metadata.set(id, formula.cellMetadata)

		if (formula.masterOffset === undefined || formula.kind !== FormulaKind.SHARED_MASTER) this.offsets.delete(id)
		else this.offsets.set(id, formula.masterOffset)

		if (formula.r1 === undefined && formula.r2 === undefined) this.inputs.delete(id)
		else
			this.inputs.set(id, {
				...(formula.r1 !== undefined && { r1: formula.r1 }),
				...(formula.r2 !== undefined && { r2: formula.r2 })
			})

		if (formula.kind === FormulaKind.SHARED_MASTER && formula.sharedIndex >= 0 && !this.masters.has(formula.sharedIndex)) {
			this.masters.set(formula.sharedIndex, id)
		}
	}

	private grow(): void {
		const size = this.kinds.length * 2
		const kinds = new Uint8Array(size)
		const flags = new Uint8Array(size)
		const shared = new Int32Array(size)

		kinds.set(this.kinds)
		flags.set(this.flags)
		shared.set(this.shared)
		this.kinds = kinds
		this.flags = flags
		this.shared = shared
	}

	// Trims the arrays once the reader is done.
	compact(): void {
		this.kinds = this.kinds.slice(0, Math.max(64, this.count))
		this.flags = this.flags.slice(0, Math.max(64, this.count))
		this.shared = this.shared.slice(0, Math.max(64, this.count))
	}

	estimateBytes(): number {
		return this.kinds.byteLength * 6 + this.texts.length * 16 + this.chars * 2 + (this.refs.size + this.metadata.size) * 64
	}
}
