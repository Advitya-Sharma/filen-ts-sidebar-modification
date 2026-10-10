import type { DocState, EditOp } from "@/features/spreadsheet/lib/edits"

// Excel's own limit.
const HISTORY_LIMIT = 100

// What the history may hold in all, in bytes as each document estimates its steps: a few edits of the
// largest size (MAX_EDIT_CELLS cells) fit, or many small ones.
export const HISTORY_BYTES = 64 * 1024 * 1024

// `before` and `after` identify the document's states either side of the step: equal ids, equal content.
interface UndoEntry<S> {
	step: S
	op: EditOp
	before: number
	after: number
	weight: number
}

// What an op held for redo costs: typed or pasted text is the bulk of it.
function opBytes(op: EditOp): number {
	switch (op.type) {
		case "setCells": {
			let bytes = 64

			for (const cell of op.cells) bytes += 48 + cell.input.length

			return bytes
		}
		case "resize":
			return 64 + op.sizes.length * 32
		default:
			return 64
	}
}

// A document's undo and redo history, S being what undoing one step needs. The oldest steps go once there
// are more than HISTORY_LIMIT or they hold more than `budget` bytes in all (as `weigh` estimates them),
// always keeping the latest.
// State ids are never reused, so a save point whose step falls out of history stays unreachable, and dirty
// stays true, without bookkeeping.
export class EditHistory<S> {
	private readonly undoSteps: UndoEntry<S>[] = []
	private redoSteps: { op: EditOp; after: number }[] = []
	private weight = 0
	// Identifies the current state; `saved` is the state last written to the file (0: as opened).
	private current = 0
	private nextState = 1
	private saved = 0
	private readonly weigh: (step: S) => number
	private readonly budget: number

	constructor(weigh: (step: S) => number, budget = HISTORY_BYTES) {
		this.weigh = weigh
		this.budget = budget
	}

	// The current state's id.
	get version(): number {
		return this.current
	}

	state(): DocState {
		return { dirty: this.current !== this.saved, canUndo: this.undoSteps.length > 0, canRedo: this.redoSteps.length > 0 }
	}

	// A new edit's step: what could be redone is gone.
	record(step: S, op: EditOp): void {
		const after = this.nextState++

		this.push({ step, op, before: this.current, after, weight: this.weigh(step) + opBytes(op) })
		this.redoSteps = []
		this.current = after
	}

	peekUndo(): S | undefined {
		return this.undoSteps.at(-1)?.step
	}

	// The step peekUndo gave was reverted.
	popUndo(): void {
		const last = this.undoSteps.pop()

		if (last === undefined) {
			return
		}

		this.weight -= last.weight
		this.redoSteps.push({ op: last.op, after: last.after })
		this.current = last.before
	}

	peekRedo(): EditOp | undefined {
		return this.redoSteps.at(-1)?.op
	}

	// The op peekRedo gave was run again, as `step`.
	commitRedo(step: S): void {
		const redone = this.redoSteps.pop()

		if (redone === undefined) {
			return
		}

		this.push({ step, op: redone.op, before: this.current, after: redone.after, weight: this.weigh(step) + opBytes(redone.op) })
		this.current = redone.after
	}

	markSaved(version: number): DocState {
		this.saved = version

		return this.state()
	}

	// As opened, with nothing to undo.
	untouched(): boolean {
		return this.current === 0 && this.undoSteps.length === 0
	}

	private push(entry: UndoEntry<S>): void {
		this.undoSteps.push(entry)
		this.weight += entry.weight

		while (this.undoSteps.length > 1 && (this.undoSteps.length > HISTORY_LIMIT || this.weight > this.budget)) {
			const dropped = this.undoSteps.shift()

			this.weight -= dropped?.weight ?? 0
		}
	}
}
