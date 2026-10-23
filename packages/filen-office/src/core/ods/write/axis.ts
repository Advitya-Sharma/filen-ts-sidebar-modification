import type { AxisShift } from "../../xlsx/refs/axis"

// Where each row (or column) of a sheet came from in its source table: runs of consecutive positions now that were
// consecutive positions then, in order (a structure edit never reorders them). A position in no run is new (an
// insert). Immutable: an edit gives a new one, so its undo puts the old one back.
export class SourceAxis {
	// Per run: its first position now, its first position in the source, and how many it holds.
	private readonly targets: Int32Array
	private readonly sources: Int32Array
	private readonly lengths: Int32Array

	private constructor(targets: Int32Array, sources: Int32Array, lengths: Int32Array) {
		this.targets = targets
		this.sources = sources
		this.lengths = lengths
	}

	// Every position its own, up to `limit`.
	static identity(limit: number): SourceAxis {
		return new SourceAxis(Int32Array.of(0), Int32Array.of(0), Int32Array.of(limit))
	}

	get runs(): number {
		return this.targets.length
	}

	get identity(): boolean {
		return this.targets.length === 1 && this.targets[0] === 0 && this.sources[0] === 0
	}

	// The run holding target position `index`, -1 for none.
	private runAt(index: number): number {
		let low = 0
		let high = this.targets.length - 1

		while (low <= high) {
			const mid = (low + high) >>> 1

			if ((this.targets[mid] ?? 0) <= index) low = mid + 1
			else high = mid - 1
		}

		return high >= 0 && index < (this.targets[high] ?? 0) + (this.lengths[high] ?? 0) ? high : -1
	}

	// The source position of target position `index`, -1 for a new one.
	source(index: number): number {
		const run = this.runAt(index)

		return run < 0 ? -1 : (this.sources[run] ?? 0) + index - (this.targets[run] ?? 0)
	}

	// The target position of source position `index`, -1 when it was deleted.
	target(index: number): number {
		let low = 0
		let high = this.sources.length - 1

		while (low <= high) {
			const mid = (low + high) >>> 1

			if ((this.sources[mid] ?? 0) <= index) low = mid + 1
			else high = mid - 1
		}

		return high >= 0 && index < (this.sources[high] ?? 0) + (this.lengths[high] ?? 0)
			? (this.targets[high] ?? 0) + index - (this.sources[high] ?? 0)
			: -1
	}

	// How many source positions from `index` on keep their order next to each other now (to the end of their run),
	// 0 for a deleted one.
	runFrom(index: number): number {
		const target = this.target(index)

		if (target < 0) return 0

		const run = this.runAt(target)

		return (this.targets[run] ?? 0) + (this.lengths[run] ?? 0) - target
	}

	// The first source position at or after `index` that is kept, -1 for none (every one from there was deleted).
	nextKept(index: number): number {
		let low = 0
		let high = this.sources.length - 1

		// The last run starting at or before `index`.
		while (low <= high) {
			const mid = (low + high) >>> 1

			if ((this.sources[mid] ?? 0) <= index) low = mid + 1
			else high = mid - 1
		}

		if (high >= 0 && index < (this.sources[high] ?? 0) + (this.lengths[high] ?? 0)) return index

		return high + 1 < this.sources.length ? (this.sources[high + 1] ?? -1) : -1
	}

	// The last source position at or before `index` that is kept, -1 for none.
	previousKept(index: number): number {
		let low = 0
		let high = this.sources.length - 1

		while (low <= high) {
			const mid = (low + high) >>> 1

			if ((this.sources[mid] ?? 0) <= index) low = mid + 1
			else high = mid - 1
		}

		if (high < 0) return -1

		return Math.min(index, (this.sources[high] ?? 0) + (this.lengths[high] ?? 0) - 1)
	}

	// Each run as [target, source, length].
	forEach(visit: (target: number, source: number, length: number) => void): void {
		for (let k = 0; k < this.targets.length; k++) visit(this.targets[k] ?? 0, this.sources[k] ?? 0, this.lengths[k] ?? 0)
	}

	// After rows or columns inserted or deleted on this axis (a shift on the other axis, or confined to a band of
	// it, moves no whole row or column: the same axis).
	shifted(shift: AxisShift, axis: "rows" | "cols", limit: number): SourceAxis {
		if (shift.axis !== axis || shift.band !== undefined) return this

		const targets: number[] = []
		const sources: number[] = []
		const lengths: number[] = []
		const add = (target: number, source: number, length: number): void => {
			const clipped = Math.min(length, limit - target)

			if (clipped <= 0) return

			const last = targets.length - 1

			if (
				last >= 0 &&
				(targets[last] ?? 0) + (lengths[last] ?? 0) === target &&
				(sources[last] ?? 0) + (lengths[last] ?? 0) === source
			) {
				lengths[last] = (lengths[last] ?? 0) + clipped

				return
			}

			targets.push(target)
			sources.push(source)
			lengths.push(clipped)
		}
		const { at, count } = shift

		this.forEach((target, source, length) => {
			const end = target + length

			if (shift.type === "insert") {
				if (end <= at) add(target, source, length)
				else if (target >= at) add(target + count, source, length)
				else {
					add(target, source, at - target)
					add(at + count, source + at - target, end - at)
				}

				return
			}

			const cut = at + count

			if (end <= at) add(target, source, length)
			else if (target >= cut) add(target - count, source, length)
			else {
				if (target < at) add(target, source, at - target)
				if (end > cut) add(at, source + cut - target, end - cut)
			}
		})

		return new SourceAxis(Int32Array.from(targets), Int32Array.from(sources), Int32Array.from(lengths))
	}
}
