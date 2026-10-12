import { Fragment, type Node, type Schema } from "prosemirror-model"
import type { Transaction } from "prosemirror-state"
import { Step, StepMap, StepResult, type Mappable } from "prosemirror-transform"

// A node whose attributes a SetAttrsStep sets: its position (where it starts) and the attributes changed,
// with their new values (the others stay).
export interface AttrsEntry {
	readonly pos: number
	readonly attrs: Readonly<Record<string, unknown>>
}

// Sets attributes of many nodes as one step. A step per node (setNodeAttribute) copies the block list around
// the node for every node, and a transaction keeps every intermediate document: a paragraph command over a
// long document that way takes seconds and gigabytes. This step rebuilds the blocks around its nodes once,
// keeping every other block (identity included). It moves no position (an empty map, like AttrStep); its
// inverse holds the values each node had, so an undo restores exactly.
export class SetAttrsStep extends Step {
	// In document order, one per node.
	readonly entries: readonly AttrsEntry[]

	constructor(entries: readonly AttrsEntry[]) {
		super()

		this.entries = entries
	}

	// The span of the nodes' start tokens: [first node's start, last node's start + 1).
	get from(): number {
		return this.entries[0]?.pos ?? 0
	}

	get to(): number {
		return (this.entries[this.entries.length - 1]?.pos ?? -1) + 1
	}

	// Calls `f` with each node of the step in `doc` (the document it applies to: the nodes as they were) in
	// document order, in one walk.
	eachNode(doc: Node, f: (node: Node, pos: number, attrs: Readonly<Record<string, unknown>>) => void): void {
		const cursor = new Cursor(this.entries)
		const visit = (node: Node, start: number): void => {
			node.forEach((child, offset) => {
				const pos = start + offset

				if (!cursor.within(pos, pos + child.nodeSize)) return

				const entry = cursor.at(pos)

				if (entry !== undefined) f(child, pos, entry.attrs)
				if (cursor.within(pos + 1, pos + child.nodeSize)) visit(child, pos + 1)
			})
		}

		visit(doc, 0)

		if (cursor.missed()) throw new RangeError("No node at a position of SetAttrsStep")
	}

	override apply(doc: Node): StepResult {
		try {
			const cursor = new Cursor(this.entries)
			const next = this.rebuild(doc, 0, cursor)

			if (cursor.missed()) return StepResult.fail("No node at a position of SetAttrsStep")

			return StepResult.ok(next)
		} catch (error) {
			return StepResult.fail(error instanceof Error ? error.message : String(error))
		}
	}

	// `node` (its content starting at `start`) with its descendants' attributes set.
	private rebuild(node: Node, start: number, cursor: Cursor): Node {
		const out: Node[] = []
		let changed = false

		node.forEach((child, offset) => {
			const pos = start + offset

			if (!cursor.within(pos, pos + child.nodeSize)) {
				out.push(child)

				return
			}

			const entry = cursor.at(pos)
			const content = cursor.within(pos + 1, pos + child.nodeSize) ? this.rebuild(child, pos + 1, cursor).content : child.content
			const next =
				entry === undefined
					? content === child.content
						? child
						: child.copy(content)
					: child.type.create(withAttrs(child.attrs, entry.attrs), content, child.marks)

			if (next !== child) changed = true

			out.push(next)
		})

		return changed ? node.copy(Fragment.fromArray(out)) : node
	}

	override getMap(): StepMap {
		return StepMap.empty
	}

	override invert(doc: Node): Step {
		const entries: AttrsEntry[] = []

		this.eachNode(doc, (node, pos, attrs) => {
			const old: Record<string, unknown> = {}

			for (const name of Object.keys(attrs)) old[name] = node.attrs[name]

			entries.push({ pos, attrs: old })
		})

		return new SetAttrsStep(entries)
	}

	// As AttrStep: a node deleted after its position is dropped.
	override map(mapping: Mappable): Step | null {
		const entries: AttrsEntry[] = []

		for (const entry of this.entries) {
			const result = mapping.mapResult(entry.pos, 1)

			if (!result.deletedAfter) entries.push({ pos: result.pos, attrs: entry.attrs })
		}

		return entries.length === 0 ? null : new SetAttrsStep(entries)
	}

	override merge(): Step | null {
		return null
	}

	override toJSON(): { stepType: string; entries: [number, Readonly<Record<string, unknown>>][] } {
		return { stepType: "docxSetAttrs", entries: this.entries.map(entry => [entry.pos, entry.attrs]) }
	}

	static override fromJSON(_schema: Schema, json: { entries: [number, Record<string, unknown>][] }): SetAttrsStep {
		if (!Array.isArray(json.entries) || json.entries.some(entry => typeof entry[0] !== "number" || typeof entry[1] !== "object"))
			throw new RangeError("Invalid input for SetAttrsStep.fromJSON")

		return new SetAttrsStep(json.entries.map(([pos, attrs]) => ({ pos, attrs })))
	}
}

Step.jsonID("docxSetAttrs", SetAttrsStep)

// The nodes and attributes of a SetAttrsStep gathered (a value a node already has is left out), made into the
// step at the end.
export class AttrChanges {
	private readonly entries = new Map<number, Record<string, unknown>>()

	// Sets `attrs` on `node` at `pos` (values compared by identity, as setNodeAttribute's callers did).
	set(node: Node, pos: number, attrs: Readonly<Record<string, unknown>>): void {
		let entry = this.entries.get(pos)

		for (const [name, value] of Object.entries(attrs)) {
			if (node.attrs[name] === value && (entry === undefined || !(name in entry))) continue

			if (entry === undefined) {
				entry = {}
				this.entries.set(pos, entry)
			}

			entry[name] = value
		}
	}

	get empty(): boolean {
		return this.entries.size === 0
	}

	// Adds the step to `tr` (none without changes).
	step(tr: Transaction): void {
		if (this.entries.size === 0) return

		const entries = [...this.entries].map(([pos, attrs]) => ({ pos, attrs }))

		entries.sort((a, b) => a.pos - b.pos)
		tr.step(new SetAttrsStep(entries))
	}
}

// `attrs` with `changes` laid over (copied as AttrStep does: a spread of attribute objects is several times
// slower).
function withAttrs(attrs: Readonly<Record<string, unknown>>, changes: Readonly<Record<string, unknown>>): Record<string, unknown> {
	const out: Record<string, unknown> = Object.create(null) as Record<string, unknown>

	for (const name in attrs) out[name] = attrs[name]
	for (const name in changes) out[name] = changes[name]

	return out
}

// The entries a walk in document order reaches next.
class Cursor {
	private index = 0
	private readonly entries: readonly AttrsEntry[]

	constructor(entries: readonly AttrsEntry[]) {
		this.entries = entries
	}

	// Whether an entry lies in [pos, end): a node there or inside.
	within(pos: number, end: number): boolean {
		const entry = this.entries[this.index]

		return entry !== undefined && entry.pos >= pos && entry.pos < end
	}

	// The entry of the node at `pos`, taken.
	at(pos: number): AttrsEntry | undefined {
		const entry = this.entries[this.index]

		if (entry?.pos !== pos) return undefined

		this.index++

		return entry
	}

	// Whether an entry was passed without a node at its position.
	missed(): boolean {
		return this.index < this.entries.length
	}
}
