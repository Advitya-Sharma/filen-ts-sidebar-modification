import { Fragment, type Node, type Schema } from "prosemirror-model"
import type { Transaction } from "prosemirror-state"
import { Step, StepMap, StepResult, type Mappable } from "prosemirror-transform"
import { asBuilt, builtBlock } from "./origin"

// A node whose attributes a SetAttrsStep sets: its position (where it starts) and the attributes changed,
// with their new values (the others stay).
export interface AttrsEntry {
	readonly pos: number
	readonly attrs: Readonly<Record<string, unknown>>
	// An inverse's: the attributes object the node had, set again when the node's other attributes are still
	// its (an undo gives the node back the very object, which asBuilt and identity caches recognise).
	readonly previous?: Readonly<Record<string, unknown>>
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
			const next = this.rebuild(doc, 0, cursor, undefined)

			if (cursor.missed()) return StepResult.fail("No node at a position of SetAttrsStep")

			return StepResult.ok(next)
		} catch (error) {
			return StepResult.fail(error instanceof Error ? error.message : String(error))
		}
	}

	// `node` (its content starting at `start`) with its descendants' attributes set, a block the builder made
	// when it is that again (asBuilt; `built` is what it made in `node`'s place).
	private rebuild(node: Node, start: number, cursor: Cursor, built: Node | undefined): Node {
		const out: Node[] = []
		let changed = false
		const paired = built !== undefined && built.childCount === node.childCount ? built : undefined

		node.forEach((child, offset, index) => {
			const pos = start + offset

			if (!cursor.within(pos, pos + child.nodeSize)) {
				out.push(child)

				return
			}

			const entry = cursor.at(pos)
			const attrs = entry === undefined ? child.attrs : (restored(child.attrs, entry) ?? withAttrs(child.attrs, entry.attrs))
			const own = paired === undefined ? (builtBlock(attrs) ?? builtBlock(child.attrs)) : paired.child(index)
			const content = cursor.within(pos + 1, pos + child.nodeSize) ? this.rebuild(child, pos + 1, cursor, own).content : child.content
			// An inverse giving back the built node's attributes object (an undo) rebuilds from that node, which keeps
			// it: NodeType.create would copy it, and asBuilt could no longer find the built block.
			const rebuilt =
				entry === undefined
					? content === child.content
						? child
						: child.copy(content)
					: own?.attrs === attrs
						? own.mark(child.marks).copy(content)
						: child.type.create(attrs, content, child.marks)
			const next = rebuilt === child ? child : asBuilt(rebuilt, own, attrs)

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

			entries.push({ pos, attrs: old, previous: node.attrs })
		})

		return new SetAttrsStep(entries)
	}

	// As AttrStep: a node deleted after its position is dropped.
	override map(mapping: Mappable): Step | null {
		const entries: AttrsEntry[] = []

		for (const entry of this.entries) {
			const result = mapping.mapResult(entry.pos, 1)

			if (!result.deletedAfter) entries.push({ ...entry, pos: result.pos })
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

// The attributes object an inverse's entry gives back: the node's previous one, when the node's other
// attributes are still that object's (undefined when they are not, or the entry has none).
function restored(attrs: Readonly<Record<string, unknown>>, entry: AttrsEntry): Readonly<Record<string, unknown>> | undefined {
	const previous = entry.previous

	if (previous === undefined) return undefined

	for (const name in entry.attrs) if (entry.attrs[name] !== previous[name]) return undefined
	for (const name in attrs) if (!Object.hasOwn(entry.attrs, name) && attrs[name] !== previous[name]) return undefined
	for (const name in previous) if (!Object.hasOwn(attrs, name)) return undefined

	return previous
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
