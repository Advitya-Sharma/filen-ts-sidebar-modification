import { Mark, type Node } from "prosemirror-model"
import type { BytePieces } from "../../opc/index"
import type { SliceClock } from "../../slicing"
import type { DocxEditContext } from "../edit/context"
import type { StoryDoc } from "../edit/build"
import { NO_SOURCE, docxSchema, type MarkerList } from "../edit/schema"
import type { StoryPart } from "../model/index"
import { closeTag, copyNode, Names, openTag, Out, spliceChildren } from "./markup"
import { MARKER_ITEMS, writeListMarker } from "./markers"
import { copySource, type NestedStories } from "./nested"
import { writeNewStory, type NewStoryContainer } from "./newStory"
import { writeParagraph } from "./paragraph"
import { finalSectionMarkup } from "./sections"
import type { SaveParts } from "./parts/index"
import { writeSdt } from "./sdt"
import { R_NAMESPACE, W14_NAMESPACE, W_NAMESPACE, cellWithin, inCell, refuseBlockMarks, src } from "./source"
import { writeTable } from "./tables"
import { DocxWriteError } from "./values"

export { DocxWriteError, type WriteFailure } from "./values"
export { R_NAMESPACE, W_NAMESPACE, cellWithin, inCell, sectionWritten } from "./source"

// What a whole save shares across stories: new ids and the session-level registries resolved to file values.
export interface WriteEnvironment {
	readonly context: DocxEditContext
	// A new w:id for a revision wrapper written a second time.
	nextRevisionId(): number
	// The w:numId a command's new list (negative id) became.
	numId(id: number): number
	// A relationship id for a new external hyperlink from `part`.
	hyperlink(part: string, target: string): string
	// The parts of the save being written, for writers that add to them or reference new ones (pictures,
	// notes); absent where nothing new may be (gate 1, the layout of edited blocks).
	readonly parts?: SaveParts
	// Stories rewritten inside source bytes this story copies (see nested.ts).
	readonly nested?: NestedStories
}

export interface StoryWriteOptions {
	// Gate 1: write this block through the edit path even though it is unchanged (runs rebuilt from their
	// bases, containers from their parts).
	readonly force?: Node
	// A story the source does not have (its item is -1): the container it is written in.
	readonly container?: NewStoryContainer
	// Prefixes of namespaces the save declares on the part's root (by URI), where the source's root lacks them.
	readonly prefixes?: ReadonlyMap<string, string>
}

const nodes = docxSchema.nodes

// A story as written: the container element's bytes, the spans copied as whole source nodes (see Out),
// the paragraphs rebuilt through the edit path and the property elements of containers written as copies,
// each as a byte range of `bytes`.
export interface StoryOutput {
	// Joined (writeStory), or the pieces a save assembles a part from.
	readonly bytes: Uint8Array | BytePieces
	readonly trusted: readonly number[]
	readonly paragraphs: readonly WrittenParagraph[]
	readonly containers: readonly WrittenContainer[]
}

export interface WrittenParagraph {
	readonly node: Node
	readonly from: number
	readonly to: number
}

// Property elements (source nodes) of a container copy, written without the children named in `strip`.
export interface WrittenContainer {
	readonly properties: readonly number[]
	readonly strip: readonly string[]
	readonly from: number
	readonly to: number
}

// A story written but not yet joined into one byte array: the caller can compare it with its source cheaply
// (Out.equals) and assemble it with the rest of its part (Out.pieces).
export interface StoryDraft {
	readonly out: Out
	readonly paragraphs: readonly WrittenParagraph[]
	readonly containers: readonly WrittenContainer[]
}

const NO_SIBLINGS: readonly number[] = []

// Where source item `item` is among a container's children (items ascend in source order), or -1.
function indexOf(siblings: readonly number[], item: number): number {
	let low = 0
	let high = siblings.length - 1

	while (low <= high) {
		const middle = (low + high) >> 1
		const at = siblings[middle] ?? 0

		if (at === item) return middle
		if (at < item) low = middle + 1
		else high = middle - 1
	}

	return -1
}

// Runs a write to its end (without a clock nothing yields).
function complete<T>(steps: Generator<void, T>): T {
	for (;;) {
		const step = steps.next()

		if (step.done === true) return step.value
	}
}

async function drive<T>(steps: Generator<void, T>, clock: SliceClock): Promise<T> {
	for (;;) {
		const step = steps.next()

		if (step.done === true) return step.value

		await clock.pause()
	}
}

// The first block of each source item as the story was built, per story (built once: the story is never
// changed).
const originalsCache = new WeakMap<StoryDoc, Map<number, Node>>()

function originalBlocks(story: StoryDoc): Map<number, Node> {
	return originalsCache.get(story) ?? complete(findOriginals(story, () => false))
}

function* findOriginals(story: StoryDoc, due: () => boolean): Generator<void, Map<number, Node>> {
	const found = new Map<number, Node>()
	const visit = (child: Node): boolean => {
		if (child.isInline) return false

		const own = src(child)

		if (own !== NO_SOURCE && !found.has(own)) found.set(own, child)

		return !child.isTextblock
	}
	const doc = story.doc

	for (let index = 0; index < doc.childCount; index++) {
		const child = doc.child(index)

		if (visit(child)) child.descendants(visit)

		if (due()) yield
	}

	originalsCache.set(story, found)

	return found
}

// Writes one story (body, header, footer, note, comment) of an editing document back into its part:
// everything that can still be copied as source bytes is, the rest goes through the edit path described in
// StoryDoc. The output keeps the story container element and replaces only its content.
export class StoryWriter {
	readonly story: StoryDoc
	readonly part: StoryPart
	readonly names: Names
	private readonly writing: StoryWriting

	constructor(env: WriteEnvironment, story: StoryDoc, options: StoryWriteOptions = {}) {
		this.writing = new StoryWriting(env, story, options)
		this.story = story
		this.part = story.part
		this.names = this.writing.names
	}

	// The story container element (w:body, w:hdr, w:footnote...) written from `doc`, with what gate 2 needs
	// to check it without parsing it whole.
	writeStory(doc: Node): StoryOutput & { readonly bytes: Uint8Array } {
		const draft = complete(this.writing.draft(doc))

		return { bytes: draft.out.finish(), trusted: draft.out.trusted, paragraphs: draft.paragraphs, containers: draft.containers }
	}

	// writeStory without joining the output (see StoryDraft).
	draftStorySync(doc: Node): StoryDraft {
		return complete(this.writing.draft(doc))
	}

	// writeStory in slices: the clock's pause runs between them (a rejecting pause abandons the write). The
	// output is not joined yet (see StoryDraft).
	async draftStory(doc: Node, clock: SliceClock): Promise<StoryDraft> {
		this.writing.clock = clock

		try {
			return await drive(this.writing.draft(doc), clock)
		} finally {
			this.writing.clock = undefined
		}
	}

	// One block through the edit path, for gate 1: runs rebuilt from their bases, properties spliced,
	// containers rebuilt from their parts (a table's paragraphs are copied, each is proved on its own).
	// `copy`: an earlier block of the story holds the same source (a pasted copy), as writeStory finds it.
	// `cell`: whether the block is in a table cell (by default, where its source is).
	writeBlockAlone(node: Node, copy = false, cell?: boolean): Out {
		return this.writing.writeBlockAlone(node, copy, cell)
	}

	// Whether a paragraph is again the one the story was built with (an undo): its source bytes stand for it.
	isSource(node: Node): boolean {
		return (
			node.type === docxSchema.nodes.paragraph &&
			src(node) !== NO_SOURCE &&
			node.marks.length === 0 &&
			(this.story.pristine.has(node) || sameParagraph(this.original(node), node))
		)
	}

	private original(node: Node): Node | undefined {
		return originalBlocks(this.story).get(src(node))
	}
}

// The state of one story's write, shared by the modules that write its parts (paragraph.ts, tables.ts,
// sdt.ts, wrappers.ts, markers.ts, atoms/): the walker and the containers live here.
export class StoryWriting {
	readonly story: StoryDoc
	readonly part: StoryPart
	readonly names: Names
	readonly env: WriteEnvironment
	readonly bytes: Uint8Array
	force: Node | undefined
	// A story the source does not have: the container it is written in.
	readonly container: NewStoryContainer | undefined
	// Table cells the block being written is in.
	cells = 0
	private readonly owners = new Map<number, Node>()
	// Source items whose owning node object stands at more than one place, those already written, and the
	// depth of later occurrences being written.
	private readonly shared = new Set<number>()
	private readonly claimed = new Set<number>()
	// Some source item stands on two different nodes (a copy).
	private duplicated = false
	private repeating = 0
	// Whether a block can be written as its source bytes, outside table cells and inside one.
	private readonly copyable = new Map<Node, boolean>()
	private readonly copyableInCell = new Map<Node, boolean>()
	private readonly siblings = new Map<number, number[]>()
	// Set while writing in slices.
	clock: SliceClock | undefined
	readonly copiedRuns = new Set<number>()
	readonly startedRuns = new Set<number>()
	readonly openedWrappers = new Set<number>()
	readonly paragraphs: WrittenParagraph[] = []
	// The w14:textId attribute's qualified name, and the w14 prefix the part binds ("" when it binds none).
	readonly textId: ReadonlySet<string> | undefined
	readonly w14: string
	private readonly containers: WrittenContainer[] = []
	// Bases whose w:rPrChange has been written (copied runs included).
	readonly writtenChanges = new Set<number>()

	constructor(env: WriteEnvironment, story: StoryDoc, options: StoryWriteOptions) {
		this.env = env
		this.story = story
		this.part = story.part
		this.bytes = story.part.document.bytes
		this.force = options.force
		this.container = options.container

		const reader = story.part.reader
		const root = story.part.document.documentElement()

		const w14 = reader.namespaces.prefixFor(root, W14_NAMESPACE)

		this.w14 = w14 ?? options.prefixes?.get(W14_NAMESPACE) ?? ""
		this.textId = w14 === undefined || w14 === "" ? undefined : new Set([`${w14}:textId`])
		this.names = new Names(reader.namespaces.prefixFor(root, W_NAMESPACE) ?? "", reader.namespaces.prefixFor(root, R_NAMESPACE) ?? "")
	}

	*draft(doc: Node): Generator<void, StoryDraft> {
		this.owners.clear()
		this.copyable.clear()
		this.copyableInCell.clear()
		this.shared.clear()
		this.claimed.clear()

		if (!originalsCache.has(this.story)) yield* findOriginals(this.story, () => this.due())

		yield* this.findOwners(doc)
		this.paragraphs.length = 0
		this.containers.length = 0
		this.cells = 0

		// In slices, what can be copied is known before writing starts, children first: deciding it for a
		// large table while writing would be one long step. Nodes standing at two places are decided as they
		// are written (a claim changes the answer).
		if (this.clock !== undefined && this.shared.size === 0) yield* this.decideCopies(doc)

		const out = new Out()

		if (this.container !== undefined) {
			yield* writeNewStory(this, out, doc, this.container)

			return { out, paragraphs: [...this.paragraphs], containers: [...this.containers] }
		}

		const document = this.part.document
		const container = this.part.node(this.story.story.item)

		const sectPr = doc.attrs["sectPr"] as number
		// The body's own section properties as an edit changed them (sections.ts), else the source's.
		const final = finalSectionMarkup(this, doc)

		// A story without blocks, its empty paragraph untouched: as its source.
		if (final === undefined && untouchedEmpty(this.story, doc)) {
			this.copyWhole(out, container)

			return { out, paragraphs: [], containers: [] }
		}

		openTag(out, document, container)

		const tail = yield* this.writeChildren(out, doc, this.story.story.item, (child, at) => this.writeBlock(at, child), {
			from: document.openEnd(container),
			to: document.contentEnd(container),
			tail: sectPr === NO_SOURCE || final !== undefined ? [] : [sectPr]
		})

		if (final !== undefined) out.text(final)
		else if (!tail && sectPr !== NO_SOURCE) this.copyWhole(out, this.part.node(sectPr))

		closeTag(out, document, container)

		return { out, paragraphs: [...this.paragraphs], containers: [...this.containers] }
	}

	writeBlockAlone(node: Node, copy: boolean, cell: boolean | undefined): Out {
		this.force = node
		// Each block is written alone once (a proof keeps its writer): nothing of one is kept for the next.
		this.siblings.clear()
		this.owners.clear()
		this.copyable.clear()
		this.copyableInCell.clear()
		this.copiedRuns.clear()
		this.startedRuns.clear()
		this.openedWrappers.clear()
		this.writtenChanges.clear()
		this.owners.set(src(node), node)
		this.shared.clear()
		this.claimed.clear()
		complete(this.findOwners(node))
		this.cells = (cell ?? (node.type === nodes.paragraph && src(node) !== NO_SOURCE && inCell(this.part, src(node)))) ? 1 : 0

		// The earlier block owns the sources and wrote its runs first: this one is written as a copy.
		if (copy) {
			const other = copyOwner()

			for (const own of this.owners.keys()) this.owners.set(own, other)

			node.descendants(child => {
				const run = docxSchema.marks.run.isInSet(child.marks)

				if (run !== undefined && src(run) !== NO_SOURCE) this.copiedRuns.add(src(run))

				return true
			})
		}

		// Only a whole story's output reports these; a block alone would keep every node it wrote alive.
		this.paragraphs.length = 0
		this.containers.length = 0

		const out = new Out()

		complete(this.writeBlock(out, node))

		return out
	}

	// ---- identity ----------------------------------------------------------------------------------------

	// The first node of each source item owns it. ProseMirror shares unchanged nodes, so one node object can
	// stand at two places (an unsanitised copy): its first occurrence owns, the later ones are copies.
	private *findOwners(doc: Node): Generator<void, void> {
		const visit = (node: Node): boolean => {
			if (node.isInline) return false

			const own = src(node)

			if (own !== NO_SOURCE) {
				const owner = this.owners.get(own)

				if (owner === undefined) this.owners.set(own, node)
				else if (owner === node) this.shared.add(own)
				else this.duplicated = true
			}

			return !node.isTextblock && node.type !== nodes.locked_block
		}

		this.duplicated = false

		for (let index = 0; index < doc.childCount; index++) {
			const child = doc.child(index)

			if (visit(child)) child.descendants(visit)

			if (this.due()) yield
		}
	}

	// Whether the write in progress should give the host a turn (only in slices).
	due(): boolean {
		return this.clock !== undefined && this.clock.due
	}

	// canCopy for every block, children before their container (each container then looks only at its
	// children's answers), in the cell context writing uses. A container still the very node the story was
	// built with, where its source was, is decided at once (see unchangedContainer).
	private *decideCopies(doc: Node): Generator<void, void> {
		const stack: { node: Node; cell: boolean; index: number }[] = []

		for (let top = 0; top < doc.childCount; top++) {
			stack.push({ node: doc.child(top), cell: this.cells > 0, index: 0 })

			while (stack.length > 0) {
				const frame = stack[stack.length - 1] as { node: Node; cell: boolean; index: number }
				const node = frame.node

				if (frame.index === 0 && this.unchangedContainer(node, frame.cell)) {
					stack.pop()
					this.canCopy(node, frame.cell)

					continue
				}

				if (!node.isTextblock && node.type !== nodes.locked_block && frame.index < node.childCount) {
					stack.push({ node: node.child(frame.index++), cell: frame.cell || node.type === nodes.table_cell, index: 0 })

					continue
				}

				stack.pop()
				this.canCopy(node, frame.cell)

				if (this.due()) yield
			}
		}
	}

	// Whether `node` is the first occurrence of its source item, which it is then written as.
	owns(node: Node): boolean {
		if (this.repeating > 0) return false

		const own = src(node)

		return own !== NO_SOURCE && this.owners.get(own) === node && !this.claimed.has(own)
	}

	// After an occurrence of `node` was written: later occurrences of shared nodes in it are copies.
	private claim(node: Node): void {
		if (this.shared.size === 0) return

		const own = src(node)

		if (own !== NO_SOURCE && this.shared.has(own)) this.claimed.add(own)

		if (!node.isTextblock && !node.isLeaf) node.descendants(child => (this.claim(child), false))
	}

	// The block the story was built with for `node`'s source item.
	original(node: Node): Node | undefined {
		return originalBlocks(this.story).get(src(node))
	}

	// Whether a block can be written as its source bytes (see StoryDoc), in a table cell (`cell`) or not.
	private canCopy(node: Node, cell = this.cells > 0): boolean {
		if (this.repeating > 0 || this.claimed.has(src(node))) return false

		const memo = cell ? this.copyableInCell : this.copyable
		let result = memo.get(node)

		if (result !== undefined) return result

		// A block carrying node marks (what a feature recorded on it) is written through the edit path.
		if (node === this.force || !this.owns(node) || node.marks.length > 0) {
			result = false
		} else if (node.type === nodes.paragraph) {
			// An edited paragraph that is again what it was (an undo; markers before it are written apart) is
			// its bytes too, unless they hold a section break and it moved into or out of a table cell.
			const row = this.part.aux[src(node)] ?? -1

			result =
				(this.story.pristine.has(node) || sameParagraph(this.original(node), node)) &&
				((this.part.paragraphs.sectPr[row] ?? -1) < 0 || inCell(this.part, src(node)) === cell)
		} else if (node.type === nodes.locked_block) {
			// One whose section break moved into or out of a table cell is refused (see writeBlock).
			result = !this.lockedBreakMoved(node, cell)
		} else if (this.force !== undefined && this.force.type !== nodes.paragraph) {
			// A table under proof is rebuilt down to its paragraphs.
			result = false
		} else if (this.unchangedContainer(node, cell)) {
			result = true
		} else {
			const original = this.original(node)

			result =
				original !== undefined &&
				original.attrs === node.attrs &&
				original.childCount === node.childCount &&
				original.content.content.every(
					(child, index) =>
						src(child) === src(node.child(index)) && this.canCopy(node.child(index), cell || node.type === nodes.table_cell)
				)
		}

		memo.set(node, result)

		return result
	}

	// A container (table, row, cell, content control) that is still the very node the story was built with, so
	// is everything in it, written where its source was (in a table cell or not, as it was): its source bytes,
	// as canCopy would find child by child. Only while each source item stands at one place in the document
	// (then every node owns its source) and no block is forced through the edit path.
	private unchangedContainer(node: Node, cell: boolean): boolean {
		if (this.duplicated || this.shared.size > 0 || this.force !== undefined || this.repeating > 0) return false
		if (node.isInline || node.isTextblock || node.type === nodes.locked_block) return false

		const own = src(node)

		return own !== NO_SOURCE && this.owners.get(own) === node && this.original(node) === node && inCell(this.part, own) === cell
	}

	// Whether locked block `node`, in a table cell (`cell`) or not, has a section break that moved into or out
	// of one: a break in a cell of the block moves with it, one of the block itself only when the block did.
	private lockedBreakMoved(node: Node, cell: boolean): boolean {
		const sectPr = node.attrs["sectPr"] as number
		const own = src(node)

		return sectPr >= 0 && !cellWithin(this.part, sectPr, own) && inCell(this.part, own) !== cell
	}

	// ---- containers --------------------------------------------------------------------------------------

	private siblingsOf(container: number): number[] {
		let siblings = this.siblings.get(container)

		if (siblings === undefined) {
			siblings = []

			for (let item = container + 1; item < this.part.end(container); item = this.part.end(item)) siblings.push(item)

			this.siblings.set(container, siblings)
		}

		return siblings
	}

	// The children of a container: blocks that can be copied keep the bytes around them (whitespace,
	// comments, proofErr) when their source neighbours are also their neighbours here. `region` is the
	// container's content in the source and what must follow the last child there (the body's final
	// section properties); when everything after the last child is as it was, it is copied whole and true
	// returned.
	*writeChildren(
		out: Out,
		container: Node,
		sourceContainer: number,
		write: (child: Node, out: Out) => Generator<void, void>,
		region?: { readonly from: number; readonly to: number; readonly tail: readonly number[] }
	): Generator<void, boolean> {
		const siblings = sourceContainer === NO_SOURCE ? NO_SIBLINGS : this.siblingsOf(sourceContainer)
		let previous = -1

		for (let at = 0; at < container.childCount; at++) {
			const child = container.child(at)
			const before = (child.attrs["before"] as MarkerList | undefined) ?? []

			if (this.due()) yield

			if (this.canCopy(child)) {
				const own = src(child)
				const index = indexOf(siblings, own)
				const leading = index < 0 ? [] : this.leadingMarkers(siblings, index)
				const together = leading.length === before.length && leading.every((item, k) => before[k]?.src === item)
				const first = together && leading.length > 0 ? (leading[0] ?? own) : own

				if (!together) for (const marker of before) writeListMarker(this, out, marker)

				const firstIndex = indexOf(siblings, first)

				if (previous >= 0 && firstIndex > 0 && siblings[firstIndex - 1] === previous) {
					this.copySource(out, this.part.byteEnd(previous), this.part.start(first), true)
				} else if (previous === -1 && firstIndex === 0 && region !== undefined) {
					this.copySource(out, region.from, this.part.start(first), true)
				}

				this.copySource(out, this.part.start(first), this.part.byteEnd(own), true)
				this.claim(child)
				previous = own

				continue
			}

			for (const marker of before) writeListMarker(this, out, marker)

			// A later occurrence of a shared node is written as a copy, all the way down.
			const repeat = this.claimed.has(src(child))

			if (repeat) this.repeating++

			yield* write(child, out)

			if (repeat) this.repeating--
			else this.claim(child)

			previous = -2
		}

		const end = (container.attrs["end"] as MarkerList | undefined) ?? []

		if (region !== undefined && previous >= 0) {
			const rest = siblings.slice(indexOf(siblings, previous) + 1)
			const expected = [...end.map(marker => marker.src), ...region.tail]

			if (rest.length === expected.length && rest.every((item, k) => item === expected[k])) {
				this.copySource(out, this.part.byteEnd(previous), region.to, true)

				return true
			}
		}

		for (const marker of end) writeListMarker(this, out, marker)

		return false
	}

	private leadingMarkers(siblings: readonly number[], index: number): number[] {
		let start = index

		while (start > 0 && MARKER_ITEMS.has(this.part.kinds[siblings[start - 1] ?? -1] ?? 0)) start--

		return siblings.slice(start, index)
	}

	*writeBlock(out: Out, node: Node): Generator<void, void> {
		switch (node.type) {
			case nodes.paragraph:
				writeParagraph(this, out, node)

				return
			case nodes.table:
				refuseBlockMarks(node)
				yield* writeTable(this, out, node)

				return
			case nodes.sdt:
				refuseBlockMarks(node)
				yield* writeSdt(this, out, node)

				return
			case nodes.locked_block:
				refuseBlockMarks(node)

				if (!this.owns(node)) throw new DocxWriteError("missingSource", "A locked block appears twice")

				// Its bytes are all it can be written as: a section break in them must stay in or out of a table
				// cell as it was.
				if (this.lockedBreakMoved(node, this.cells > 0))
					throw new DocxWriteError("unsupported", "A locked section break moved into or out of a table cell")

				this.copyWhole(out, this.part.node(src(node)))

				return
			default:
				throw new DocxWriteError("unsupported", `Unexpected block ${node.type.name}`)
		}
	}

	// Source bytes [from, to) of the part: whole nodes copied as they are when `trusted` (see Out.copy), or
	// part of new markup; stories rewritten inside them are spliced in (env.nested).
	copySource(out: Out, from: number, to: number, trusted: boolean): void {
		copySource(out, this.part, from, to, trusted, this.env.nested)
	}

	// Source node `node` as part of new markup (a run, an atom), with stories rewritten inside it spliced in.
	copyNode(out: Out, node: number): void {
		if (this.env.nested === undefined) copyNode(out, this.part.document, node)
		else this.copySource(out, this.part.document.start(node), this.part.document.end(node), false)
	}

	copyWhole(out: Out, node: number): void {
		this.copySource(out, this.part.document.start(node), this.part.document.end(node), true)
	}

	// A container's property elements for a copy of it: everything but what carries an id of its own.
	copyProperties(out: Out, properties: readonly number[], strip: readonly string[]): void {
		const changes = new Map(strip.map(local => [local, ""]))
		const sorted = properties.filter(property => property >= 0).sort((a, b) => a - b)
		const from = out.size

		for (const property of sorted) spliceChildren(out, this.part.reader, property, changes, [])

		this.containers.push({ properties: sorted, strip, from, to: out.size })
	}

	// Container attributes no command changes yet: the writer has no way to express a change to them.
	checkUnchanged(node: Node, names: readonly string[], message: string): void {
		const original = this.original(node)

		if (original !== undefined && names.some(name => original.attrs[name] !== node.attrs[name]))
			throw new DocxWriteError("unsupported", message)
	}
}

function sameParagraph(original: Node | undefined, node: Node): boolean {
	if (original === undefined) return false
	if (original.eq(node)) return true
	if (!original.content.eq(node.content) || !Mark.sameSet(original.marks, node.marks)) return false

	for (const name of Object.keys(original.attrs)) {
		if (name !== "before" && original.attrs[name] !== node.attrs[name]) return false
	}

	return true
}

let copyOwnerNode: Node | undefined

// Stands for the block owning sources a copy is written without.
function copyOwner(): Node {
	copyOwnerNode ??= docxSchema.nodes.paragraph.create()

	return copyOwnerNode
}

// Whether `doc` is a story the source has no block in, holding only the builder's empty paragraph untouched (see
// StoryDoc.pristine) and nothing an edit added around it.
export function untouchedEmpty(story: StoryDoc, doc: Node): boolean {
	const only = doc.childCount === 1 ? doc.firstChild : null

	return (
		only !== null &&
		src(only) === NO_SOURCE &&
		story.pristine.has(only) &&
		((doc.attrs["end"] as MarkerList | undefined)?.length ?? 0) === 0
	)
}
