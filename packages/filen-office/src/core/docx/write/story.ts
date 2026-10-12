import { Mark, type MarkType, type Node } from "prosemirror-model"
import { decodeUtf8, escapeAttribute, escapeText } from "../../xml/index"
import type { DocxEditContext } from "../edit/context"
import type { StoryDoc } from "../edit/build"
import type { RunBase } from "../edit/formatting"
import { NO_SOURCE, WRAPPER_MARKS, docxSchema, endsWithTable, inlineLevel, type IndentAttrs, type MarkerList } from "../edit/schema"
import { Item, type StoryPart } from "../model/index"
import type { BytePieces } from "../../opc/index"
import type { SliceClock } from "../../slicing"
import { closeTag, copyNode, freshContainer, Names, openTag, openTagWithout, Out, spliceChildren, startTagWith } from "./markup"
import { JUSTIFICATION, NUMBERING_PROPERTIES, PARAGRAPH_MARK_PROPERTIES, PARAGRAPH_PROPERTIES } from "./order.generated"
import { sameFormatting, writeRunProperties } from "./runs"
import { DocxWriteError, indent as checkIndent, integer, name as checkName, oneOf } from "./values"

export const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
export const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"

export { DocxWriteError, type WriteFailure } from "./values"

// What a whole save shares across stories: new ids and the session-level registries resolved to file values.
export interface WriteEnvironment {
	readonly context: DocxEditContext
	// A new w:id for a revision wrapper written a second time.
	nextRevisionId(): number
	// The w:numId a command's new list (negative id) became.
	numId(id: number): number
	// A relationship id for a new external hyperlink from `part`.
	hyperlink(part: string, target: string): string
}

export interface StoryWriteOptions {
	// Gate 1: write this block through the edit path even though it is unchanged (runs rebuilt from their
	// bases, containers from their parts).
	readonly force?: Node
}

const nodes = docxSchema.nodes
const marks = docxSchema.marks
const wrapperTypes = new Set(WRAPPER_MARKS.map(name => marks[name]))
const WRAPPER_ITEMS = new Set<number>([
	Item.Hyperlink,
	Item.SmartTag,
	Item.CustomXml,
	Item.Sdt,
	Item.Insertion,
	Item.Deletion,
	Item.MoveFrom,
	Item.MoveTo,
	Item.Bidi
])
const MARKER_ITEMS = new Set<number>([
	Item.BookmarkStart,
	Item.BookmarkEnd,
	Item.CommentRangeStart,
	Item.CommentRangeEnd,
	Item.MoveFromRangeStart,
	Item.MoveFromRangeEnd,
	Item.MoveToRangeStart,
	Item.MoveToRangeEnd,
	Item.PermissionStart,
	Item.PermissionEnd,
	Item.CustomXmlRangeMarker
])
// Word marks text xml:space="preserve" exactly when it starts or ends with whitespace (inner runs of spaces
// it writes as they are, and reads them so).
const PRESERVE = /^[ \t\r\n]|[ \t\r\n]$/
// The prefix a hyperlink declares for the relationships namespace when the part binds none (a name no
// WordprocessingML part uses for anything else).
const LOCAL_R = "filenr"
// Children of container properties a copy leaves out: tracked changes, whose ids are the original's.
const TRACKED_TABLE = ["tblPrChange", "tblGridChange"]
const TRACKED_ROW = ["ins", "del", "trPrChange", "tblPrExChange"]
const TRACKED_CELL = ["cellIns", "cellDel", "cellMerge", "tcPrChange"]
const MARK_REVISIONS = new Map([
	["ins", ""],
	["del", ""],
	["moveFrom", ""],
	["moveTo", ""],
	["rPrChange", ""]
])

// Whether source item `item` lies inside a table cell.
export function inCell(part: StoryPart, item: number): boolean {
	for (let at = part.parent(item); at >= 0; at = part.parent(at)) if (part.kinds[at] === Item.Cell) return true

	return false
}

// Whether source item `item` lies in a table cell inside source item `container`.
export function cellWithin(part: StoryPart, item: number, container: number): boolean {
	for (let at = part.parent(item); at >= 0 && at !== container; at = part.parent(at)) if (part.kinds[at] === Item.Cell) return true

	return false
}

// Whether a paragraph ending with the section break of source paragraph `sectPr` writes it, in a table cell
// or not: a break written in a cell is no section break (the model ignores it), so one is written only
// where its source paragraph had it, in a cell or out of one.
export function sectionWritten(part: StoryPart, sectPr: number, cell: boolean): boolean {
	return sectPr !== NO_SOURCE && inCell(part, sectPr) === cell
}

function src(node: Node | Mark): number {
	const value: unknown = node.attrs["src"]

	return typeof value === "number" ? value : NO_SOURCE
}

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

// A run of a paragraph's inline content written as one unit: consecutive content of one source run (or of
// no run) under the same wrappers, or one paragraph-level atom.
interface Unit {
	readonly nodes: Node[]
	readonly run: Mark | undefined
	readonly wrappers: readonly Mark[]
	readonly atom: boolean
	copy: boolean
}

// A story written but not yet joined into one byte array: the caller can compare it with its source cheaply
// (Out.equals) and assemble it with the rest of its part (Out.pieces).
export interface StoryDraft {
	readonly out: Out
	readonly paragraphs: readonly WrittenParagraph[]
	readonly containers: readonly WrittenContainer[]
}

// A source container's children in order, and each one's index.
interface Siblings {
	readonly items: readonly number[]
	readonly index: ReadonlyMap<number, number>
}

const NO_SIBLINGS: Siblings = { items: [], index: new Map() }

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
	private readonly env: WriteEnvironment
	private readonly bytes: Uint8Array
	private force: Node | undefined
	// Table cells the block being written is in.
	private cells = 0
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
	private readonly siblings = new Map<number, Siblings>()
	// Set while writing in slices.
	private clock: SliceClock | undefined
	private readonly copiedRuns = new Set<number>()
	private readonly startedRuns = new Set<number>()
	private readonly openedWrappers = new Set<number>()
	private readonly paragraphs: WrittenParagraph[] = []
	private readonly textId: ReadonlySet<string> | undefined
	private readonly containers: WrittenContainer[] = []
	// Bases whose w:rPrChange has been written (copied runs included).
	private readonly writtenChanges = new Set<number>()

	constructor(env: WriteEnvironment, story: StoryDoc, options: StoryWriteOptions = {}) {
		this.env = env
		this.story = story
		this.part = story.part
		this.bytes = story.part.document.bytes
		this.force = options.force

		const reader = story.part.reader
		const root = story.part.document.documentElement()

		const w14 = reader.namespaces.prefixFor(root, W14_NAMESPACE)

		this.textId = w14 === undefined || w14 === "" ? undefined : new Set([`${w14}:textId`])
		this.names = new Names(reader.namespaces.prefixFor(root, W_NAMESPACE) ?? "", reader.namespaces.prefixFor(root, R_NAMESPACE) ?? "")
	}

	// The story container element (w:body, w:hdr, w:footnote...) written from `doc`, with what gate 2 needs
	// to check it without parsing it whole.
	writeStory(doc: Node): StoryOutput & { readonly bytes: Uint8Array } {
		const draft = complete(this.draft(doc))

		return { bytes: draft.out.finish(), trusted: draft.out.trusted, paragraphs: draft.paragraphs, containers: draft.containers }
	}

	// writeStory without joining the output (see StoryDraft).
	draftStorySync(doc: Node): StoryDraft {
		return complete(this.draft(doc))
	}

	// writeStory in slices: the clock's pause runs between them (a rejecting pause abandons the write). The
	// output is not joined yet (see StoryDraft).
	async draftStory(doc: Node, clock: SliceClock): Promise<StoryDraft> {
		this.clock = clock

		try {
			return await drive(this.draft(doc), clock)
		} finally {
			this.clock = undefined
		}
	}

	private *draft(doc: Node): Generator<void, StoryDraft> {
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
		const document = this.part.document
		const container = this.part.node(this.story.story.item)

		const sectPr = doc.attrs["sectPr"] as number

		openTag(out, document, container)

		const tail = yield* this.writeChildren(out, doc, this.story.story.item, (child, at) => this.writeBlock(at, child), {
			from: document.openEnd(container),
			to: document.contentEnd(container),
			tail: sectPr === NO_SOURCE ? [] : [sectPr]
		})

		if (!tail && sectPr !== NO_SOURCE) this.copyWhole(out, this.part.node(sectPr))

		closeTag(out, document, container)

		return { out, paragraphs: [...this.paragraphs], containers: [...this.containers] }
	}

	// One block through the edit path, for gate 1: runs rebuilt from their bases, properties spliced,
	// containers rebuilt from their parts (a table's paragraphs are copied, each is proved on its own).
	// `copy`: an earlier block of the story holds the same source (a pasted copy), as writeStory finds it.
	// `cell`: whether the block is in a table cell (by default, where its source is).
	writeBlockAlone(node: Node, copy = false, cell?: boolean): Out {
		this.force = node
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
				const run = marks.run.isInSet(child.marks)

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
	private due(): boolean {
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

	private owns(node: Node): boolean {
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

	private original(node: Node): Node | undefined {
		return originalBlocks(this.story).get(src(node))
	}

	// Whether a block can be written as its source bytes (see StoryDoc), in a table cell (`cell`) or not.
	private canCopy(node: Node, cell = this.cells > 0): boolean {
		if (this.repeating > 0 || this.claimed.has(src(node))) return false

		const memo = cell ? this.copyableInCell : this.copyable
		let result = memo.get(node)

		if (result !== undefined) return result

		if (node === this.force || !this.owns(node)) {
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

	private siblingsOf(container: number): Siblings {
		let siblings = this.siblings.get(container)

		if (siblings === undefined) {
			const items: number[] = []
			const index = new Map<number, number>()

			for (let item = container + 1; item < this.part.end(container); item = this.part.end(item)) {
				index.set(item, items.length)
				items.push(item)
			}

			siblings = { items, index }
			this.siblings.set(container, siblings)
		}

		return siblings
	}

	// The children of a container: blocks that can be copied keep the bytes around them (whitespace,
	// comments, proofErr) when their source neighbours are also their neighbours here. `region` is the
	// container's content in the source and what must follow the last child there (the body's final
	// section properties); when everything after the last child is as it was, it is copied whole and true
	// returned.
	private *writeChildren(
		out: Out,
		container: Node,
		sourceContainer: number,
		write: (child: Node, out: Out) => Generator<void, void>,
		region?: { readonly from: number; readonly to: number; readonly tail: readonly number[] }
	): Generator<void, boolean> {
		const { items: siblings, index: positions } = sourceContainer === NO_SOURCE ? NO_SIBLINGS : this.siblingsOf(sourceContainer)
		let previous = -1

		for (let at = 0; at < container.childCount; at++) {
			const child = container.child(at)
			const before = (child.attrs["before"] as MarkerList | undefined) ?? []

			if (this.due()) yield

			if (this.canCopy(child)) {
				const own = src(child)
				const index = positions.get(own) ?? -1
				const leading = index < 0 ? [] : this.leadingMarkers(siblings, index)
				const together = leading.length === before.length && leading.every((item, k) => before[k]?.src === item)
				const first = together && leading.length > 0 ? (leading[0] ?? own) : own

				if (!together) for (const marker of before) this.writeMarker(out, marker.src)

				const firstIndex = positions.get(first) ?? -1

				if (previous >= 0 && firstIndex > 0 && siblings[firstIndex - 1] === previous) {
					out.copy(this.bytes, this.part.byteEnd(previous), this.part.start(first))
				} else if (previous === -1 && firstIndex === 0 && region !== undefined) {
					out.copy(this.bytes, region.from, this.part.start(first))
				}

				out.copy(this.bytes, this.part.start(first), this.part.byteEnd(own))
				this.claim(child)
				previous = own

				continue
			}

			for (const marker of before) this.writeMarker(out, marker.src)

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
			const rest = siblings.slice((positions.get(previous) ?? -1) + 1)
			const expected = [...end.map(marker => marker.src), ...region.tail]

			if (rest.length === expected.length && rest.every((item, k) => item === expected[k])) {
				out.copy(this.bytes, this.part.byteEnd(previous), region.to)

				return true
			}
		}

		for (const marker of end) this.writeMarker(out, marker.src)

		return false
	}

	private leadingMarkers(siblings: readonly number[], index: number): number[] {
		let start = index

		while (start > 0 && MARKER_ITEMS.has(this.part.kinds[siblings[start - 1] ?? -1] ?? 0)) start--

		return siblings.slice(start, index)
	}

	private writeMarker(out: Out, item: number): void {
		if (item === NO_SOURCE) throw new DocxWriteError("missingSource", "A new range marker")

		this.copyWhole(out, this.part.node(item))
	}

	private *writeBlock(out: Out, node: Node): Generator<void, void> {
		switch (node.type) {
			case nodes.paragraph:
				this.writeParagraph(out, node)

				return
			case nodes.table:
				yield* this.writeTable(out, node)

				return
			case nodes.sdt:
				yield* this.writeSdt(out, node)

				return
			case nodes.locked_block:
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

	private *writeTable(out: Out, node: Node): Generator<void, void> {
		const document = this.part.document
		const names = this.names

		const item = src(node)

		if (!this.owns(node) && item !== NO_SOURCE && this.part.kinds[item] === Item.Table) {
			// A second table from one source (a split, a move of part of it): its own properties, no ids.
			this.checkUnchanged(node, ["styleId", "grid"], "Changed table properties")

			const row = this.part.aux[item] ?? -1

			out.text(`<${names.el("tbl")}>`)
			this.copyProperties(out, [this.part.tables.tblPr[row] ?? -1, this.part.tables.tblGrid[row] ?? -1], TRACKED_TABLE)
			yield* this.writeChildren(out, node, NO_SOURCE, (child, rowOut) => this.writeRow(rowOut, child))
			out.text(`</${names.el("tbl")}>`)

			return
		}

		if (!this.owns(node)) {
			let columns = (node.attrs["grid"] as readonly number[]).length

			node.forEach(row => {
				let count = 0

				row.forEach(cell => {
					count += cell.attrs["colspan"] as number
				})
				columns = Math.max(columns, count)
			})

			const grid =
				(node.attrs["grid"] as readonly number[]).length > 0
					? (node.attrs["grid"] as readonly number[])
					: Array.from({ length: columns }, () => 2000)
			const styleId = node.attrs["styleId"] as string | null

			out.text(
				`<${names.el("tbl")}><${names.el("tblPr")}>${styleId === null ? "" : names.val("tblStyle", styleId)}${names.empty("tblW", [
					["w", "0"],
					["type", "auto"]
				])}</${names.el("tblPr")}><${names.el("tblGrid")}>${grid.map(width => names.empty("gridCol", [["w", String(width)]])).join("")}</${names.el("tblGrid")}>`
			)
			yield* this.writeChildren(out, node, NO_SOURCE, (row, rowOut) => this.writeRow(rowOut, row))
			out.text(`</${names.el("tbl")}>`)

			return
		}

		const row = this.part.aux[item] ?? -1

		this.checkUnchanged(node, ["styleId", "grid"], "Changed table properties")

		const element = this.part.node(item)

		openTag(out, document, element)

		let from = document.openEnd(element)

		for (const property of [this.part.tables.tblPr[row] ?? -1, this.part.tables.tblGrid[row] ?? -1].sort((a, b) => a - b)) {
			if (property < 0) continue

			this.copyWhole(out, property)
			from = document.end(property)
		}

		yield* this.writeChildren(out, node, item, (child, rowOut) => this.writeRow(rowOut, child), {
			from,
			to: document.contentEnd(element),
			tail: []
		})
		closeTag(out, document, element)
	}

	private *writeRow(out: Out, node: Node): Generator<void, void> {
		const document = this.part.document
		const own = src(node)

		if (!this.owns(node) && own !== NO_SOURCE && this.part.kinds[own] === Item.Row) {
			const row = this.part.aux[own] ?? -1

			out.text(`<${this.names.el("tr")}>`)
			this.copyProperties(out, [this.part.rows.trPr[row] ?? -1, this.part.rows.tblPrEx[row] ?? -1], TRACKED_ROW)
			yield* this.writeChildren(out, node, NO_SOURCE, (cell, cellOut) => this.writeCell(cellOut, cell))
			out.text(`</${this.names.el("tr")}>`)

			return
		}

		if (!this.owns(node)) {
			out.text(`<${this.names.el("tr")}>`)
			yield* this.writeChildren(out, node, NO_SOURCE, (cell, cellOut) => this.writeCell(cellOut, cell))
			out.text(`</${this.names.el("tr")}>`)

			return
		}

		const item = src(node)
		const row = this.part.aux[item] ?? -1
		const element = this.part.node(item)

		openTag(out, document, element)

		// w:tblPrEx and w:trPr in the order the source has them (CT_Row puts tblPrEx first).
		let from = document.openEnd(element)

		for (const property of [this.part.rows.trPr[row] ?? -1, this.part.rows.tblPrEx[row] ?? -1].sort((a, b) => a - b)) {
			if (property < 0) continue

			this.copyWhole(out, property)
			from = document.end(property)
		}

		yield* this.writeChildren(out, node, item, (cell, cellOut) => this.writeCell(cellOut, cell), {
			from,
			to: document.contentEnd(element),
			tail: []
		})
		closeTag(out, document, element)
	}

	private *writeCell(out: Out, node: Node): Generator<void, void> {
		this.cells++

		try {
			yield* this.writeCellContent(out, node)
		} finally {
			this.cells--
		}
	}

	private *writeCellContent(out: Out, node: Node): Generator<void, void> {
		const document = this.part.document
		const names = this.names
		// Word wants a cell to end with a paragraph, which a table as its last block would not.
		const endsWithParagraph = !endsWithTable(node)

		const own = src(node)

		if (!this.owns(node) && own !== NO_SOURCE && this.part.kinds[own] === Item.Cell) {
			this.checkUnchanged(node, ["colspan", "vMerge"], "Changed cell merging")
			out.text(`<${names.el("tc")}>`)
			this.copyProperties(out, [this.part.cells.tcPr[this.part.aux[own] ?? -1] ?? -1], TRACKED_CELL)
			yield* this.writeChildren(out, node, NO_SOURCE, (child, blockOut) => this.writeBlock(blockOut, child))

			if (!endsWithParagraph) out.text(`<${names.el("p")}/>`)

			out.text(`</${names.el("tc")}>`)

			return
		}

		if (!this.owns(node)) {
			const span = node.attrs["colspan"] as number
			const vMerge = node.attrs["vMerge"] as string | null
			const tcPr = `${span > 1 ? names.val("gridSpan", String(span)) : ""}${vMerge === null ? "" : vMerge === "continue" ? names.empty("vMerge") : names.val("vMerge", vMerge)}`

			out.text(`<${names.el("tc")}>${tcPr === "" ? "" : `<${names.el("tcPr")}>${tcPr}</${names.el("tcPr")}>`}`)
			yield* this.writeChildren(out, node, NO_SOURCE, (child, blockOut) => this.writeBlock(blockOut, child))

			if (!endsWithParagraph) out.text(`<${names.el("p")}/>`)

			out.text(`</${names.el("tc")}>`)

			return
		}

		const item = own
		const row = this.part.aux[item] ?? -1

		this.checkUnchanged(node, ["colspan", "vMerge"], "Changed cell merging")

		const element = this.part.node(item)
		const tcPr = this.part.cells.tcPr[row] ?? -1

		openTag(out, document, element)

		if (tcPr >= 0) this.copyWhole(out, tcPr)

		yield* this.writeChildren(out, node, item, (child, blockOut) => this.writeBlock(blockOut, child), {
			from: tcPr >= 0 ? document.end(tcPr) : document.openEnd(element),
			to: document.contentEnd(element),
			tail: []
		})

		if (!endsWithParagraph) out.text(`<${names.el("p")}/>`)

		closeTag(out, document, element)
	}

	private *writeSdt(out: Out, node: Node): Generator<void, void> {
		const document = this.part.document
		const names = this.names

		const own = src(node)

		if (!this.owns(node) && own !== NO_SOURCE && this.part.kinds[own] === Item.Sdt) {
			const row = this.part.aux[own] ?? -1

			out.text(`<${names.el("sdt")}>`)
			this.copyProperties(out, [this.part.sdts.sdtPr[row] ?? -1, this.part.sdts.sdtEndPr[row] ?? -1], ["id"])
			out.text(`<${names.el("sdtContent")}>`)
			yield* this.writeChildren(out, node, NO_SOURCE, (child, blockOut) => this.writeBlock(blockOut, child))
			out.text(`</${names.el("sdtContent")}></${names.el("sdt")}>`)

			return
		}

		if (!this.owns(node)) {
			out.text(`<${names.el("sdt")}><${names.el("sdtPr")}/><${names.el("sdtContent")}>`)
			yield* this.writeChildren(out, node, NO_SOURCE, (child, blockOut) => this.writeBlock(blockOut, child))
			out.text(`</${names.el("sdtContent")}></${names.el("sdt")}>`)

			return
		}

		const item = src(node)
		const content = this.part.sdts.content[this.part.aux[item] ?? -1] ?? -1

		if (content < 0) throw new DocxWriteError("missingSource", "A content control without content")

		const element = this.part.node(item)

		out.bytes(this.bytes, document.start(element), document.openEnd(content))
		yield* this.writeChildren(out, node, item, (child, blockOut) => this.writeBlock(blockOut, child), {
			from: document.openEnd(content),
			to: document.contentEnd(content),
			tail: []
		})
		out.bytes(this.bytes, document.contentEnd(content), document.end(element))
	}

	private copyWhole(out: Out, node: number): void {
		out.copy(this.bytes, this.part.document.start(node), this.part.document.end(node))
	}

	// A container's property elements for a copy of it: everything but what carries an id of its own.
	private copyProperties(out: Out, properties: readonly number[], strip: readonly string[]): void {
		const changes = new Map(strip.map(local => [local, ""]))
		const sorted = properties.filter(property => property >= 0).sort((a, b) => a - b)
		const from = out.size

		for (const property of sorted) spliceChildren(out, this.part.reader, property, changes, [])

		this.containers.push({ properties: sorted, strip, from, to: out.size })
	}

	// Container attributes no command changes yet: the writer has no way to express a change to them.
	private checkUnchanged(node: Node, names: readonly string[], message: string): void {
		const original = this.original(node)

		if (original !== undefined && names.some(name => original.attrs[name] !== node.attrs[name]))
			throw new DocxWriteError("unsupported", message)
	}

	// ---- paragraphs --------------------------------------------------------------------------------------

	private writeParagraph(out: Out, node: Node): void {
		const names = this.names
		const document = this.part.document
		const own = src(node)
		const source = own !== NO_SOURCE && this.part.kinds[own] === Item.Paragraph ? own : NO_SOURCE
		const owner = source !== NO_SOURCE && this.owns(node)

		const from = out.size

		// Word's w14:textId hashes the paragraph's text: an edited paragraph drops it (Word writes a new one);
		// the paragraph keeps its w14:paraId.
		if (owner && this.force === undefined && this.textId !== undefined)
			openTagWithout(out, document, this.part.node(source), this.textId)
		else if (owner) openTag(out, document, this.part.node(source))
		else out.text(`<${names.el("p")}>`)

		this.writeParagraphProperties(out, node, source, owner)
		this.writeInline(out, node)

		if (owner) closeTag(out, document, this.part.node(source))
		else out.text(`</${names.el("p")}>`)

		this.paragraphs.push({ node, from, to: out.size })
	}

	private nested(write: (out: Out) => void): string {
		const out = new Out()

		write(out)

		return decodeUtf8(out.finish())
	}

	private writeParagraphProperties(out: Out, node: Node, source: number, owner: boolean): void {
		const names = this.names
		const part = this.part
		const reader = part.reader
		const document = part.document
		const row = source === NO_SOURCE ? -1 : (part.aux[source] ?? -1)
		const pPr = row < 0 ? -1 : (part.paragraphs.pPr[row] ?? -1)
		const props = this.env.context.store.paragraph(document, pPr)
		const changes = new Map<string, string>()
		const styleId = node.attrs["styleId"] as string | null
		const numId = node.attrs["numId"] as number | null
		const ilvl = node.attrs["ilvl"] as number | null
		const align = node.attrs["align"] as string | null
		const indent = node.attrs["indent"] as IndentAttrs | null
		const sectPr = node.attrs["sectPr"] as number
		const sourceNumPr = row < 0 ? -1 : (part.paragraphs.numPr[row] ?? -1)

		if (styleId !== (props.styleId ?? null)) {
			if (styleId !== null) checkName("w:pStyle", styleId)

			changes.set("pStyle", styleId === null ? "" : names.val("pStyle", styleId))
		}

		const numberingChanged = numId !== (props.numbering?.numId ?? null) || ilvl !== (props.numbering?.ilvl ?? null)

		if (numberingChanged || (!owner && sourceNumPr >= 0)) {
			const fileNumId = numId === null ? null : numId < 0 ? this.env.numId(numId) : numId

			if (fileNumId !== null) integer("w:numId", fileNumId, 0)
			if (ilvl !== null) integer("w:ilvl", ilvl, 0, 8)
			const numberingChanges = new Map<string, string>([
				["ilvl", ilvl === null ? "" : names.val("ilvl", String(ilvl))],
				["numId", fileNumId === null ? "" : names.val("numId", String(fileNumId))]
			])

			if (!owner) {
				numberingChanges.set("numberingChange", "")
				numberingChanges.set("ins", "")
			}

			if (numId === null && ilvl === null) changes.set("numPr", "")
			else if (sourceNumPr >= 0)
				changes.set(
					"numPr",
					this.nested(sub => spliceChildren(sub, reader, sourceNumPr, numberingChanges, NUMBERING_PROPERTIES))
				)
			else changes.set("numPr", freshContainer(names, "numPr", numberingChanges, NUMBERING_PROPERTIES))
		}

		if (align !== (props.alignment ?? null)) {
			if (align !== null) oneOf("w:jc", JUSTIFICATION, align)

			changes.set("jc", align === null ? "" : names.val("jc", align))
		}

		if (indent !== (props.indent ?? null)) {
			checkIndent(indent)
			changes.set(
				"ind",
				indent === null
					? ""
					: names.empty("ind", [
							["left", indent.start?.toString()],
							["leftChars", indent.startChars?.toString()],
							["right", indent.end?.toString()],
							["rightChars", indent.endChars?.toString()],
							["hanging", indent.hanging?.toString()],
							["hangingChars", indent.hangingChars?.toString()],
							["firstLine", indent.firstLine?.toString()],
							["firstLineChars", indent.firstLineChars?.toString()]
						])
			)
		}

		const hasSection = row >= 0 && (part.paragraphs.sectPr[row] ?? -1) >= 0

		if (!sectionWritten(part, sectPr, this.cells > 0)) {
			if (hasSection) changes.set("sectPr", "")
		} else if (sectPr !== source || !hasSection) {
			const from = part.paragraphs.sectPr[part.aux[sectPr] ?? -1] ?? -1

			if (from < 0) throw new DocxWriteError("missingSource", "A section break without its properties")

			changes.set("sectPr", decodeUtf8(this.bytes, document.start(from), document.end(from)))
		}

		if (!owner && pPr >= 0) {
			// A copy of the paragraph's properties is a new paragraph mark: tracked changes stay with the original.
			changes.set("pPrChange", "")

			const markRPr = part.paragraphs.markRPr[row] ?? -1

			if (markRPr >= 0)
				changes.set(
					"rPr",
					this.nested(sub => spliceChildren(sub, reader, markRPr, MARK_REVISIONS, PARAGRAPH_MARK_PROPERTIES))
				)
		}

		if (pPr >= 0) {
			if (changes.size === 0) {
				copyNode(out, document, pPr)

				return
			}

			// Properties that lost their last child are left out rather than written empty.
			const spliced = this.nested(sub => spliceChildren(sub, reader, pPr, changes, PARAGRAPH_PROPERTIES))

			if (!/^<[^>]*>\s*<\/[^>]*>$/.test(spliced)) out.text(spliced)
		} else {
			out.text(freshContainer(names, "pPr", changes, PARAGRAPH_PROPERTIES))
		}
	}

	// ---- inline content ----------------------------------------------------------------------------------

	private units(paragraph: Node): Unit[] {
		const units: Unit[] = []

		paragraph.forEach(child => {
			const run = marks.run.isInSet(child.marks)
			const wrappers = this.wrappersOf(child)
			const inRun = inlineLevel(child) === "run" || (child.type === nodes.marker && run !== undefined)
			const last = units[units.length - 1]

			if (inRun && last !== undefined && !last.atom && sameMark(last.run, run) && sameMarks(last.wrappers, wrappers)) {
				last.nodes.push(child)
			} else {
				units.push({ nodes: [child], run, wrappers, atom: !inRun, copy: false })
			}
		})

		return units
	}

	// The node's wrapper marks, outermost first: source wrappers by their item (an outer element comes
	// first in document order), new ones innermost.
	private wrappersOf(node: Node): Mark[] {
		const wrappers = node.marks.filter(mark => wrapperTypes.has(mark.type)).sort((a, b) => wrapperRank(a) - wrapperRank(b))
		const link = wrappers.findIndex(mark => mark.type === marks.link && src(mark) === NO_SOURCE)
		const revision = wrappers.findIndex(mark => REVISION_WRAPPERS.has(mark.type))

		// A tracked change holds runs only (CT_RunTrackChange): a new link goes around it.
		if (link >= 0 && revision >= 0 && revision < link) wrappers.splice(revision, 0, ...wrappers.splice(link, 1))

		return wrappers
	}

	// Content tokens of a source run (text and atoms by item), as its unit would show them.
	private sourceTokens(run: number): string {
		const part = this.part
		let out = ""

		for (let item = run + 1; item < part.end(run); item = part.end(item)) {
			const kind = part.kinds[item]

			if (kind === Item.Text || kind === Item.DeletedText) out += part.text(item)
			else out += `\u0000${item}\u0000`
		}

		return out
	}

	private unitTokens(unit: Unit): string {
		let out = ""

		for (const node of unit.nodes) out += node.isText ? (node.text ?? "") : `\u0000${src(node)}\u0000`

		return out
	}

	private sourceWrappers(run: number): number[] {
		const out: number[] = []
		const part = this.part

		for (let at = part.parent(run); at >= 0 && part.kinds[at] !== Item.Paragraph; at = part.parent(at)) {
			if (WRAPPER_ITEMS.has(part.kinds[at] ?? 0)) out.push(at)
		}

		return out.reverse()
	}

	private base(run: Mark | undefined): RunBase | undefined {
		const id = run?.attrs["base"] as number | undefined

		return id === undefined || id < 0 ? undefined : this.env.context.bases.get(id)
	}

	private canCopyUnit(unit: Unit, forced: boolean): boolean {
		if (forced || unit.atom || unit.run === undefined) return false

		const run = src(unit.run)

		if (run === NO_SOURCE || this.copiedRuns.has(run) || this.part.kinds[run] !== Item.Run) return false

		const base = this.base(unit.run)

		for (const node of unit.nodes) if (node.type !== nodes.marker && !sameFormatting(node.marks, base)) return false

		const wrappers = this.sourceWrappers(run)

		if (wrappers.length !== unit.wrappers.length || unit.wrappers.some((mark, k) => src(mark) !== wrappers[k])) return false

		return this.unitTokens(unit) === this.sourceTokens(run)
	}

	private writeInline(out: Out, paragraph: Node): void {
		const forced = paragraph === this.force
		const units = this.units(paragraph)

		for (const unit of units) {
			unit.copy = this.canCopyUnit(unit, forced)

			if (unit.copy && unit.run !== undefined) {
				this.copiedRuns.add(src(unit.run))

				const base = this.base(unit.run)

				if (base?.props.change !== undefined) this.writtenChanges.add(base.id)
			}
		}

		const stack: Mark[] = []
		const document = this.part.document

		units.forEach((unit, index) => {
			const node = unit.nodes[0]

			if (node === undefined) return

			// Rendered page breaks are Word's layout cache: rebuilt runs leave them out (and an otherwise
			// empty wrapper with them).
			if (!forced && !unit.copy && unit.nodes.every(child => child.type === nodes.rendered_break)) return

			// Bookkeeping marks stay only between content written as it was.
			if (node.type === nodes.proof_err) {
				const previous = units[index - 1]
				const next = units[index + 1]

				// Gate 1 keeps them: it proves the run rebuilding, and with them most paragraphs come out as
				// their very bytes.
				if (
					!forced &&
					((previous !== undefined && !previous.atom && !previous.copy) || (next !== undefined && !next.atom && !next.copy))
				) {
					return
				}
			}

			this.transition(out, stack, unit.wrappers)

			if (unit.atom) {
				this.writeAtom(out, node)

				return
			}

			if (unit.copy && unit.run !== undefined) {
				copyNode(out, document, this.part.node(src(unit.run)))

				return
			}

			this.writeRuns(out, unit, forced)
		})

		this.transition(out, stack, [])
	}

	private transition(out: Out, stack: Mark[], wanted: readonly Mark[]): void {
		let common = 0

		while (common < stack.length && common < wanted.length && stack[common]?.eq(wanted[common] as Mark) === true) common++

		while (stack.length > common) {
			const mark = stack.pop()

			if (mark !== undefined) this.closeWrapper(out, mark)
		}

		for (let k = common; k < wanted.length; k++) {
			const mark = wanted[k] as Mark

			this.openWrapper(out, mark)
			stack.push(mark)
		}
	}

	private openWrapper(out: Out, mark: Mark): void {
		const names = this.names
		const own = src(mark)

		if (own === NO_SOURCE) {
			if (mark.type !== marks.link) throw new DocxWriteError("missingSource", `A new ${mark.type.name} wrapper`)

			const anchor = mark.attrs["anchor"] as string | null
			const tooltip = mark.attrs["tooltip"] as string | null
			let attributes = ""

			const target = (mark.attrs["target"] as string | null) ?? ""

			// An external target is a relationship; a bookmark is w:anchor; a link may have both.
			if (target !== "") {
				const id = escapeAttribute(this.env.hyperlink(this.part.name, target))

				// A part that never needed relationships may not bind their namespace: declare it here.
				attributes += names.r === "" ? ` xmlns:${LOCAL_R}="${R_NAMESPACE}" ${LOCAL_R}:id="${id}"` : ` ${names.relId()}="${id}"`
			}

			if (anchor !== null && anchor !== "") attributes += ` ${names.attr("anchor")}="${escapeAttribute(anchor)}"`

			if (tooltip !== null) attributes += ` ${names.attr("tooltip")}="${escapeAttribute(tooltip)}"`

			out.text(`<${names.el("hyperlink")}${attributes}>`)

			return
		}

		const part = this.part
		const document = part.document
		const element = part.node(own)
		const again = this.openedWrappers.has(own)

		this.openedWrappers.add(own)

		switch (part.kinds[own]) {
			case Item.Sdt: {
				const row = part.aux[own] ?? -1
				const content = part.sdts.content[row] ?? -1
				const sdtPr = part.sdts.sdtPr[row] ?? -1

				if (content < 0) throw new DocxWriteError("missingSource", "A content control without content")

				if (!again) {
					out.bytes(this.bytes, document.start(element), document.openEnd(content))

					return
				}

				// A content control split in two: the second part has no w:id of its own.
				openTag(out, document, element)

				for (
					let child = document.firstChildElement(element);
					child >= 0 && child !== content;
					child = document.nextSiblingElement(child)
				) {
					if (child === sdtPr) spliceChildren(out, part.reader, sdtPr, new Map([["id", ""]]), [])
					else copyNode(out, document, child)
				}

				openTag(out, document, content)

				return
			}
			case Item.SmartTag:
			case Item.CustomXml: {
				openTag(out, document, element)

				const properties = part.aux[own] ?? -1

				if (properties >= 0) copyNode(out, document, properties)

				return
			}
			case Item.Insertion:
			case Item.Deletion:
			case Item.MoveFrom:
			case Item.MoveTo:
				if (again) startTagWith(out, document, element, names.attr("id"), String(this.env.nextRevisionId()))
				else openTag(out, document, element)

				return
			default:
				openTag(out, document, element)
		}
	}

	private closeWrapper(out: Out, mark: Mark): void {
		const own = src(mark)

		if (own === NO_SOURCE) {
			out.text(`</${this.names.el("hyperlink")}>`)

			return
		}

		const part = this.part
		const document = part.document
		const element = part.node(own)

		if (part.kinds[own] === Item.Sdt) {
			const content = part.sdts.content[part.aux[own] ?? -1] ?? -1

			out.bytes(this.bytes, document.contentEnd(content), document.end(element))

			return
		}

		closeTag(out, document, element)
	}

	private writeAtom(out: Out, node: Node): void {
		const document = this.part.document
		const own = src(node)

		switch (node.type) {
			case nodes.proof_err: {
				const xmlNode = node.attrs["node"] as number

				if (xmlNode >= 0) copyNode(out, document, xmlNode)

				return
			}
			case nodes.field:
				if (own !== NO_SOURCE) {
					out.bytes(this.bytes, this.part.start(own), this.part.byteEnd(node.attrs["end"] as number))

					return
				}

				this.writeNewField(out, node)

				return
			default:
				if (own === NO_SOURCE) throw new DocxWriteError("missingSource", `A new ${node.type.name}`)

				copyNode(out, document, this.part.node(own))
		}
	}

	private writeNewField(out: Out, node: Node): void {
		const names = this.names
		const instruction = escapeAttribute(node.attrs["instruction"] as string)
		const result = node.attrs["result"] as string

		if (result === "") {
			out.text(`<${names.el("fldSimple")} ${names.attr("instr")}="${instruction}"/>`)

			return
		}

		out.text(
			`<${names.el("fldSimple")} ${names.attr("instr")}="${instruction}"><${names.el("r")}>${this.textElement(result, false)}</${names.el("r")}></${names.el("fldSimple")}>`
		)
	}

	private textElement(text: string, deleted: boolean): string {
		const local = this.names.el(deleted ? "delText" : "t")
		const space = PRESERVE.test(text) ? ' xml:space="preserve"' : ""

		return `<${local}${space}>${escapeText(text)}</${local}>`
	}

	// A unit rebuilt: consecutive nodes with the same marks become one w:r with the base's properties and
	// their changes; markers Word had put inside the run go between the runs.
	private writeRuns(out: Out, unit: Unit, forced: boolean): void {
		const names = this.names
		const document = this.part.document
		const run = unit.run === undefined ? NO_SOURCE : src(unit.run)
		const base = this.base(unit.run)
		let open: readonly Mark[] | null = null
		let raw = false

		const close = (): void => {
			if (open === null) return

			out.text(raw ? `</${document.name(this.part.node(run))}>` : `</${names.el("r")}>`)
			open = null
		}

		for (const node of unit.nodes) {
			if (node.type === nodes.marker) {
				close()
				this.writeAtom(out, node)

				continue
			}

			// Word's record of its last page break goes stale with the edit (gate 1 keeps it, see above).
			if (node.type === nodes.rendered_break && !forced) continue

			if (open === null || !sameMarks(open, node.marks)) {
				close()
				raw = run !== NO_SOURCE && this.part.kinds[run] === Item.Run && !this.startedRuns.has(run)

				if (raw) {
					this.startedRuns.add(run)
					openTag(out, document, this.part.node(run))
				} else {
					out.text(`<${names.el("r")}>`)
				}

				// A tracked formatting change written a second time needs an id of its own.
				const changed = base?.props.change !== undefined

				writeRunProperties(
					out,
					names,
					node.marks,
					base,
					changed && this.writtenChanges.has(base.id) ? this.env.nextRevisionId() : undefined
				)

				if (changed) this.writtenChanges.add(base.id)
				open = node.marks
			}

			this.writeRunContent(out, node)
		}

		close()
	}

	private writeRunContent(out: Out, node: Node): void {
		const names = this.names
		const own = src(node)

		if (node.isText) {
			// Deleted text is w:delText; moved-away text stays w:t, as Word writes it.
			const deleted = marks.deletion.isInSet(node.marks) !== undefined

			out.text(this.textElement(node.text ?? "", deleted))

			return
		}

		if (own !== NO_SOURCE) {
			copyNode(out, this.part.document, this.part.node(own))

			return
		}

		switch (node.type) {
			case nodes.tab:
				out.text(names.empty("tab"))

				return
			case nodes.break: {
				if (node.attrs["cr"] === true) {
					out.text(names.empty("cr"))

					return
				}

				const type = node.attrs["type"] as string
				const clear = node.attrs["clear"] as string

				out.text(
					names.empty("br", [
						["type", type === "textWrapping" ? null : type],
						["clear", clear === "none" ? null : clear]
					])
				)

				return
			}
			case nodes.hyphen:
				out.text(names.empty(node.attrs["soft"] === true ? "softHyphen" : "noBreakHyphen"))

				return
			case nodes.symbol:
				out.text(
					names.empty("sym", [
						["font", node.attrs["font"] as string | null],
						["char", (node.attrs["char"] as number).toString(16).toUpperCase().padStart(4, "0")]
					])
				)

				return
			default:
				throw new DocxWriteError("missingSource", `A new ${node.type.name}`)
		}
	}
}

const REVISION_WRAPPERS: ReadonlySet<MarkType> = new Set([marks.insertion, marks.deletion, marks.move_from, marks.move_to])

function sameParagraph(original: Node | undefined, node: Node): boolean {
	if (original === undefined) return false
	if (original.eq(node)) return true
	if (!original.content.eq(node.content) || !Mark.sameSet(original.marks, node.marks)) return false

	for (const name of Object.keys(original.attrs)) {
		if (name !== "before" && original.attrs[name] !== node.attrs[name]) return false
	}

	return true
}

function wrapperRank(mark: Mark): number {
	const own = src(mark)

	return own === NO_SOURCE ? Number.MAX_SAFE_INTEGER : own
}

function sameMark(a: Mark | undefined, b: Mark | undefined): boolean {
	return a === undefined ? b === undefined : b !== undefined && a.eq(b)
}

function sameMarks(a: readonly Mark[], b: readonly Mark[]): boolean {
	if (a === b) return true
	if (a.length !== b.length) return false

	for (let k = 0; k < a.length; k++) if (!(a[k] as Mark).eq(b[k] as Mark)) return false

	return true
}

let copyOwnerNode: Node | undefined

// Stands for the block owning sources a copy is written without.
function copyOwner(): Node {
	copyOwnerNode ??= docxSchema.nodes.paragraph.create()

	return copyOwnerNode
}
