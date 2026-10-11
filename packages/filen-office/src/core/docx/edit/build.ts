import { Mark, type Node } from "prosemirror-model"
import { NodeKind } from "../../xml/index"
import { FieldCharType, ITEM_NAMES, Item, type Story, type StoryPart } from "../model/index"
import { El, Ns } from "../model/namespaces"
import type { DocxEditContext } from "./context"
import {
	NO_MARKERS,
	NO_SOURCE,
	docxSchema,
	makeMark,
	makeNode,
	type LockReason,
	type LockedKind,
	type MarkerAttrs,
	type MarkerKind,
	type MarkerList
} from "./schema"

// One story as an editing document. What a writer may copy as source bytes (only the first node in
// document order that claims a `src` may; a split paragraph's second half or a copy does not):
// - a paragraph while the document still holds the very node the builder made, which `pristine` records
//   (ProseMirror keeps the identity of untouched nodes across transactions);
// - a locked block always (its attributes are protected and it has no content);
// - a table, row, cell or content control while its attributes still say what its source does and its
//   children are its source's child items, in order, each of them copyable.
// Containers are not recorded: a set entry per row and cell costs more than deriving it at save time.
export interface StoryDoc {
	readonly doc: Node
	readonly story: Story
	readonly part: StoryPart
	readonly pristine: WeakSet<Node>
	readonly stats: StoryBuildStats
}

export interface StoryBuildStats {
	readonly paragraphs: number
	readonly tables: number
	readonly lockedBlocks: number
	readonly lockReasons: Readonly<Partial<Record<LockReason, number>>>
}

const MARKER_KINDS: Partial<Record<number, MarkerKind>> = {
	[Item.BookmarkStart]: "bookmarkStart",
	[Item.BookmarkEnd]: "bookmarkEnd",
	[Item.CommentRangeStart]: "commentRangeStart",
	[Item.CommentRangeEnd]: "commentRangeEnd",
	[Item.MoveFromRangeStart]: "moveFromRangeStart",
	[Item.MoveFromRangeEnd]: "moveFromRangeEnd",
	[Item.MoveToRangeStart]: "moveToRangeStart",
	[Item.MoveToRangeEnd]: "moveToRangeEnd",
	[Item.PermissionStart]: "permissionStart",
	[Item.PermissionEnd]: "permissionEnd",
	[Item.CustomXmlRangeMarker]: "customXmlRange"
}

// Range markers the model keeps opaque because they sit inside a w:r.
const MARKER_ELEMENTS: Partial<Record<number, MarkerKind>> = {
	[El.BookmarkStart]: "bookmarkStart",
	[El.BookmarkEnd]: "bookmarkEnd",
	[El.CommentRangeStart]: "commentRangeStart",
	[El.CommentRangeEnd]: "commentRangeEnd",
	[El.MoveFromRangeStart]: "moveFromRangeStart",
	[El.MoveFromRangeEnd]: "moveFromRangeEnd",
	[El.MoveToRangeStart]: "moveToRangeStart",
	[El.MoveToRangeEnd]: "moveToRangeEnd",
	[El.PermStart]: "permissionStart",
	[El.PermEnd]: "permissionEnd",
	[El.CustomXmlRangeMarker]: "customXmlRange"
}

const REVISION_MARK_OF: Partial<Record<number, "insertion" | "deletion" | "move_from" | "move_to">> = {
	[Item.Insertion]: "insertion",
	[Item.Deletion]: "deletion",
	[Item.MoveFrom]: "move_from",
	[Item.MoveTo]: "move_to"
}

const WRAPPER_KINDS: Partial<Record<number, string>> = {
	[Item.SmartTag]: "smartTag",
	[Item.CustomXml]: "customXml"
}

const BREAK_TYPES = ["textWrapping", "page", "column"] as const
const BREAK_CLEARS = ["none", "left", "right", "all"] as const

const EDGE_WHITESPACE = /^[ \t\r\n]|[ \t\r\n]$/
const TRIM = /^[ \t\r\n]+|[ \t\r\n]+$/g

// Longest preview text kept on a locked block.
const PREVIEW_LIMIT = 2000

// w:proofErr elements of a part by the inline container holding them (they are no model items): each
// sits before the child item `before` (-1: after the last), `node` is its XML node.
interface ProofErrors {
	readonly before: number[]
	readonly node: number[]
}

interface PartInfo {
	readonly plan: FieldPlan
	// XML nodes whose markup could not be read, ascending.
	readonly malformed: Int32Array
	readonly proofErrors: ReadonlyMap<number, ProofErrors>
}

// The item whose XML node is `node`, or -1 (item nodes ascend: items are made in document order).
function itemOfNode(part: StoryPart, node: number): number {
	const nodes = part.nodes
	let low = 0
	let high = part.count

	while (low < high) {
		const middle = (low + high) >> 1

		if ((nodes[middle] ?? 0) < node) low = middle + 1
		else high = middle
	}

	return low < part.count && nodes[low] === node ? low : -1
}

function findProofErrors(part: StoryPart): Map<number, ProofErrors> {
	const out = new Map<number, ProofErrors>()
	const document = part.document
	const reader = part.reader
	// Names that can be w:proofErr (any prefix): most documents have none, and comparing interned name ids
	// is far cheaper than classifying every element.
	const names = new Set<number>()

	document.names.forEach((name, id) => {
		if (name === "proofErr" || name.endsWith(":proofErr")) names.add(id)
	})

	if (names.size === 0) return out

	const nameIds = document.nameIds

	for (let node = 0; node < document.count; node++) {
		if (!names.has(nameIds[node] ?? -1) || document.kinds[node] !== NodeKind.Element || reader.code(node) !== El.ProofErr) continue

		let parent = document.parents[node] ?? -1

		if (parent >= 0 && reader.is(parent, Ns.W, "sdtContent")) parent = document.parents[parent] ?? -1

		const container = itemOfNode(part, parent)

		if (container < 0) continue

		const start = document.start(node)
		let before = -1

		for (let child = container + 1; child < part.end(container); child = part.end(child)) {
			if (part.start(child) > start) {
				before = child

				break
			}
		}

		let list = out.get(container)

		if (list === undefined) {
			list = { before: [], node: [] }
			out.set(container, list)
		}

		list.before.push(before)
		list.node.push(node)
	}

	return out
}

// Complex fields of a part: those kept as one atom (begin run -> end run) and paragraphs that must be
// locked because a field does not fit in one.
interface FieldPlan {
	readonly atoms: Map<number, number>
	readonly locked: Map<number, LockReason>
}

function paragraphOf(part: StoryPart, item: number): number {
	for (let at = item; at >= 0; at = part.parent(at)) {
		const kind = part.kinds[at] ?? 0

		if (kind === Item.Paragraph) return at
		if (kind >= Item.Body && kind <= Item.Textbox) return -1
	}

	return -1
}

function planFields(part: StoryPart): FieldPlan {
	const fields = part.fields
	const atoms = new Map<number, number>()
	const locked = new Map<number, LockReason>()
	const lock = (paragraph: number, reason: LockReason): void => {
		if (paragraph >= 0 && locked.get(paragraph) !== "fieldSpansBlocks") locked.set(paragraph, reason)
	}

	for (let field = 0; field < fields.count; field++) {
		if (fields.simple[field] === 1) continue

		const begin = fields.begin[field] ?? -1
		const end = fields.end[field] ?? -1

		if (begin < 0) continue

		const first = paragraphOf(part, begin)

		if (end < 0) {
			lock(first, "fieldStructure")

			continue
		}

		const last = paragraphOf(part, end)

		if (first !== last) {
			const from = Math.min(first, last)
			const to = Math.max(first, last)

			if (from < 0) {
				lock(to, "fieldSpansBlocks")

				continue
			}

			for (let item = from; item <= to; item++) {
				if (part.kinds[item] === Item.Paragraph) locked.set(item, "fieldSpansBlocks")
			}

			continue
		}

		// A nested field travels inside its parent's atom.
		if ((fields.parent[field] ?? -1) >= 0) continue

		const beginRun = part.parent(begin)
		const endRun = part.parent(end)
		const whole =
			part.kinds[beginRun] === Item.Run &&
			part.kinds[endRun] === Item.Run &&
			part.parent(beginRun) === part.parent(endRun) &&
			beginRun + 1 === begin &&
			part.end(endRun) === end + 1 &&
			beginRun <= endRun

		if (whole) atoms.set(beginRun, endRun)
		else lock(first, "fieldStructure")
	}

	return { atoms, locked }
}

// Inline content collector: adjacent text with the same mark array (identity: arrays are cached per
// wrapper set and base) becomes one text node.
class InlineSink {
	readonly nodes: Node[] = []
	private text = ""
	private marks: readonly Mark[] = Mark.none

	addText(text: string, marks: readonly Mark[]): void {
		if (text === "") return

		if (marks !== this.marks) {
			this.flush()
			this.marks = marks
		}

		this.text += text
	}

	add(node: Node): void {
		this.flush()
		this.nodes.push(node)
	}

	flush(): void {
		if (this.text !== "") this.nodes.push(docxSchema.text(this.text, this.marks))

		this.text = ""
	}
}

// Builds editing documents from the model, one per story. Structural impossibilities (fields across
// paragraphs, odd table structure, unmodelled block-level elements) become locked blocks here; the
// round-trip proof can lock more later.
export class DocxDocBuilder {
	private readonly context: DocxEditContext
	private readonly infos = new Map<StoryPart, PartInfo>()
	private readonly combined = new WeakMap<readonly Mark[], Map<number, readonly Mark[]>>()
	private part!: StoryPart
	private plan!: FieldPlan
	private info!: PartInfo
	private pristine!: WeakSet<Node>
	private paragraphs = 0
	private tables = 0
	private lockedBlocks = 0
	private lockReasons: Partial<Record<LockReason, number>> = {}
	// Why the table being built cannot be edited, once something fails.
	private tableFailure: LockReason | null = null

	constructor(context: DocxEditContext) {
		this.context = context
	}

	build(story: Story): StoryDoc {
		const model = this.context.model
		const part = model.partOf(story)

		this.part = part
		this.plan = this.planOf(part)
		this.pristine = new WeakSet()
		this.paragraphs = 0
		this.tables = 0
		this.lockedBlocks = 0
		this.lockReasons = {}

		const blocks: Node[] = []
		let sectPr = NO_SOURCE
		const end = this.blocks(story.item, blocks, item => {
			if (part.kinds[item] !== Item.SectionProperties) return false

			sectPr = item

			return true
		})

		// A story with no block at all still needs a place for the cursor; the empty paragraph is new content.
		if (blocks.length === 0) blocks.push(docxSchema.nodes.paragraph.create())

		const doc = makeNode(docxSchema.nodes.doc, { story: story.index, sectPr, end }, blocks)

		return {
			doc,
			story,
			part,
			pristine: this.pristine,
			stats: { paragraphs: this.paragraphs, tables: this.tables, lockedBlocks: this.lockedBlocks, lockReasons: this.lockReasons }
		}
	}

	// A locked block standing for `node` (a source paragraph, table or content control of `story`), for a
	// block the round-trip proof rejected. It is written as its bytes.
	lockedBlock(story: StoryDoc, node: Node, reason: LockReason): Node {
		const src = node.attrs["src"]

		if (typeof src !== "number" || src < 0) throw new RangeError("Only source blocks can be locked")

		const kind: LockedKind =
			node.type === docxSchema.nodes.paragraph ? "paragraph" : node.type === docxSchema.nodes.table ? "table" : "sdt"

		this.part = story.part
		this.plan = this.planOf(story.part)

		return this.locked(src, (node.attrs["before"] as MarkerList | undefined) ?? NO_MARKERS, kind, reason)
	}

	private planOf(part: StoryPart): FieldPlan {
		let info = this.infos.get(part)

		if (info === undefined) {
			info = {
				plan: planFields(part),
				malformed: Int32Array.from(part.reader.malformedNodes).sort(),
				proofErrors: findProofErrors(part)
			}
			this.infos.set(part, info)
		}

		this.info = info

		return info.plan
	}

	// Whether markup inside the element of `item` could not be read (its own start tag included).
	private malformedWithin(node: number): boolean {
		const malformed = this.info.malformed

		if (node < 0 || malformed.length === 0) return false

		let low = 0
		let high = malformed.length

		while (low < high) {
			const middle = (low + high) >> 1

			if ((malformed[middle] ?? 0) < node) low = middle + 1
			else high = middle
		}

		const first = malformed[low] ?? -1

		return first >= 0 && this.part.document.start(first) < this.part.document.end(node)
	}

	// ---- blocks -----------------------------------------------------------------------------------------

	// Builds the block children of `container` into `out`; returns the markers after the last block.
	private blocks(container: number, out: Node[], special?: (item: number) => boolean): MarkerList {
		const part = this.part
		const end = part.end(container)
		let pending: MarkerAttrs[] | null = null

		for (let item = container + 1; item < end; item = part.end(item)) {
			const marker = this.marker(item)

			if (marker !== undefined) {
				;(pending ??= []).push(marker)

				continue
			}

			if (special?.(item) === true) continue

			out.push(this.block(item, pending ?? NO_MARKERS))
			pending = null
		}

		return pending ?? NO_MARKERS
	}

	private block(item: number, before: MarkerList): Node {
		const part = this.part

		switch (part.kinds[item]) {
			case Item.Paragraph:
				return this.paragraph(item, before)
			case Item.Table:
				return this.table(item, before)
			case Item.Sdt:
				return this.sdt(item, before)
			case Item.AltChunk:
				return this.locked(item, before, "element", "altChunk")
			case Item.CustomXml:
			case Item.Insertion:
			case Item.Deletion:
			case Item.MoveFrom:
			case Item.MoveTo:
			case Item.AlternateContent:
				return this.locked(item, before, "element", "blockWrapper")
			case Item.Opaque:
				return this.locked(item, before, "element", part.reader.malformedNodes.has(part.node(item)) ? "malformed" : "opaque")
			default:
				return this.locked(item, before, "element", "opaque")
		}
	}

	private keep(node: Node): Node {
		this.pristine.add(node)

		return node
	}

	private locked(item: number, before: MarkerList, kind: LockedKind, reason: LockReason, deletable: boolean = true): Node {
		const part = this.part
		const end = part.end(item)
		let sectPr = NO_SOURCE
		let spans = reason === "fieldSpansBlocks"

		for (let at = item; at < end; at++) {
			if (part.kinds[at] !== Item.Paragraph) continue

			if ((part.paragraphs.sectPr[part.aux[at] ?? -1] ?? -1) >= 0) sectPr = at
			if (this.plan.locked.get(at) === "fieldSpansBlocks") spans = true
		}

		this.lockedBlocks++
		this.lockReasons[reason] = (this.lockReasons[reason] ?? 0) + 1

		return makeNode(docxSchema.nodes.locked_block, {
			src: item,
			before,
			kind,
			reason,
			text: this.preview(item),
			markers: this.survivors(item + 1, end),
			deletable: deletable && sectPr === NO_SOURCE && !spans,
			sectPr
		})
	}

	private preview(item: number): string {
		const part = this.part
		const kind = part.kinds[item]

		if (kind === Item.Paragraph) return part.paragraphText(item).slice(0, PREVIEW_LIMIT)

		let out = ""
		const end = part.end(item)

		for (let at = item + 1; at < end && out.length < PREVIEW_LIMIT; at++) {
			if (part.kinds[at] === Item.Paragraph) {
				out += (out === "" ? "" : "\n") + part.paragraphText(at)
				at = part.end(at) - 1
			}
		}

		return out.slice(0, PREVIEW_LIMIT)
	}

	// Range markers and comment references in items [start, end), which outlive the atom holding them.
	private survivors(start: number, end: number): MarkerList {
		let out: MarkerAttrs[] | null = null

		for (let item = start; item < end; item++) {
			const marker = this.marker(item) ?? this.commentReference(item)

			if (marker !== undefined) (out ??= []).push(marker)
		}

		return out ?? NO_MARKERS
	}

	private marker(item: number): MarkerAttrs | undefined {
		const part = this.part
		const itemKind = part.kinds[item] ?? 0
		const kind = itemKind === Item.Opaque ? MARKER_ELEMENTS[part.reader.code(part.node(item))] : MARKER_KINDS[itemKind]

		return kind === undefined ? undefined : { src: item, kind, id: part.id(item) ?? null }
	}

	private commentReference(item: number): MarkerAttrs | undefined {
		return this.part.kinds[item] === Item.CommentReference
			? { src: item, kind: "commentReference", id: this.part.id(item) ?? null }
			: undefined
	}

	private paragraph(item: number, before: MarkerList): Node {
		const part = this.part
		const reason = this.plan.locked.get(item)

		this.paragraphs++

		if (reason !== undefined) return this.locked(item, before, "paragraph", reason)

		// Checked on the XML, not through the props store: a malformed w:rPr or w:pPr is reported there once
		// per distinct bytes only.
		if (this.malformedWithin(part.node(item))) return this.locked(item, before, "paragraph", "malformed")

		const row = part.aux[item] ?? -1
		const pPr = part.paragraphs.pPr[row] ?? -1
		const props = this.context.store.paragraph(part.document, pPr)
		const sink = new InlineSink()

		this.inline(item, Mark.none, sink)
		sink.flush()

		const sectPr = part.paragraphs.sectPr[row] ?? -1

		return this.keep(
			makeNode(
				docxSchema.nodes.paragraph,
				{
					src: item,
					before,
					styleId: props.styleId ?? null,
					numId: props.numbering?.numId ?? null,
					ilvl: props.numbering?.ilvl ?? null,
					align: props.alignment ?? null,
					indent: props.indent ?? null,
					sectPr: sectPr >= 0 ? item : NO_SOURCE,
					revision: part.paragraphs.markRevision[row] ?? 0
				},
				sink.nodes
			)
		)
	}

	private sdt(item: number, before: MarkerList): Node {
		const part = this.part
		const row = part.aux[item] ?? -1
		const sdtPr = part.sdts.sdtPr[row] ?? -1
		const lockNode = sdtPr < 0 ? -1 : part.reader.child(sdtPr, Ns.W, "lock")
		const lock = lockNode < 0 ? null : (part.reader.attr(lockNode, Ns.W, "val") ?? null)

		if (lock === "contentLocked" || lock === "sdtContentLocked")
			return this.locked(item, before, "sdt", "contentLocked", lock !== "sdtContentLocked")

		const content: Node[] = []
		const end = this.blocks(item, content)

		if (content.length === 0) return this.locked(item, before, "sdt", "opaque")

		return makeNode(docxSchema.nodes.sdt, { src: item, before, end, type: part.sdtType(row), lock }, content)
	}

	private table(item: number, before: MarkerList): Node {
		const part = this.part
		const saved = this.tableFailure
		const savedLocked = this.lockedBlocks
		const savedReasons = { ...this.lockReasons }
		const savedParagraphs = this.paragraphs

		const tableRow = part.aux[item] ?? -1

		this.tables++
		this.tableFailure =
			this.malformedWithin(part.tables.tblPr[tableRow] ?? -1) || this.malformedWithin(part.tables.tblGrid[tableRow] ?? -1)
				? "malformed"
				: null

		const rows: Node[] = []
		const end = part.end(item)
		let pending: MarkerAttrs[] | null = null

		for (let child = item + 1; child < end; child = part.end(child)) {
			const marker = this.marker(child)

			if (marker !== undefined) {
				;(pending ??= []).push(marker)
			} else if (part.kinds[child] === Item.Row) {
				rows.push(this.row(child, pending ?? NO_MARKERS))
				pending = null
			} else {
				this.tableFailure ??= "tableStructure"
			}
		}

		const failure = this.tableFailure ?? (rows.length === 0 ? "tableStructure" : null)

		this.tableFailure = saved

		if (failure !== null) {
			// What the rows built is discarded: the whole table is one locked block.
			this.lockedBlocks = savedLocked
			this.lockReasons = savedReasons
			this.paragraphs = savedParagraphs

			return this.locked(item, before, "table", failure)
		}

		const props = this.context.store.table(part.document, part.tables.tblPr[tableRow] ?? -1)

		return makeNode(
			docxSchema.nodes.table,
			{
				src: item,
				before,
				end: pending ?? NO_MARKERS,
				styleId: props.styleId ?? null,
				grid: this.grid(part.tables.tblGrid[tableRow] ?? -1)
			},
			rows
		)
	}

	private grid(tblGrid: number): readonly number[] {
		if (tblGrid < 0) return []

		const reader = this.part.reader
		const document = this.part.document
		const out: number[] = []

		for (let child = document.firstChildElement(tblGrid); child >= 0; child = document.nextSiblingElement(child)) {
			if (reader.is(child, Ns.W, "gridCol")) out.push(reader.int(child, Ns.W, "w", 0))
		}

		return out
	}

	private row(item: number, before: MarkerList): Node {
		const part = this.part
		const rowRow = part.aux[item] ?? -1

		if (this.malformedWithin(part.rows.trPr[rowRow] ?? -1) || this.malformedWithin(part.rows.tblPrEx[rowRow] ?? -1)) {
			this.tableFailure ??= "malformed"
		}

		const cells: Node[] = []
		const end = part.end(item)
		let pending: MarkerAttrs[] | null = null

		for (let child = item + 1; child < end; child = part.end(child)) {
			const marker = this.marker(child)

			if (marker !== undefined) {
				;(pending ??= []).push(marker)
			} else if (part.kinds[child] === Item.Cell) {
				cells.push(this.cell(child, pending ?? NO_MARKERS))
				pending = null
			} else {
				this.tableFailure ??= "tableStructure"
			}
		}

		if (cells.length === 0) {
			this.tableFailure ??= "tableStructure"

			return docxSchema.nodes.table_row.create()
		}

		return makeNode(
			docxSchema.nodes.table_row,
			{ src: item, before, end: pending ?? NO_MARKERS, revision: part.rows.revision[part.aux[item] ?? -1] ?? 0 },
			cells
		)
	}

	private cell(item: number, before: MarkerList): Node {
		const part = this.part
		const row = part.aux[item] ?? -1

		if (this.malformedWithin(part.cells.tcPr[row] ?? -1)) this.tableFailure ??= "malformed"

		const props = this.context.store.cell(part.document, part.cells.tcPr[row] ?? -1)
		const content: Node[] = []
		const end = this.blocks(item, content)

		if (content.length === 0) {
			this.tableFailure ??= "tableStructure"
			content.push(docxSchema.nodes.paragraph.create())
		}

		return makeNode(
			docxSchema.nodes.table_cell,
			{
				src: item,
				before,
				end,
				colspan: Math.max(1, props.gridSpan ?? 1),
				rowspan: 1,
				colwidth: null,
				vMerge: props.vMerge ?? null,
				revision: part.cells.revision[row] ?? 0
			},
			content
		)
	}

	// ---- inline content ---------------------------------------------------------------------------------

	private marksFor(wrappers: readonly Mark[], base: number, baseMarks: readonly Mark[]): readonly Mark[] {
		if (wrappers.length === 0) return baseMarks

		let byBase = this.combined.get(wrappers)

		if (byBase === undefined) {
			byBase = new Map()
			this.combined.set(wrappers, byBase)
		}

		let marks = byBase.get(base)

		if (marks === undefined) {
			marks = baseMarks

			for (const wrapper of wrappers) marks = wrapper.addToSet(marks)

			byBase.set(base, marks)
		}

		return marks
	}

	// The wrapper mark of a transparent inline container, or undefined for anything else.
	private wrapperMark(item: number): Mark | undefined {
		const part = this.part
		const model = this.context.model
		const kind = part.kinds[item] ?? 0
		const marks = docxSchema.marks
		const revision = REVISION_MARK_OF[kind]

		if (revision !== undefined) {
			const row = part.aux[item] ?? -1
			const id = part.revisions.id[row] ?? -1

			return docxSchema.marks[revision].create({
				src: item,
				id: id < 0 ? null : id,
				author: model.string(part.revisions.author[row] ?? -1) ?? null,
				date: model.string(part.revisions.date[row] ?? -1) ?? null
			})
		}

		switch (kind) {
			case Item.Hyperlink: {
				const target = part.hyperlink(item)

				return marks.link.create({
					src: item,
					anchor: target.anchor ?? null,
					target: target.target ?? null,
					tooltip: target.tooltip ?? null
				})
			}
			case Item.Sdt:
				return marks.content_control.create({ src: item, type: part.sdtType(part.aux[item] ?? -1) })
			case Item.Bidi:
				return marks.wrapper.create({ src: item, kind: "dir", rtl: (part.aux[item] ?? 0) === 1 })
		}

		const wrapper = WRAPPER_KINDS[kind]

		return wrapper === undefined ? undefined : marks.wrapper.create({ src: item, kind: wrapper })
	}

	// Paragraph-level children of `container`, with the w:proofErr elements between them.
	private inline(container: number, wrappers: readonly Mark[], sink: InlineSink): void {
		const part = this.part
		const nodes = docxSchema.nodes
		const end = part.end(container)
		const proofErrors = this.info.proofErrors.get(container)
		let pending = 0
		// Emits the proofErr elements placed before `item` (-1: all that are left); those inside an atom's
		// span are skipped, they are part of its bytes.
		const proofErrorsBefore = (item: number): void => {
			if (proofErrors === undefined) return

			while (pending < proofErrors.before.length) {
				const before = proofErrors.before[pending] ?? -1

				if (item >= 0 && (before < 0 || before > item)) return

				if (item < 0 || before === item)
					sink.add(makeNode(nodes.proof_err, { node: proofErrors.node[pending] ?? -1 }, null, wrappers))

				pending++
			}
		}

		for (let item = container + 1; item < end; item = part.end(item)) {
			const kind = part.kinds[item] ?? 0

			proofErrorsBefore(item)

			switch (kind) {
				case Item.Run: {
					const atomEnd = this.plan.atoms.get(item)

					if (atomEnd !== undefined) {
						sink.add(this.field(item, atomEnd, wrappers))
						item = atomEnd

						continue
					}

					this.run(item, wrappers, sink)

					continue
				}
				case Item.FieldSimple:
					sink.add(this.field(item, item, wrappers))

					continue
				case Item.Math:
				case Item.MathParagraph:
					sink.add(
						makeNode(
							nodes.math,
							{ src: item, display: kind === Item.MathParagraph, markers: this.survivors(item + 1, part.end(item)) },
							null,
							wrappers
						)
					)

					continue
				case Item.CommentReference:
					sink.add(makeNode(nodes.comment_ref, { src: item, id: part.id(item) ?? null }, null, wrappers))

					continue
			}

			const marker = this.marker(item)

			if (marker !== undefined) {
				sink.add(makeNode(nodes.marker, marker, null, wrappers))

				continue
			}

			const wrapper = this.wrapperMark(item)

			if (wrapper !== undefined) {
				this.inline(item, wrapper.addToSet(wrappers), sink)

				continue
			}

			sink.add(
				makeNode(
					nodes.opaque_inline,
					{
						src: item,
						kind: ITEM_NAMES[kind] ?? "Opaque",
						text: this.atomText(item),
						markers: this.survivors(item + 1, part.end(item))
					},
					null,
					wrappers
				)
			)
		}

		proofErrorsBefore(-1)
	}

	// What an atom kept as bytes shows as text: a positional tab's tab, the selected branch of an
	// mc:AlternateContent.
	private atomText(item: number): string {
		const part = this.part

		switch (part.kinds[item]) {
			case Item.PositionalTab:
				return "\t"
			case Item.AlternateContent:
				return part.paragraphText(item).slice(0, PREVIEW_LIMIT)
			default:
				return ""
		}
	}

	private field(first: number, last: number, wrappers: readonly Mark[]): Node {
		const part = this.part
		const simple = part.kinds[first] === Item.FieldSimple
		const field = simple ? (part.aux[first] ?? -1) : part.fieldOf(first + 1)

		return makeNode(
			docxSchema.nodes.field,
			{
				src: first,
				end: last,
				simple,
				type: field < 0 ? "" : part.fieldType(field),
				instruction: field < 0 ? "" : part.fieldInstruction(field),
				result: this.fieldResult(first, last, simple, field),
				markers: this.survivors(first, part.end(last))
			},
			null,
			wrappers
		)
	}

	// The cached result as text: what follows the separator (a simple field's content), field codes of
	// nested fields left out.
	private fieldResult(first: number, last: number, simple: boolean, field: number): string {
		const part = this.part
		const separate = simple || field < 0 ? first : (part.fields.separate[field] ?? -1)

		if (separate < 0) return ""

		const end = part.end(last)
		let out = ""
		let depth = 0
		const codes: boolean[] = []
		let code = false

		for (let item = separate + 1; item < end; item++) {
			switch (part.kinds[item]) {
				case Item.FieldChar: {
					const type = part.fieldCharType(item)

					if (type === FieldCharType.Begin) {
						codes.push(code)
						depth++
						code = true
					} else if (type === FieldCharType.Separate && depth > 0) {
						code = codes[codes.length - 1] ?? false
					} else if (type === FieldCharType.End && depth > 0) {
						depth--
						code = codes.pop() ?? false
					}

					break
				}
				case Item.Text:
				case Item.DeletedText:
					if (!code) out += part.text(item)

					break
				case Item.Tab:
					if (!code) out += "\t"

					break
				case Item.Break:
				case Item.CarriageReturn:
					if (!code) out += "\n"

					break
			}
		}

		return out
	}

	// A w:r: its content in order, each piece carrying the wrappers, the run's own mark and its base
	// formatting. Range markers Word put inside the run keep the wrappers and the run mark only.
	private run(item: number, wrappers: readonly Mark[], sink: InlineSink): void {
		const part = this.part
		const nodes = docxSchema.nodes
		const row = part.aux[item] ?? -1
		const base = this.context.bases.of(part.document, part.runs.rPr[row] ?? -1)
		const runMark = makeMark(docxSchema.marks.run, { src: item, base: base?.id ?? NO_SOURCE })
		const marks = runMark.addToSet(base === undefined ? wrappers : this.marksFor(wrappers, base.id, base.marks))
		const end = part.end(item)

		for (let child = item + 1; child < end; child = part.end(child)) {
			const kind = part.kinds[child] ?? 0

			switch (kind) {
				case Item.Text:
				case Item.DeletedText:
					sink.addText(this.runText(child), marks)

					break
				case Item.LastRenderedPageBreak:
					sink.add(makeNode(nodes.rendered_break, { src: child }, null, marks))

					break
				case Item.Tab:
					sink.add(makeNode(nodes.tab, { src: child }, null, marks))

					break
				case Item.Break:
				case Item.CarriageReturn:
					sink.add(
						makeNode(
							nodes.break,
							{
								src: child,
								type:
									kind === Item.CarriageReturn ? "textWrapping" : (BREAK_TYPES[part.breakType(child)] ?? "textWrapping"),
								clear: kind === Item.CarriageReturn ? "none" : (BREAK_CLEARS[part.breakClear(child)] ?? "none"),
								cr: kind === Item.CarriageReturn
							},
							null,
							marks
						)
					)

					break
				case Item.SoftHyphen:
				case Item.NoBreakHyphen:
					sink.add(makeNode(nodes.hyphen, { src: child, soft: kind === Item.SoftHyphen }, null, marks))

					break
				case Item.Symbol: {
					const symbol = part.symbol(child)

					sink.add(makeNode(nodes.symbol, { src: child, char: symbol.char, font: symbol.font ?? null }, null, marks))

					break
				}
				case Item.FootnoteReference:
				case Item.EndnoteReference:
					sink.add(
						makeNode(
							nodes.note_ref,
							{ src: child, kind: kind === Item.FootnoteReference ? "footnote" : "endnote", id: part.id(child) ?? null },
							null,
							marks
						)
					)

					break
				case Item.CommentReference:
					sink.add(makeNode(nodes.comment_ref, { src: child, id: part.id(child) ?? null }, null, marks))

					break
				case Item.Drawing:
				case Item.VmlPicture:
				case Item.EmbeddedObject:
				case Item.AlternateContent:
					sink.add(this.drawing(child, marks))

					break
				default: {
					const marker = kind === Item.Opaque ? this.marker(child) : undefined

					if (marker !== undefined) {
						sink.add(makeNode(nodes.marker, marker, null, runMark.addToSet(wrappers)))

						break
					}

					sink.add(
						makeNode(
							nodes.run_object,
							{
								src: child,
								kind: ITEM_NAMES[kind] ?? "Opaque",
								text: this.atomText(child),
								markers: this.survivors(child + 1, part.end(child))
							},
							null,
							marks
						)
					)
				}
			}
		}
	}

	// Text as Word reads it: without xml:space="preserve" (on the element or inherited), whitespace at the
	// edges is dropped (ECMA-376 §17.3.3.31 by way of XML 1.0 §2.10; XML whitespace only).
	private runText(item: number): string {
		const part = this.part
		const text = part.text(item)

		if (!EDGE_WHITESPACE.test(text) || part.preservesSpace(item) || this.inheritsPreserve(part.node(item))) return text

		return text.replace(TRIM, "")
	}

	private inheritsPreserve(node: number): boolean {
		const document = this.part.document
		const reader = this.part.reader

		for (let at = document.parent(node); at > 0; at = document.parent(at)) {
			const space = reader.attr(at, Ns.Xml, "space")

			if (space !== undefined) return space === "preserve"
		}

		return false
	}

	private drawing(item: number, marks: readonly Mark[]): Node {
		const part = this.part
		const end = part.end(item)
		let inner = -1

		for (let at = item; at < end && inner < 0; at++) {
			const kind = part.kinds[at]

			if (kind === Item.Drawing || kind === Item.VmlPicture || kind === Item.EmbeddedObject) inner = at
		}

		const markers = this.survivors(item + 1, end)

		if (inner < 0) {
			return makeNode(
				docxSchema.nodes.run_object,
				{ src: item, kind: ITEM_NAMES[part.kinds[item] ?? 0] ?? "Opaque", text: this.atomText(item), markers },
				null,
				marks
			)
		}

		const kind = part.kinds[inner]
		const index = part.aux[inner] ?? -1
		const anchored = kind === Item.Drawing && this.context.model.drawings[index]?.placement === "anchor"

		return makeNode(
			docxSchema.nodes.drawing,
			{ src: item, kind: kind === Item.Drawing ? "drawing" : kind === Item.VmlPicture ? "vml" : "object", index, anchored, markers },
			null,
			marks
		)
	}
}
