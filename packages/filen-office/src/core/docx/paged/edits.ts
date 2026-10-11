import type { Node } from "prosemirror-model"
import type { Transaction } from "prosemirror-state"
import {
	AddMarkStep,
	AddNodeMarkStep,
	AttrStep,
	RemoveMarkStep,
	RemoveNodeMarkStep,
	ReplaceAroundStep,
	ReplaceStep
} from "prosemirror-transform"
import { Relationships } from "../../opc/index"
import { encodeUtf8, parseXml, XmlError, type XmlDocument } from "../../xml/index"
import { parseNumbering, type Numbering } from "../numbering/index"
import { parseStyles, type DocxStyles } from "../styles/index"
import { DocxDocBuilder, docxSchema, type DocxEditContext, type StoryDoc } from "../edit/index"
import type { NewList } from "../edit/formats"
import { ref, refItem, refPart, type BodyBlocks, type DirtyBlock, type LayoutEdits } from "../layout/index"
import {
	Item,
	ModelTables,
	RevisionView,
	buildPart,
	partReader,
	type DocxModel,
	type ModelExtension,
	type Story,
	type StoryPart
} from "../model/index"
import { DocxWriteError, StoryWriter, planNumbering, stylesWithDefinitions, type WriteEnvironment } from "../write/index"

// What the layout lays out while a document is edited, kept in step with the editing sessions: every
// edited block is written as it will be saved (StoryWriter), parsed again into a small overlay part and laid
// out in place of its source ("lay out what you save"); blocks added, removed or moved change the body's
// block list. Pages map back to the editing document through the same records.

const nodes = docxSchema.nodes
const HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

// The styles and numbering parts as opened (the planned definitions are added to them as a save adds them).
export interface EditDefinitions {
	readonly context: DocxEditContext
	readonly styles: XmlDocument | undefined
	readonly numbering: XmlDocument | undefined
}

// A block laid out from an overlay part.
export interface OverlayRecord {
	readonly part: StoryPart
	// The block (paragraph or table) in the overlay part.
	readonly item: number
	// The source block it stands for (a ref), or -1 for a block the editor created.
	readonly base: number
	readonly node: Node
	readonly story: StoryEdits
	// The top-level entry holding the block, and its position inside the entry's node.
	entry: TopEntry
	offset: number
	// The block as the editor builds it from the overlay part (lazily; maps pages to the editing document).
	shadow: Node | null | undefined
}

// An overlay part: the blocks one edit wrote, parsed together. It lives while a block of it is laid out.
interface OverlayPart {
	// By item, ascending.
	readonly records: OverlayRecord[]
	// What the part's build added to the model's lists, dropped with it.
	readonly stories: readonly [number, number]
	readonly drawings: readonly [number, number]
	readonly vml: readonly [number, number]
}

// A top-level block of a story's editing document and the blocks it is laid out as.
export interface TopEntry {
	node: Node
	index: number
	section: number
	// Refs as the layout addresses them before substitution (the story's block list), refs as laid out, and
	// per ref the offset of the editing node it comes from inside `node` (-1 inside a locked block).
	readonly base: number[]
	readonly laid: number[]
	readonly offsets: number[]
	// Overlay records and substituted source refs this entry owns.
	readonly records: OverlayRecord[]
	readonly substituted: number[]
}

// The part of an edit the layout must redo: the changed top-level blocks as laid out now, the entries the
// edit replaced and, when blocks came or went, their neighbours (layout restarts at the first page showing
// any of those).
export interface StoryChange {
	readonly dirty: DirtyBlock[]
	readonly replaced: readonly TopEntry[]
	readonly around: readonly TopEntry[]
	// The section properties or the story's own structure changed beyond what the dirty blocks cover: lay
	// out from the story's start.
	readonly whole: boolean
}

// One document's edits for layout: the derived model the layout reads (overlay parts added), the
// substitutions and block lists of every edited story.
export class DocumentEdits implements LayoutEdits {
	readonly model: DocxModel
	readonly extension: ModelExtension
	readonly substitutes = new Map<number, number>()
	readonly view: RevisionView
	body: BodyBlocks | undefined = undefined
	// Stories other than the body whose block structure changed: their blocks (refs).
	readonly stories = new Map<number, Float64Array>()
	private readonly overlays = new Map<number, OverlayPart>()
	// Stories of the document itself; those after them belong to overlay parts.
	readonly storyCount: number
	private readonly bodyIndex: number
	private readonly bodyParts = new Set<number>()
	// The w:numId each list a command started will get when saved (the overlays write them).
	numIds: ReadonlyMap<number, number> = new Map()
	private definitions: EditDefinitions | undefined = undefined
	private plannedKey = ""
	// Relationship ids of hyperlinks commands inserted, per source part, and the parts' relationships with them.
	private readonly links = new Map<string, Map<string, string>>()
	private readonly relationships = new Map<string, Relationships>()

	constructor(model: DocxModel, view: RevisionView = RevisionView.Final) {
		const derived = model.derive()

		this.model = derived.model
		this.extension = derived.extension
		this.view = view
		this.storyCount = model.stories.length
		this.bodyIndex = model.body?.part ?? -1
	}

	// The parts a save rewrites when commands added lists or styles: the overlays and the layout use what
	// the save will write.
	define(definitions: EditDefinitions): void {
		this.definitions = definitions
	}

	// Plans the definitions for `lists` and `missing` (the sessions' new lists and styles); the styles and
	// numbering to lay out with when they changed, else undefined.
	plan(
		lists: ReadonlyMap<number, NewList>,
		missing: ReadonlyMap<string, string>
	): { readonly styles: DocxStyles; readonly numbering: Numbering | undefined } | undefined {
		const definitions = this.definitions

		if (definitions === undefined) return undefined

		const key =
			[...lists].map(([id, list]) => `${id}:${list.kind}:${String(list.basedOn)}`).join(",") + "|" + [...missing.keys()].join(",")

		if (key === this.plannedKey) return undefined

		this.plannedKey = key

		const context = definitions.context
		const stylesBytes = missing.size === 0 ? null : stylesWithDefinitions(definitions.styles, context.styles, missing)
		const styles = stylesBytes === null ? context.styles : parseStyles(parseXml(stylesBytes), context.store)
		const numbering = planNumbering(definitions.numbering, context.numbering, lists)
		const numberingBytes = numbering.bytes()

		this.numIds = numbering.numIds

		return {
			styles,
			numbering: numberingBytes === null ? context.numbering : parseNumbering(parseXml(numberingBytes), context.store, styles)
		}
	}

	// The relationship id overlays write for a hyperlink a command inserted (the save assigns its own; the
	// target is the same).
	linkId(partName: string, target: string): string {
		let ids = this.links.get(partName)

		if (ids === undefined) {
			ids = new Map()
			this.links.set(partName, ids)
		}

		let id = ids.get(target)

		if (id === undefined) {
			id = `filenPlannedLink${String(ids.size + 1)}`
			ids.set(target, id)
			this.relationships.delete(partName)
		}

		return id
	}

	// A part's relationships with the hyperlinks commands added.
	private relationshipsOf(source: StoryPart): Relationships {
		const ids = this.links.get(source.name)

		if (ids === undefined) return source.relationships

		let relationships = this.relationships.get(source.name)

		if (relationships === undefined) {
			const added = [...ids].map(([target, id]) => ({ id, type: HYPERLINK, target, external: true, partName: undefined }))

			relationships = new Relationships([...source.relationships.all, ...added])
			this.relationships.set(source.name, relationships)
		}

		return relationships
	}

	memoKey(part: number, item: number): number {
		const record = this.recordAt(part, item)

		return record !== undefined && record.base >= 0 && record.item === item ? record.base : ref(part, item)
	}

	bodyPart(part: number): boolean {
		return part === this.bodyIndex || this.bodyParts.has(part)
	}

	part(index: number): StoryPart | undefined {
		return this.extension.parts[index]
	}

	// Parses `bytes` (blocks written by StoryWriter) inside the root of `source` into a new overlay part, and
	// its top-level paragraphs and tables in order; undefined when it does not parse.
	addPart(
		source: StoryPart,
		container: number,
		bytes: Uint8Array,
		body: boolean
	): { readonly part: StoryPart; readonly items: readonly number[] } | undefined {
		const document = source.document
		const root = document.documentElement()
		const node = source.node(container)
		const open = document.bytes.subarray(document.start(root), document.openEnd(root))
		// A body block goes into a w:body; the content of a header or footer sits in the root itself.
		const inner =
			node !== root && node >= 0
				? { open: encodeUtf8(`<${document.name(node)}>`), close: encodeUtf8(`</${document.name(node)}>`) }
				: { open: new Uint8Array(0), close: new Uint8Array(0) }
		const close = encodeUtf8(`</${document.name(root)}>`)
		const wrapped = new Uint8Array(open.length + inner.open.length + bytes.length + inner.close.length + close.length)
		let at = 0

		for (const chunk of [open, inner.open, bytes, inner.close, close]) {
			wrapped.set(chunk, at)
			at += chunk.length
		}

		let parsed

		try {
			parsed = parseXml(wrapped)
		} catch (error) {
			if (error instanceof XmlError) return undefined

			throw error
		}

		const extension = this.extension
		const index = extension.parts.length
		const stories = extension.stories.length
		const drawings = extension.drawings.length
		const vml = extension.vml.length
		const tables = Object.assign(new ModelTables(), { stories: extension.stories, drawings: extension.drawings, vml: extension.vml })
		const part = buildPart(partReader(parsed), tables as ModelTables, index, source.name, this.relationshipsOf(source))

		extension.parts.push(part)

		if (body) this.bodyParts.add(index)

		this.overlays.set(index, {
			records: [],
			stories: [stories, extension.stories.length],
			drawings: [drawings, extension.drawings.length],
			vml: [vml, extension.vml.length]
		})

		const story = extension.stories[stories]
		const items: number[] = []

		if (story !== undefined) {
			for (let child = story.item + 1; child < part.end(story.item); child = part.end(child)) {
				const kind = part.kinds[child]

				if (kind === Item.Paragraph || kind === Item.Table) items.push(child)
			}
		}

		return { part, items }
	}

	// The record of the overlay block holding item `item` of part `part`, or undefined.
	recordAt(part: number, item: number): OverlayRecord | undefined {
		const records = this.overlays.get(part)?.records

		if (records === undefined || records.length === 0) return undefined

		const record = records[recordIndex(records, item)]

		return record !== undefined && record.item <= item && item < record.part.end(record.item) ? record : undefined
	}

	// Every block laid out from an overlay part.
	*overlayRecords(): Generator<OverlayRecord> {
		for (const overlay of this.overlays.values()) yield* overlay.records
	}

	addRecord(record: OverlayRecord): void {
		const records = this.overlays.get(record.part.index)?.records

		if (records === undefined) return

		// A part's blocks are added in item order: this finds the end at once.
		let at = records.length

		while (at > 0 && (records[at - 1]?.item ?? 0) > record.item) at--

		records.splice(at, 0, record)
	}

	// Drops a block laid out from an overlay part, and the part with its last block.
	dropRecord(record: OverlayRecord): void {
		const index = record.part.index
		const overlay = this.overlays.get(index)

		if (overlay === undefined) return

		const records = overlay.records
		const at = recordIndex(records, record.item)

		if (records[at] === record) records.splice(at, 1)

		this.dropUnused(record.part)
	}

	// Drops an overlay part no block is laid out from, and what its build added to the model's lists.
	dropUnused(part: StoryPart): void {
		const index = part.index
		const overlay = this.overlays.get(index)

		if (overlay === undefined || overlay.records.length > 0) return

		const extension = this.extension

		extension.parts[index] = undefined
		this.bodyParts.delete(index)
		this.overlays.delete(index)

		for (let k = overlay.stories[0]; k < overlay.stories[1]; k++) extension.stories[k] = undefined
		for (let k = overlay.drawings[0]; k < overlay.drawings[1]; k++) extension.drawings[k] = undefined
		for (let k = overlay.vml[0]; k < overlay.vml[1]; k++) extension.vml[k] = undefined
	}

	// The blocks the layout's flow of a source container (a story root, a cell, a locked wrapper) finds in
	// [start, end) of `part`: paragraphs and tables, wrappers entered as layout enters them.
	expand(part: StoryPart, start: number, end: number, out: number[]): void {
		const kinds = part.kinds
		const view = this.view

		for (let item = start; item < end;) {
			const kind = kinds[item] ?? 0

			if (kind === Item.Paragraph || kind === Item.Table) {
				out.push(ref(part.index, item))
				item = part.end(item)
			} else if (
				kind === Item.Sdt ||
				kind === Item.CustomXml ||
				kind === Item.AlternateContent ||
				((kind === Item.Insertion || kind === Item.MoveTo) && view !== RevisionView.Original) ||
				((kind === Item.Deletion || kind === Item.MoveFrom) && view !== RevisionView.Final)
			) {
				item++
			} else {
				item = part.end(item)
			}
		}
	}
}

// The edits of one story (the body, a header, a footer, a note) for layout, and the index from its editing
// document to what the layout laid out.
export class StoryEdits {
	readonly edits: DocumentEdits
	readonly story: StoryDoc
	readonly context: DocxEditContext
	private readonly writer: StoryWriter
	private readonly body: boolean
	readonly entries: TopEntry[] = []
	// Start position of each entry's node in the document.
	starts = new Int32Array(0)
	private doc: Node
	// Whether the story's block structure differs from its part's (the layout then reads a block list).
	structured = false
	// Entries by the source items they lay out, ascending (rebuilt when the structure changes).
	private sourceFrom = new Int32Array(0)
	private sourceTo = new Int32Array(0)
	private sourceEntry = new Int32Array(0)
	private sourcesValid = false
	// The entry laying out each source block (ref): when an edit leaves two blocks with the same source (a
	// split paragraph keeps its identity in both halves), the first owns it and the other is a new block.
	private readonly owners = new Map<number, TopEntry>()
	// Entries built by the edit being applied, not in `entries` yet.
	private readonly building = new Set<TopEntry>()
	// Tables of the document as opened (the builder records paragraphs only): written as their source.
	private readonly openedTables = new WeakSet<Node>()
	// Blocks the edit being applied wrote, parsed into one overlay part when it is done (place()).
	private readonly queued: QueuedBlock[] = []
	private shadows: DocxDocBuilder | undefined = undefined

	// `doc` is the session's document now (blocks the proof locked replaced).
	constructor(edits: DocumentEdits, story: StoryDoc, doc: Node, context: DocxEditContext, env?: WriteEnvironment) {
		this.edits = edits
		this.story = story
		this.context = context
		this.body = story.story.kind === "body"

		let revision = Math.max(0, context.model.maxIds.revision)

		this.writer = new StoryWriter(
			env ?? {
				context,
				nextRevisionId: () => ++revision,
				numId: id => {
					const numId = edits.numIds.get(id)

					if (numId === undefined) throw new DocxWriteError("missingSource", `Unknown new list ${String(id)}`)

					return numId
				},
				hyperlink: (part, target) => edits.linkId(part, target)
			},
			story
		)
		this.doc = doc
		doc.descendants(node => {
			if (node.type === nodes.table) this.openedTables.add(node)

			return !node.isTextblock
		})
		this.rebuildAll(doc)
	}

	get part(): StoryPart {
		return this.story.part
	}

	get document(): Node {
		return this.doc
	}

	// ---- keeping up with the editing document -------------------------------------------------------------

	// Brings the layout's view of the story up to date with transactions the session applied, in order;
	// undefined when nothing the layout shows changed.
	applyAll(transactions: readonly Transaction[]): StoryChange | undefined {
		let out: StoryChange | undefined

		for (const tr of transactions) {
			const change = this.apply(tr)

			if (change === undefined) continue

			out =
				out === undefined
					? change
					: {
							dirty: out.dirty.concat(change.dirty),
							replaced: out.replaced.concat(change.replaced),
							around: out.around.concat(change.around),
							whole: out.whole || change.whole
						}
		}

		return out
	}

	apply(tr: Transaction): StoryChange | undefined {
		if (tr.steps.length === 0) return undefined

		const before = this.doc
		const after = tr.doc
		const span = before === tr.before ? changedSpan(tr) : { from: 0, to: after.content.size, whole: true }

		this.doc = after

		if (span === undefined) return undefined

		if (span.whole || before.attrs["sectPr"] !== after.attrs["sectPr"]) {
			const replaced = this.entries.slice()

			this.rebuildAll(after)

			return { dirty: this.dirtyOf(0, this.entries.length - 1), replaced, around: [], whole: true }
		}

		const oldCount = before.childCount
		const newCount = after.childCount
		let first = childAt(after, span.from, false)
		let last = Math.max(first - 1, childAt(after, span.to, true))

		// Unchanged neighbours are the same node objects: widen until they are.
		while (first > 0 && after.child(first - 1) !== before.child(first - 1)) first--

		let oldLast = last + (oldCount - newCount)

		while (last + 1 < newCount && oldLast + 1 < oldCount && after.child(last + 1) !== before.child(oldLast + 1)) {
			last++
			oldLast++
		}

		if (oldLast < first - 1 || last < first - 1 || oldLast >= oldCount) {
			const replaced = this.entries.slice()

			this.rebuildAll(after)

			return { dirty: this.dirtyOf(0, this.entries.length - 1), replaced, around: [], whole: true }
		}

		const removed = this.entries.slice(first, oldLast + 1)
		const reuse = new Map<Node, TopEntry>()
		const children = new Set<Node>()

		for (const entry of removed) reuse.set(entry.node, entry)
		for (let k = first; k <= last; k++) children.add(after.child(k))

		// Entries not taken over by a node of the new document give up their overlays at the end.
		const pending = new Set(removed)
		const fresh: TopEntry[] = []
		let structural = last - first !== oldLast - first

		for (let k = first; k <= last; k++) {
			const node = after.child(k)
			const kept = reuse.get(node)

			if (kept !== undefined) {
				reuse.delete(node)
				pending.delete(kept)
				kept.index = k
				fresh.push(kept)

				if (removed[k - first] !== kept) structural = true

				continue
			}

			// The entry that stood here is the block's previous version unless its node lives on elsewhere.
			const candidate = removed[k - first]
			const old = candidate !== undefined && pending.has(candidate) && !children.has(candidate.node) ? candidate : undefined

			if (old !== undefined) pending.delete(old)

			const entry = this.entry(node, k, old)

			if (old === undefined || !sameRefs(old.base, entry.base)) structural = true

			fresh.push(entry)
		}

		this.place()

		for (const entry of pending) this.release(entry)

		this.entries.splice(first, oldLast - first + 1, ...fresh)
		this.building.clear()

		const delta = after.content.size - before.content.size
		// Section boundaries move only when a block carrying section properties came or went (or gained or
		// lost them).
		const sections = removed.some(entry => endsSection(entry.node)) || fresh.some(entry => endsSection(entry.node))

		if (structural || newCount !== oldCount || sections) {
			for (let k = first; k < this.entries.length; k++) {
				const entry = this.entries[k]

				if (entry !== undefined) entry.index = k
			}

			this.reposition(after, first)
			this.structured = true
			this.sourcesValid = false

			if (sections) {
				this.renumberSections()
				this.publish()
			} else {
				// The section of the blocks replaced; for blocks added, that of the block before (or the next
				// section when that block ends its own).
				const previous = this.entries[first - 1]
				const section =
					removed[0]?.section ?? (previous === undefined ? 0 : previous.section + (endsSection(previous.node) ? 1 : 0))

				for (const entry of fresh) entry.section = section

				this.publish(first, removed, fresh)
			}
		} else {
			for (let k = last + 1; k < this.entries.length; k++) this.starts[k] = (this.starts[k] ?? 0) + delta
		}

		// A removal dirties its neighbours: what comes after it moves up.
		const from = last >= first ? first : Math.max(0, first - 1)
		const to = last >= first ? last : Math.min(this.entries.length - 1, first)
		const around: TopEntry[] = []

		if (structural) {
			const before = this.entries[first - 1]
			const after = this.entries[first + fresh.length]

			if (before !== undefined) around.push(before)
			if (after !== undefined) around.push(after)
		}

		return { dirty: this.dirtyOf(from, to), replaced: removed, around, whole: false }
	}

	private dirtyOf(from: number, to: number): DirtyBlock[] {
		const out: DirtyBlock[] = []

		for (let k = Math.max(0, from); k <= to && k < this.entries.length; k++) {
			const entry = this.entries[k]

			if (entry === undefined) continue

			for (const laid of entry.laid) out.push({ section: entry.section, ref: laid })
		}

		return out
	}

	private rebuildAll(doc: Node): void {
		for (const entry of this.entries) this.release(entry)

		this.owners.clear()
		this.entries.length = 0
		doc.forEach((node, _offset, index) => {
			this.entries.push(this.entry(node, index, undefined))
		})
		this.place()
		this.building.clear()
		this.reposition(doc)
		this.sourcesValid = false
		this.structured = !this.matchesSource()
		this.renumberSections()
		this.publish()
	}

	// Start positions of the entries from `from` on (those before it are unchanged).
	private reposition(doc: Node, from = 0): void {
		const count = doc.childCount
		const starts = this.starts.length === count ? this.starts : new Int32Array(count)

		if (starts !== this.starts) starts.set(this.starts.subarray(0, Math.min(from, this.starts.length, count)))

		let pos = from === 0 ? 0 : (starts[from - 1] ?? 0) + doc.child(from - 1).nodeSize

		for (let k = from; k < count; k++) {
			starts[k] = pos
			pos += doc.child(k).nodeSize
		}

		this.starts = starts
	}

	// Whether the story's blocks are exactly what the layout reads from its part.
	private matchesSource(): boolean {
		const part = this.part
		const expected: number[] = []

		this.edits.expand(part, this.story.story.item + 1, part.end(this.story.story.item), expected)

		let at = 0

		for (const entry of this.entries) {
			for (const value of entry.base) if (expected[at++] !== value) return false
		}

		return at === expected.length
	}

	// Section of each body entry: sections end with the block carrying their properties.
	private renumberSections(): void {
		if (!this.body) return

		const sections = this.edits.model.sections
		let section = 0

		for (const entry of this.entries) {
			entry.section = Math.min(section, Math.max(0, sections.length - 1))

			if (endsSection(entry.node)) section++
		}
	}

	// Hands the story's block list to the layout once its structure is its own. With the change (entries
	// `fresh` at `first` replacing `removed`, sections unchanged) the previous list is patched instead.
	private publish(first?: number, removed?: readonly TopEntry[], fresh?: readonly TopEntry[]): void {
		if (!this.structured) {
			if (this.body) this.edits.body = undefined
			else this.edits.stories.delete(this.story.story.index)

			return
		}

		const previous = this.body ? this.edits.body : this.edits.stories.get(this.story.story.index)
		const previousRefs = previous === undefined ? undefined : "refs" in previous ? previous.refs : previous

		if (first !== undefined && removed !== undefined && fresh !== undefined && previousRefs !== undefined) {
			let at = 0

			for (let k = 0; k < first; k++) at += this.entries[k]?.base.length ?? 0

			let gone = 0
			let added = 0

			for (const entry of removed) gone += entry.base.length
			for (const entry of fresh) added += entry.base.length

			const refs = new Float64Array(previousRefs.length - gone + added)

			refs.set(previousRefs.subarray(0, at))

			let write = at

			for (const entry of fresh) for (const value of entry.base) refs[write++] = value

			refs.set(previousRefs.subarray(at + gone), write)

			if (!this.body) {
				this.edits.stories.set(this.story.story.index, refs)

				return
			}

			const ends = (previous !== undefined && "sectionEnds" in previous ? previous.sectionEnds : new Int32Array(1)).slice()

			for (let k = 0; k < ends.length; k++) if ((ends[k] ?? 0) > at) ends[k] = (ends[k] ?? 0) - gone + added

			this.edits.body = { refs, count: refs.length, sectionEnds: ends }

			return
		}

		let count = 0

		for (const entry of this.entries) count += entry.base.length

		const refs = new Float64Array(count)
		let at = 0

		for (const entry of this.entries) for (const value of entry.base) refs[at++] = value

		if (!this.body) {
			this.edits.stories.set(this.story.story.index, refs)

			return
		}

		const sections = this.edits.model.sections
		const ends = new Int32Array(Math.max(1, sections.length))
		let section = 0

		at = 0

		for (const entry of this.entries) {
			at += entry.base.length

			if (endsSection(entry.node) && section < ends.length - 1) ends[section++] = at
		}

		while (section < ends.length) ends[section++] = at

		this.edits.body = { refs, count, sectionEnds: ends }
	}

	// ---- entries ------------------------------------------------------------------------------------------

	private entry(node: Node, index: number, old: TopEntry | undefined): TopEntry {
		const entry: TopEntry = {
			node,
			index,
			section: old?.section ?? 0,
			base: [],
			laid: [],
			offsets: [],
			records: [],
			substituted: []
		}

		this.building.add(entry)

		this.block(entry, node, 0, old)

		if (old !== undefined) this.release(old)

		return entry
	}

	private block(entry: TopEntry, node: Node, offset: number, old: TopEntry | undefined): void {
		const part = this.part
		const src = (node.attrs["src"] as number | undefined) ?? -1
		const pristine = this.story.pristine.has(node) || this.openedTables.has(node)

		switch (node.type) {
			case nodes.paragraph:
			case nodes.table: {
				const base = src >= 0 && this.claim(ref(part.index, src), entry, old) ? ref(part.index, src) : -1

				if (base >= 0 && pristine) {
					this.source(entry, base, offset)

					return
				}

				if (node.type === nodes.table && old !== undefined && this.cellEdits(entry, node, offset, base, old)) return

				this.overlay(entry, node, offset, base)

				return
			}
			case nodes.sdt:
				node.forEach((child, childOffset) => this.block(entry, child, offset + 1 + childOffset, old))

				return
			default: {
				// A locked block: what layout finds inside its element, never edited.
				if (src < 0) return

				const refs: number[] = []

				this.edits.expand(part, src, part.end(src), refs)

				for (const value of refs) {
					entry.base.push(value)
					entry.laid.push(value)
					entry.offsets.push(-1)
				}
			}
		}
	}

	// Whether `entry` (replacing `old`) lays out source block `base`: no other live entry does.
	private claim(base: number, entry: TopEntry, old: TopEntry | undefined): boolean {
		const owner = this.owners.get(base)

		if (owner !== undefined && owner !== old && owner !== entry && (this.entries[owner.index] === owner || this.building.has(owner)))
			return false

		this.owners.set(base, entry)

		return true
	}

	// A source block laid out as it is.
	private source(entry: TopEntry, value: number, offset: number): void {
		const substitutes = this.edits.substitutes
		const substitute = substitutes.get(value)

		// An earlier edit's overlay is dropped: the block is its source again (undo back to the original).
		if (substitute !== undefined) {
			substitutes.delete(value)

			const record = this.edits.recordAt(refPart(substitute), refItem(substitute))

			if (record !== undefined) this.edits.dropRecord(record)
		}

		entry.base.push(value)
		entry.laid.push(value)
		entry.offsets.push(offset)
	}

	// A table whose rows and cells are as they were, with paragraphs edited inside: those paragraphs are laid
	// out from overlays, the table from its source, or from its own overlay when an edit changed its
	// structure (or the editor made it). False when its structure changed now.
	private cellEdits(entry: TopEntry, node: Node, offset: number, base: number, old: TopEntry): boolean {
		const previous = old.node.type === nodes.table && old.node.attrs["src"] === node.attrs["src"] ? old.node : undefined

		if (previous === undefined) return false

		const part = this.part
		const substitute = base >= 0 ? this.edits.substitutes.get(base) : old.laid[0]
		// The table's own overlay, as last written.
		const table =
			substitute === undefined
				? undefined
				: old.records.find(record => record.node.type === nodes.table && ref(record.part.index, record.item) === substitute)
		const changed: ChangedParagraph[] = []

		if (table === undefined) {
			if (base < 0 || substitute !== undefined || !sameStructure(previous, node)) return false

			changedParagraphs(previous, node, offset, [], changed)
		} else {
			if (table.offset !== offset || !sameShape(table.node, node) || !sameShape(previous, node)) return false

			changedParagraphs(previous, node, offset, [], changed)

			const shadow = this.shadowOf(table)

			if (shadow === null) return false

			// Paragraphs of the overlay table are those at the same place in the table it was written from.
			for (const paragraph of changed) {
				const src = atPath(shadow, paragraph.path)?.attrs["src"]

				if (typeof src !== "number" || src < 0) return false

				paragraph.base = ref(table.part.index, src)
			}
		}

		// Substitutions of paragraphs that did not change carry over.
		for (const kept of old.substituted) entry.substituted.push(kept)
		for (const record of old.records) entry.records.push(record)

		old.substituted.length = 0
		old.records.length = 0

		for (const record of entry.records) record.entry = entry

		for (const paragraph of changed) {
			const value = paragraph.base >= 0 ? paragraph.base : ref(part.index, paragraph.node.attrs["src"] as number)
			const existing = this.edits.substitutes.get(value)

			if (existing !== undefined) {
				const record = this.edits.recordAt(refPart(existing), refItem(existing))

				if (record !== undefined) {
					this.edits.dropRecord(record)
					entry.records.splice(entry.records.indexOf(record), 1)
				}
			} else {
				entry.substituted.push(value)
			}

			const bytes = this.written(paragraph.node)

			if (bytes === undefined) this.edits.substitutes.delete(value)
			else this.queued.push({ entry, node: paragraph.node, offset: paragraph.offset, base: value, bytes, slot: -1 })
		}

		for (const record of entry.records) {
			if (record !== table && !changed.some(paragraph => paragraph.node === record.node))
				record.offset = offsetOf(node, offset, record.node) ?? record.offset
		}

		const laid = table === undefined ? base : ref(table.part.index, table.item)

		entry.base.push(base >= 0 ? base : laid)
		entry.laid.push(laid)
		entry.offsets.push(offset)

		return true
	}

	// A block laid out from an overlay of what it writes as (once the edit's blocks are parsed, see place()).
	private overlay(entry: TopEntry, node: Node, offset: number, base: number): void {
		const bytes = this.written(node)

		// What the writer refuses (the save refuses it too) is laid out as its source.
		if (bytes === undefined) {
			if (base >= 0) {
				this.edits.substitutes.delete(base)
				entry.base.push(base)
				entry.laid.push(base)
				entry.offsets.push(offset)
			}

			return
		}

		if (base >= 0) entry.substituted.push(base)

		// Placeholders until placed.
		entry.base.push(base)
		entry.laid.push(base)
		entry.offsets.push(offset)
		this.queued.push({ entry, node, offset, base, bytes, slot: entry.laid.length - 1 })
	}

	private written(node: Node): Uint8Array | undefined {
		try {
			return this.writer.writeBlockAlone(node).finish()
		} catch (error) {
			if (error instanceof DocxWriteError) return undefined

			throw error
		}
	}

	// Parses the blocks the edit wrote into one overlay part and lays them out from it. A block the batch
	// cannot place (its output does not parse) is laid out as its source, or left out when it has none.
	private place(): void {
		if (this.queued.length === 0) return

		const blocks = this.queued.splice(0)
		let size = 0

		for (const block of blocks) size += block.bytes.length

		const bytes = new Uint8Array(size)
		let at = 0

		for (const block of blocks) {
			bytes.set(block.bytes, at)
			at += block.bytes.length
		}

		const added = this.edits.addPart(this.part, this.story.story.item, bytes, this.body)

		if (added !== undefined && added.items.length === blocks.length) {
			blocks.forEach((block, k) => this.placed(block, added.part, added.items[k] ?? -1))

			return
		}

		if (added !== undefined) this.edits.dropUnused(added.part)

		const unplaced = new Map<TopEntry, Set<number>>()

		for (const block of blocks) {
			const one = blocks.length === 1 ? added : this.edits.addPart(this.part, this.story.story.item, block.bytes, this.body)
			const item = one?.items[0]

			if (one !== undefined && item !== undefined) {
				this.placed(block, one.part, item)

				continue
			}

			if (one !== undefined) this.edits.dropUnused(one.part)

			if (block.base >= 0) {
				this.edits.substitutes.delete(block.base)
			} else if (block.slot >= 0) {
				const slots = unplaced.get(block.entry) ?? new Set<number>()

				slots.add(block.slot)
				unplaced.set(block.entry, slots)
			}
		}

		for (const [entry, slots] of unplaced) {
			const keep = (_: number, k: number): boolean => !slots.has(k)
			const base = entry.base.filter(keep)
			const laid = entry.laid.filter(keep)
			const offsets = entry.offsets.filter(keep)

			entry.base.splice(0, entry.base.length, ...base)
			entry.laid.splice(0, entry.laid.length, ...laid)
			entry.offsets.splice(0, entry.offsets.length, ...offsets)
		}
	}

	private placed(block: QueuedBlock, part: StoryPart, item: number): void {
		const record: OverlayRecord = {
			part,
			item,
			base: block.base,
			node: block.node,
			story: this,
			entry: block.entry,
			offset: block.offset,
			shadow: undefined
		}
		const laid = ref(part.index, item)

		this.edits.addRecord(record)
		block.entry.records.push(record)

		if (block.base >= 0) this.edits.substitutes.set(block.base, laid)

		if (block.slot >= 0) {
			block.entry.laid[block.slot] = laid

			if (block.base < 0) block.entry.base[block.slot] = laid
		}
	}

	private release(entry: TopEntry): void {
		for (const base of entry.base) if (this.owners.get(base) === entry) this.owners.delete(base)
		for (const record of entry.records) this.edits.dropRecord(record)
		for (const base of entry.substituted) {
			const substitute = this.edits.substitutes.get(base)

			if (substitute !== undefined && this.edits.recordAt(refPart(substitute), refItem(substitute)) === undefined)
				this.edits.substitutes.delete(base)
		}

		entry.records.length = 0
		entry.substituted.length = 0
	}

	// ---- lookups -------------------------------------------------------------------------------------------

	// The entry whose node holds position `pos`.
	entryAt(pos: number): number {
		const starts = this.starts
		let low = 0
		let high = this.entries.length - 1

		while (low < high) {
			const middle = (low + high + 1) >> 1

			if ((starts[middle] ?? 0) <= pos) low = middle
			else high = middle - 1
		}

		return low
	}

	// The entry a laid-out source (part, item) belongs to, or undefined.
	entryOf(part: number, item: number): TopEntry | undefined {
		const record = this.edits.recordAt(part, item)

		if (record !== undefined) return record.story === this ? record.entry : undefined

		if (part !== this.part.index) return undefined

		this.indexSources()

		const from = this.sourceFrom
		let low = 0
		let high = from.length - 1

		if (high < 0 || item < (from[0] ?? 0)) return undefined

		while (low < high) {
			const middle = (low + high + 1) >> 1

			if ((from[middle] ?? 0) <= item) low = middle
			else high = middle - 1
		}

		if (item >= (this.sourceTo[low] ?? 0)) return undefined

		return this.entries[this.sourceEntry[low] ?? -1]
	}

	private indexSources(): void {
		if (this.sourcesValid) return

		const part = this.part
		let count = 0

		for (const entry of this.entries) for (const value of entry.base) if (refPart(value) === part.index) count++

		const from = new Int32Array(count)
		const to = new Int32Array(count)
		const owner = new Int32Array(count)
		let at = 0
		let sorted = true

		this.entries.forEach((entry, index) => {
			for (const value of entry.base) {
				if (refPart(value) !== part.index) continue

				const item = refItem(value)

				if (at > 0 && item < (from[at - 1] ?? 0)) sorted = false

				from[at] = item
				to[at] = part.end(item)
				owner[at] = index
				at++
			}
		})

		// Blocks moved out of document order (dragged paragraphs) need sorting; most documents never do.
		if (!sorted) {
			const order = Array.from({ length: count }, (_, k) => k).sort((a, b) => (from[a] ?? 0) - (from[b] ?? 0))

			this.sourceFrom = Int32Array.from(order, k => from[k] ?? 0)
			this.sourceTo = Int32Array.from(order, k => to[k] ?? 0)
			this.sourceEntry = Int32Array.from(order, k => owner[k] ?? 0)
		} else {
			this.sourceFrom = from
			this.sourceTo = to
			this.sourceEntry = owner
		}

		this.sourcesValid = true
	}

	// The block a laid-out source paragraph (part, item) stands for in the editing document: its position
	// and node, with the paragraph as built from the part the layout read (`shadow`, the same node for an
	// unedited source paragraph). Undefined for paragraphs the editing document has no node for (inside locked
	// blocks).
	paragraphOf(part: number, item: number): { pos: number; node: Node; shadow: Node; layoutPart: StoryPart } | undefined {
		const entry = this.entryOf(part, item)

		if (entry === undefined) return undefined

		const start = this.starts[entry.index] ?? 0
		const record = this.edits.recordAt(part, item)

		if (record !== undefined) {
			const shadow = this.shadowOf(record)

			if (shadow === null) return undefined

			if (item === record.item && record.node.type === nodes.paragraph) {
				return { pos: start + record.offset, node: record.node, shadow, layoutPart: record.part }
			}

			// A paragraph inside an overlay table: found in the shadow, at the same place in the table as it is
			// now (cells edited since keep the shape).
			const path = pathOfSource(shadow, item)
			const inner = path === undefined ? undefined : atPath(shadow, path)
			const table = record.offset === 0 ? entry.node : entry.node.nodeAt(record.offset - 1)
			const found = path === undefined || table === null ? undefined : offsetAtPath(table, path)

			return inner === undefined || found === undefined || !found.node.isTextblock
				? undefined
				: { pos: start + record.offset + found.offset, node: found.node, shadow: inner, layoutPart: record.part }
		}

		// A source paragraph (or one inside a source table): the editing node with that source item, laid out
		// from its overlay when it was edited.
		const found = findSource(entry.node, item, true)

		if (found === undefined) return undefined

		const pos = start + found.offset
		const substitute = this.edits.substitutes.get(ref(part, item))
		const overlay = substitute === undefined ? undefined : this.edits.recordAt(refPart(substitute), refItem(substitute))

		if (overlay !== undefined) {
			const shadow = this.shadowOf(overlay)

			return shadow === null ? undefined : { pos, node: found.node, shadow, layoutPart: overlay.part }
		}

		return { pos, node: found.node, shadow: found.node, layoutPart: this.part }
	}

	// The source key (part, item) the layout reports for the paragraph at `pos` (a textblock start inside the
	// story's document), or undefined.
	keyOf(pos: number, node: Node): { part: number; item: number } | undefined {
		const entry = this.entries[this.entryAt(pos)]

		if (entry === undefined) return undefined

		const start = this.starts[entry.index] ?? 0

		for (const record of entry.records) {
			if (record.node === node && start + record.offset === pos) return { part: record.part.index, item: record.item }
		}

		// Inside a table laid out from an overlay: the shadow's paragraph at the same place.
		for (const record of entry.records) {
			if (record.node.type !== nodes.table) continue

			const table = record.offset === 0 ? entry.node : entry.node.nodeAt(record.offset - 1)
			const offset = pos - start - record.offset - 1

			if (table === null || offset < 0 || offset >= table.content.size) continue

			const shadow = this.shadowOf(record)
			const path = pathAt(table, offset)
			const inner = shadow === null || path === undefined ? undefined : atPath(shadow, path)
			const src = inner?.attrs["src"]

			return typeof src === "number" && src >= 0 ? { part: record.part.index, item: src } : undefined
		}

		const src = node.attrs["src"]

		if (typeof src !== "number" || src < 0) return undefined

		const substitute = this.edits.substitutes.get(ref(this.part.index, src))

		return substitute === undefined ? { part: this.part.index, item: src } : { part: refPart(substitute), item: refItem(substitute) }
	}

	// The record's block as the editor builds it from the overlay part (sources = overlay items).
	shadowOf(record: OverlayRecord): Node | null {
		if (record.shadow !== undefined) return record.shadow

		const model = this.edits.model
		const story = model.storyOf(record.part, record.part.parent(record.item)) as Story | undefined
		let shadow: Node | null = null

		if (story !== undefined) {
			this.shadows ??= new DocxDocBuilder({ ...this.context, model })

			try {
				shadow = this.shadows.buildBlock(story, record.item)
			} catch {
				shadow = null
			}
		}

		record.shadow = shadow

		return shadow
	}
}

// Index of the last record whose item is at or before `item` (0 when none is).
function recordIndex(records: readonly OverlayRecord[], item: number): number {
	let low = 0
	let high = records.length - 1

	while (low < high) {
		const middle = (low + high + 1) >> 1

		if ((records[middle]?.item ?? 0) <= item) low = middle
		else high = middle - 1
	}

	return low
}

// Whether the source item list of two entries is the same.
function sameRefs(a: readonly number[], b: readonly number[]): boolean {
	if (a.length !== b.length) return false

	for (let k = 0; k < a.length; k++) if (a[k] !== b[k]) return false

	return true
}

function endsSection(node: Node): boolean {
	if (node.type === nodes.sdt) {
		let ends = false

		node.forEach(child => {
			if (endsSection(child)) ends = true
		})

		return ends
	}

	const sectPr = node.attrs["sectPr"]

	return typeof sectPr === "number" && sectPr >= 0
}

// Whether two tables have the same rows and cells (attributes included) and blocks of the same kinds and
// sources in each cell: only paragraph content differs.
function sameStructure(a: Node, b: Node): boolean {
	if (a === b) return true
	if (a.type !== b.type || a.childCount !== b.childCount) return false

	if (a.type === nodes.paragraph) return a.attrs["src"] === b.attrs["src"] && a.attrs["src"] !== -1 && a.sameMarkup(b)
	if (!a.sameMarkup(b)) return false

	for (let k = 0; k < a.childCount; k++) {
		const x = a.child(k)
		const y = b.child(k)

		if (x === y) continue

		if (x.type === nodes.paragraph && y.type === nodes.paragraph) {
			if (x.attrs["src"] !== y.attrs["src"] || (x.attrs["src"] as number) < 0) return false

			continue
		}

		if (!sameStructure(x, y)) return false
	}

	return true
}

// A paragraph changed inside a table: its position relative to the entry, its child indices from the table
// and the block (ref) it is laid out in place of (-1: its source paragraph).
interface ChangedParagraph {
	readonly node: Node
	readonly offset: number
	readonly path: readonly number[]
	base: number
}

// A block written by the edit being applied, waiting to be parsed with the others (slot: its index in the
// entry's lists, -1 for a paragraph inside a table).
interface QueuedBlock {
	readonly entry: TopEntry
	readonly node: Node
	readonly offset: number
	readonly base: number
	readonly bytes: Uint8Array
	readonly slot: number
}

// Paragraphs of `b` that differ from `a` (same structure), with their positions relative to `offset`.
function changedParagraphs(a: Node, b: Node, offset: number, path: readonly number[], out: ChangedParagraph[]): void {
	b.forEach((child, childOffset, index) => {
		const before = a.child(index)

		if (before === child) return

		const at = offset + 1 + childOffset
		const here = [...path, index]

		if (child.type === nodes.paragraph) out.push({ node: child, offset: at, path: here, base: -1 })
		else changedParagraphs(before, child, at, here, out)
	})
}

// Whether two tables have the same rows, cells and blocks (attributes included): only paragraph content
// differs. Unlike sameStructure, blocks match by place, not by source.
function sameShape(a: Node, b: Node): boolean {
	if (a === b) return true
	if (!a.sameMarkup(b) || a.childCount !== b.childCount) return false
	if (a.type === nodes.paragraph) return true

	for (let k = 0; k < a.childCount; k++) if (!sameShape(a.child(k), b.child(k))) return false

	return true
}

// The node at child indices `path` from `node`.
function atPath(node: Node, path: readonly number[]): Node | undefined {
	let at: Node = node

	for (const index of path) {
		if (index >= at.childCount) return undefined

		at = at.child(index)
	}

	return at
}

// The node at child indices `path` from `node`, with its position relative to the position of `node`.
function offsetAtPath(node: Node, path: readonly number[]): { node: Node; offset: number } | undefined {
	let at: Node = node
	let offset = 0

	for (const index of path) {
		if (index >= at.childCount) return undefined

		offset += 1

		for (let k = 0; k < index; k++) offset += at.child(k).nodeSize

		at = at.child(index)
	}

	return { node: at, offset }
}

// Child indices of the textblock with source item `item` inside `node`, or undefined.
function pathOfSource(node: Node, item: number): number[] | undefined {
	for (let k = 0; k < node.childCount; k++) {
		const child = node.child(k)

		if (child.isTextblock) {
			if (child.attrs["src"] === item) return [k]

			continue
		}

		const inner = pathOfSource(child, item)

		if (inner !== undefined) return [k, ...inner]
	}

	return undefined
}

// Child indices of the textblock starting at `offset` inside the content of `node`, or undefined.
function pathAt(node: Node, offset: number): number[] | undefined {
	let pos = 0

	for (let k = 0; k < node.childCount; k++) {
		const child = node.child(k)

		if (offset < pos + child.nodeSize) {
			if (child.isTextblock) return pos === offset ? [k] : undefined

			const inner = pathAt(child, offset - pos - 1)

			return inner === undefined ? undefined : [k, ...inner]
		}

		pos += child.nodeSize
	}

	return undefined
}

// Where `target` sits inside `node` (whose start is `offset`), or undefined.
function offsetOf(node: Node, offset: number, target: Node): number | undefined {
	let found: number | undefined

	node.descendants((child, pos) => {
		if (found !== undefined) return false

		if (child === target) {
			found = offset + 1 + pos

			return false
		}

		return !child.isTextblock
	})

	return found
}

// The textblock with source item `item` inside `node` (itself included when `self`), with its position
// relative to the position of `node`.
function findSource(node: Node, item: number, self = false): { node: Node; offset: number } | undefined {
	if (self && node.isTextblock) return node.attrs["src"] === item ? { node, offset: 0 } : undefined

	let found: { node: Node; offset: number } | undefined

	node.descendants((child, pos) => {
		if (found !== undefined) return false

		if (child.isTextblock) {
			if (child.attrs["src"] === item) found = { node: child, offset: pos + 1 }

			return false
		}

		return true
	})

	return found
}

// Top-level child index at `pos`: the child holding it, or at a boundary the one after (`before`: the one
// before).
function childAt(doc: Node, pos: number, before: boolean): number {
	const clamped = Math.max(0, Math.min(pos, doc.content.size))
	const $pos = doc.resolve(clamped)

	if ($pos.depth > 0) return $pos.index(0)

	return before ? $pos.index(0) - 1 : Math.min($pos.index(0), doc.childCount - 1)
}

// The range of the final document a transaction changed (content, marks or attributes), or undefined when
// nothing changed; `whole` when a step changed the document node itself.
function changedSpan(tr: Transaction): { from: number; to: number; whole: boolean } | undefined {
	let from = Infinity
	let to = -Infinity
	let whole = false

	tr.steps.forEach((step, index) => {
		let a: number
		let b: number

		if (step instanceof ReplaceStep) {
			a = step.from
			b = step.from + step.slice.size
		} else if (step instanceof ReplaceAroundStep) {
			a = step.from
			b = step.getMap().map(step.to, 1)
		} else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
			a = step.from
			b = step.to
		} else if (step instanceof AttrStep || step instanceof AddNodeMarkStep || step instanceof RemoveNodeMarkStep) {
			a = step.pos
			b = step.pos + 1
		} else {
			whole = true

			return
		}

		const rest = tr.mapping.slice(index + 1)

		from = Math.min(from, rest.map(a, -1))
		to = Math.max(to, rest.map(b, 1))
	})

	if (whole) return { from: 0, to: tr.doc.content.size, whole: true }

	return from === Infinity ? undefined : { from, to, whole: false }
}
