import { Fragment, Slice, type Mark, type MarkType, type Node } from "prosemirror-model"
import { Plugin, PluginKey, type EditorState, type Transaction } from "prosemirror-state"
import {
	AddMarkStep,
	AddNodeMarkStep,
	AttrStep,
	DocAttrStep,
	RemoveMarkStep,
	RemoveNodeMarkStep,
	ReplaceAroundStep,
	ReplaceStep,
	dropPoint,
	type Step
} from "prosemirror-transform"
import { isHistoryTransaction } from "prosemirror-history"
import { MAX_LEVEL } from "../numbering/index"
import {
	NO_MARKERS,
	NO_SOURCE,
	REVISION_MARKS,
	WRAPPER_MARKS,
	docxSchema,
	inlineLevel,
	markerKey,
	survivorsOf,
	type MarkerAttrs,
	type MarkerList
} from "./schema"

// Transactions the session makes itself (locking blocks, the repairs appended here) skip the invariants
// and the read-only gate. The key is private to the package's own modules, so no host can claim it.
const internalKey = new PluginKey("docxInternal")

export function markInternal(tr: Transaction): Transaction {
	return tr.setMeta(internalKey, true)
}

function isInternal(tr: Transaction): boolean {
	return tr.getMeta(internalKey) === true
}

const nodes = docxSchema.nodes
const revisionTypes: readonly MarkType[] = REVISION_MARKS.map(name => docxSchema.marks[name])
// Marks that tie content to a source element whose identity must not be duplicated.
const identityTypes: readonly MarkType[] = [...revisionTypes, docxSchema.marks.content_control, docxSchema.marks.wrapper]
const wrapperTypes: readonly MarkType[] = WRAPPER_MARKS.map(name => docxSchema.marks[name])
// Attributes that tie a node to its source or hold structure; only the session's own transactions change them.
const PROTECTED_ATTRS = new Set([
	"src",
	"end",
	"sectPr",
	"deletable",
	"kind",
	"reason",
	"markers",
	"before",
	"node",
	"story",
	"simple",
	"lock",
	"revision"
])
const SOURCE_BLOCKS = new Set(["paragraph", "table", "table_row", "table_cell", "sdt", "locked_block"])
// Atoms that exist only as source bytes: a copy of one has nothing to be written from.
const BYTES_ONLY = new Set([
	"drawing",
	"run_object",
	"math",
	"opaque_inline",
	"marker",
	"comment_ref",
	"locked_block",
	"proof_err",
	"rendered_break"
])

// ---- paste ------------------------------------------------------------------------------------------------

function freshMarks(marks: readonly Mark[]): readonly Mark[] {
	let out = marks
	// A base holding a tracked formatting change carries its revision id: such text keeps only its
	// editable properties.
	const tracked = marks.some(mark => mark.type === docxSchema.marks.format_change)

	for (const mark of marks) {
		const type = mark.type

		if (identityTypes.includes(type) || type === docxSchema.marks.format_change || (tracked && type === docxSchema.marks.run)) {
			out = mark.removeFromSet(out)
		} else if ((type === docxSchema.marks.link || type === docxSchema.marks.run) && (mark.attrs["src"] as number) !== NO_SOURCE) {
			out = type.create({ ...mark.attrs, src: NO_SOURCE }).addToSet(mark.removeFromSet(out))
		}
	}

	return out
}

function freshContent(fragment: Fragment): Fragment {
	const out: Node[] = []

	fragment.forEach(node => {
		const fresh = freshNode(node)

		if (fresh !== null) out.push(fresh)
	})

	return Fragment.fromArray(out)
}

function freshNode(node: Node): Node | null {
	const type = node.type

	if (node.isText) return node.mark(freshMarks(node.marks))

	if (BYTES_ONLY.has(type.name)) {
		// A locked paragraph comes back as its text; other byte-only content is left out.
		if (type === nodes.locked_block && node.attrs["kind"] === "paragraph" && node.attrs["text"] !== "") {
			return nodes.paragraph.create(null, docxSchema.text(node.attrs["text"] as string))
		}

		return null
	}

	const attrs: Record<string, unknown> = { ...node.attrs }

	if ("src" in attrs) attrs["src"] = NO_SOURCE
	if ("sectPr" in attrs) attrs["sectPr"] = NO_SOURCE
	if ("before" in attrs) attrs["before"] = NO_MARKERS
	if ("end" in attrs && type !== nodes.field) attrs["end"] = NO_MARKERS
	if (type === nodes.field) {
		attrs["end"] = NO_SOURCE
		attrs["markers"] = NO_MARKERS
	}
	if ("revision" in attrs) attrs["revision"] = 0

	if (node.isLeaf) return type.create(attrs, null, freshMarks(node.marks))

	const content = freshContent(node.content)

	// Containers that lost all their content are left out rather than made invalid.
	if (content.size === 0 && !type.validContent(content)) return type.createAndFill(attrs, null, freshMarks(node.marks))

	return type.create(attrs, content, freshMarks(node.marks))
}

// Content pasted or dropped into a story carries no source identity: blocks and fields become new
// elements, range markers, comment references and atoms that only exist as source bytes are left out,
// and revision, content-control and wrapper marks are removed. Formatting (run bases, properties) stays.
export function freshSlice(slice: Slice): Slice {
	return new Slice(freshContent(slice.content), slice.openStart, slice.openEnd)
}

// ---- helpers ----------------------------------------------------------------------------------------------

// Ranges of the final document that a transaction's steps touched, as [from, to] pairs (a node whose
// attributes changed counts as touched).
export function changedRanges(tr: Transaction): number[] {
	const maps = tr.mapping.maps
	const out: number[] = []
	const push = (index: number, start: number, end: number): void => {
		let from = start
		let to = end

		for (let later = index + 1; later < maps.length; later++) {
			const next = maps[later]

			if (next === undefined) continue

			from = next.map(from, -1)
			to = next.map(to, 1)
		}

		out.push(from, to)
	}

	maps.forEach((map, index) => {
		const step = tr.steps[index]

		if (step instanceof AttrStep || step instanceof AddNodeMarkStep || step instanceof RemoveNodeMarkStep)
			push(index, step.pos, step.pos + 1)
		else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) push(index, step.from, step.to)
		else map.forEach((_oldStart, _oldEnd, newStart, newEnd) => push(index, newStart, newEnd))
	})

	return out
}

// What each replace step inserted, as [from, to] pairs in the final document (both sides of the gap of a
// ReplaceAroundStep).
function insertedRanges(tr: Transaction): number[] {
	const maps = tr.mapping.maps
	const out: number[] = []

	tr.steps.forEach((step, index) => {
		if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep) || step.slice.size === 0) return

		maps[index]?.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
			if (newEnd <= newStart) return

			let from = newStart
			let to = newEnd

			for (let later = index + 1; later < maps.length; later++) {
				from = maps[later]?.map(from, 1) ?? from
				to = maps[later]?.map(to, -1) ?? to
			}

			if (to > from) out.push(from, to)
		})
	})

	return out
}

interface DeletedRange {
	readonly step: number
	readonly doc: Node
	readonly from: number
	readonly to: number
}

function deletedRanges(tr: Transaction): DeletedRange[] {
	const out: DeletedRange[] = []

	tr.steps.forEach((step, index) => {
		const doc = tr.docs[index]

		if (doc === undefined) return

		if (step instanceof ReplaceStep) {
			if (step.to > step.from) out.push({ step: index, doc, from: step.from, to: step.to })
		} else if (step instanceof ReplaceAroundStep) {
			if (step.gapFrom > step.from) out.push({ step: index, doc, from: step.from, to: step.gapFrom })
			if (step.to > step.gapTo) out.push({ step: index, doc, from: step.gapTo, to: step.to })
		}
	})

	return out
}

function skips(tr: Transaction): boolean {
	return !tr.docChanged || isInternal(tr) || isHistoryTransaction(tr)
}

// A mark no edit may add: revisions, content controls, wrappers, and links or runs that name a source
// element (copies of those would duplicate their identity).
function identityMark(mark: Mark): boolean {
	const type = mark.type

	if (identityTypes.includes(type)) return true

	return (type === docxSchema.marks.link || type === docxSchema.marks.run) && (mark.attrs["src"] as number) !== NO_SOURCE
}

function protectedChange(step: Step, doc: Node): boolean {
	if (step instanceof AttrStep) return PROTECTED_ATTRS.has(step.attr)
	if (step instanceof DocAttrStep) return true

	// setNodeMarkup: the same node around the same content with other attributes.
	if (step instanceof ReplaceAroundStep && step.insert === 1 && step.gapFrom === step.from + 1 && step.gapTo === step.to - 1) {
		const before = doc.nodeAt(step.from)
		const after = step.slice.content.firstChild

		if (before === null || after === null || step.slice.content.childCount !== 1) return false

		for (const name of PROTECTED_ATTRS) {
			if (name in before.attrs && before.attrs[name] !== after.attrs[name]) return true
		}
	}

	return false
}

// ---- filter -----------------------------------------------------------------------------------------------

// Why a transaction would break the document's structure, or undefined when it is acceptable.
export function invariantViolation(tr: Transaction): string | undefined {
	if (skips(tr)) return undefined

	for (let index = 0; index < tr.steps.length; index++) {
		const step = tr.steps[index]
		const doc = tr.docs[index]

		if (step === undefined || doc === undefined) continue
		if (protectedChange(step, doc)) return "protectedAttribute"
		if ((step instanceof AddMarkStep || step instanceof AddNodeMarkStep) && identityMark(step.mark)) return "identityMark"
	}

	for (const range of deletedRanges(tr)) {
		let violation: string | undefined

		range.doc.nodesBetween(range.from, range.to, (node, pos) => {
			if (violation !== undefined) return false

			const type = node.type
			const close = pos + node.nodeSize - 1
			const opens = pos >= range.from && pos < range.to
			const closes = close >= range.from && close < range.to

			if (type === nodes.paragraph) {
				if ((node.attrs["sectPr"] as number) !== NO_SOURCE && (opens || closes)) violation = "sectionBreak"

				return false
			}

			if (type === nodes.locked_block) {
				if (opens && node.attrs["deletable"] !== true) violation = "lockedBlock"

				return false
			}

			if (type === nodes.sdt && opens && node.attrs["lock"] === "sdtLocked") violation = "lockedContentControl"

			return !node.isInline
		})

		if (violation !== undefined) return violation
	}

	// Nothing may be inserted strictly inside deleted (or moved-away) text: it is read-only.
	for (let index = 0; index < tr.steps.length; index++) {
		const step = tr.steps[index]
		const doc = tr.docs[index]

		if (doc === undefined || !(step instanceof ReplaceStep) || step.slice.size === 0) continue

		const before = doc.resolve(step.from).nodeBefore
		const after = doc.resolve(step.to).nodeAfter

		if (before === null || after === null || !before.isInline || !after.isInline) continue

		for (const mark of before.marks) {
			if ((mark.type === docxSchema.marks.deletion || mark.type === docxSchema.marks.move_from) && mark.isInSet(after.marks)) {
				return "deletedText"
			}
		}
	}

	return undefined
}

// ---- repairs ----------------------------------------------------------------------------------------------

export interface ListAttrs {
	readonly numId: number | null
	readonly ilvl: number | null
}

// A hook that numbering features plug into: given a changed paragraph's list attributes, the attributes it
// should have instead (or undefined to keep them).
export type ListNormalizer = (attrs: ListAttrs, paragraph: Node) => ListAttrs | undefined

export function normalizeListAttrs(attrs: ListAttrs): ListAttrs | undefined {
	const ilvl = attrs.ilvl
	let next: ListAttrs | undefined

	if (attrs.numId === null && ilvl !== null) next = { numId: null, ilvl: null }
	else if (ilvl !== null && (ilvl < 0 || ilvl > MAX_LEVEL || !Number.isInteger(ilvl)))
		next = { numId: attrs.numId, ilvl: Math.min(MAX_LEVEL, Math.max(0, Math.trunc(ilvl) || 0)) }

	return next
}

function keysIn(node: Node, keys: Set<string>): void {
	for (const marker of survivorsOf(node)) keys.add(markerKey(marker))

	if (node.isBlock) {
		for (const name of ["before", "end"]) {
			const list: unknown = node.attrs[name]

			if (Array.isArray(list)) for (const marker of list as MarkerList) keys.add(markerKey(marker))
		}
	}
}

function markersOf(node: Node, names: readonly string[], out: Map<string, MarkerAttrs>): void {
	for (const name of names) {
		const list: unknown = node.attrs[name]

		if (Array.isArray(list)) for (const marker of list as MarkerList) out.set(markerKey(marker), marker)
	}
}

function markerNode(marker: MarkerAttrs): Node {
	return marker.kind === "commentReference"
		? nodes.comment_ref.create({ src: marker.src, id: marker.id })
		: nodes.marker.create({ src: marker.src, kind: marker.kind, id: marker.id })
}

// Puts lost markers back at `pos`: inline where the position is in a paragraph, otherwise in the
// `before` list of the next block or the `end` list of the container.
function reinsert(tr: Transaction, pos: number, markers: readonly MarkerAttrs[]): void {
	const $pos = tr.doc.resolve(Math.max(0, Math.min(pos, tr.doc.content.size)))

	if ($pos.parent.inlineContent) {
		tr.insert($pos.pos, markers.map(markerNode))

		return
	}

	const after = $pos.nodeAfter

	if (after !== null && "before" in after.attrs) {
		tr.setNodeAttribute($pos.pos, "before", [...markers, ...(after.attrs["before"] as MarkerList)])

		return
	}

	const parent = $pos.parent

	if ($pos.depth === 0) {
		tr.setDocAttribute("end", [...(parent.attrs["end"] as MarkerList), ...markers])
	} else if ("end" in parent.attrs) {
		tr.setNodeAttribute($pos.before(), "end", [...(parent.attrs["end"] as MarkerList), ...markers])
	}
}

export interface InvariantOptions {
	readonly readOnly?: () => boolean
	readonly normalizeList?: ListNormalizer
}

// Keys of the source identities a node carries: its own (a source block or atom) and its markers'.
function identityKeys(node: Node, own: boolean, out: string[]): void {
	if (own) {
		const src: unknown = node.attrs["src"]

		if (node.type === nodes.proof_err) out.push(`p${String(node.attrs["node"])}`)
		else if (typeof src === "number" && src !== NO_SOURCE && node.type !== nodes.marker && node.type !== nodes.comment_ref)
			out.push(`${node.isBlock ? "b" : "i"}${src}`)
	}

	for (const marker of survivorsOf(node)) out.push(`m${markerKey(marker)}`)

	if (node.isBlock) {
		for (const name of ["before", "end"]) {
			const list: unknown = node.attrs[name]

			if (Array.isArray(list)) for (const marker of list as MarkerList) out.push(`m${markerKey(marker)}`)
		}
	}
}

interface DuplicateFix {
	readonly pos: number
	// Replace the node with a fresh copy (or remove it), or only drop duplicate markers from its lists.
	readonly fresh: boolean
	readonly drop: ReadonlySet<string>
}

// Content inserted by a transaction that duplicates a source identity present elsewhere (a copy made
// without the paste sanitiser: a drag-copy, a host's own transaction): the copy becomes new content.
function duplicateFixes(doc: Node, inserted: readonly number[]): DuplicateFix[] {
	const inside: { node: Node; pos: number; keys: string[] }[] = []
	const size = doc.content.size
	const spans: [number, number][] = []

	for (let k = 0; k < inserted.length; k += 2) {
		const from = Math.max(0, Math.min(size, inserted[k] ?? 0))
		const to = Math.max(from, Math.min(size, inserted[k + 1] ?? 0))

		if (to > from) spans.push([from, to])
	}

	spans.sort((a, b) => a[0] - b[0])

	const seenPos = new Set<number>()

	for (const [from, to] of spans) {
		doc.nodesBetween(from, to, (node, pos) => {
			if (pos < from || pos + node.nodeSize > to) return !node.isLeaf

			if (seenPos.has(pos)) return false

			seenPos.add(pos)

			const keys: string[] = []

			identityKeys(node, node.isInline || SOURCE_BLOCKS.has(node.type.name), keys)

			if (keys.length > 0) inside.push({ node, pos, keys })

			// A source block is judged whole; a fresh one may hold copies.
			const src: unknown = node.attrs["src"]

			return !node.isLeaf && !(typeof src === "number" && src !== NO_SOURCE)
		})
	}

	if (inside.length === 0) return []

	// Identities outside the inserted content.
	const outside = new Set<string>()
	let span = 0

	doc.descendants((node, pos) => {
		while (span < spans.length && (spans[span]?.[1] ?? 0) <= pos) span++

		const current = spans[span]

		if (current !== undefined && pos >= current[0] && pos + node.nodeSize <= current[1]) return false

		const keys: string[] = []

		identityKeys(node, true, keys)

		for (const key of keys) outside.add(key)

		return !node.isLeaf
	})

	for (const key of doc.attrs["end"] as MarkerList) outside.add(`m${markerKey(key)}`)

	const fixes: DuplicateFix[] = []
	const taken = new Set<string>()

	for (const { node, pos, keys } of inside) {
		const ownKey = keys[0]
		const ownIsSource = ownKey !== undefined && !ownKey.startsWith("m")
		const markerOnly = node.type === nodes.marker || node.type === nodes.comment_ref
		const duplicate = keys.filter(key => outside.has(key) || taken.has(key))

		for (const key of keys) taken.add(key)

		if (duplicate.length === 0) continue

		if ((ownIsSource && duplicate.includes(ownKey)) || markerOnly) fixes.push({ pos, fresh: true, drop: new Set() })
		else fixes.push({ pos, fresh: false, drop: new Set(duplicate.map(key => key.slice(1))) })
	}

	return fixes
}

function filterMarkers(list: unknown, drop: ReadonlySet<string>): MarkerList | undefined {
	if (!Array.isArray(list)) return undefined

	const kept = (list as MarkerList).filter(marker => !drop.has(markerKey(marker)))

	return kept.length === list.length ? undefined : kept.length === 0 ? NO_MARKERS : kept
}

function applyDuplicateFix(repair: Transaction, fix: DuplicateFix): void {
	const pos = repair.mapping.map(fix.pos, 1)
	const node = repair.doc.nodeAt(pos)

	if (node === null) return

	if (!fix.fresh) {
		for (const name of ["before", "end", "markers"]) {
			const kept = filterMarkers(node.attrs[name], fix.drop)

			if (kept !== undefined) repair.setNodeAttribute(pos, name, kept)
		}

		return
	}

	const fresh = freshNode(node)
	const $pos = repair.doc.resolve(pos)

	if (fresh !== null) {
		repair.replaceWith(pos, pos + node.nodeSize, fresh)
	} else if (node.isInline || $pos.parent.childCount > 1) {
		repair.delete(pos, pos + node.nodeSize)
	} else {
		repair.replaceWith(pos, pos + node.nodeSize, nodes.paragraph.create())
	}
}

export interface InvariantOptions {
	readonly readOnly?: () => boolean
	readonly normalizeList?: ListNormalizer
}

// What appendTransaction adds after `transactions`: revision and stray marks off new or re-marked
// content, one holder per section break and marker, list attributes normalised, inserted copies of
// source identities made new, and lost markers put back. Null when nothing needs repair.
export function repairTransaction(
	transactions: readonly Transaction[],
	state: EditorState,
	options: InvariantOptions = {}
): Transaction | null {
	const relevant = transactions.filter(tr => !skips(tr))

	if (relevant.length === 0) return null

	const repair = state.tr
	const size = state.doc.content.size
	// Positions are in each transaction's final document; map them onto `state.doc`.
	const toState = (index: number, pos: number, assoc: number): number => {
		let mapped = pos

		for (let later = transactions.indexOf(relevant[index] as Transaction) + 1; later < transactions.length; later++) {
			mapped = transactions[later]?.mapping.map(mapped, assoc) ?? mapped
		}

		return Math.max(0, Math.min(size, mapped))
	}
	const lost = new Map<number, MarkerAttrs[]>()
	const changed: number[] = []
	const inserted: number[] = []
	let insertsRevisions = false

	relevant.forEach((tr, index) => {
		const ranges = changedRanges(tr)
		const insertions = insertedRanges(tr)

		for (let k = 0; k < ranges.length; k += 2) changed.push(toState(index, ranges[k] ?? 0, -1), toState(index, ranges[k + 1] ?? 0, 1))
		for (let k = 0; k < insertions.length; k += 2)
			inserted.push(toState(index, insertions[k] ?? 0, 1), toState(index, insertions[k + 1] ?? 0, -1))

		insertsRevisions ||= tr.steps.some(
			step =>
				(step instanceof ReplaceStep || step instanceof ReplaceAroundStep) &&
				step.slice.size > 0 &&
				carries(step.slice.content, mark => revisionTypes.includes(mark.type))
		)

		// Markers a deletion took out, by key, with where the deletion ended up.
		const candidates = new Map<string, { marker: MarkerAttrs; pos: number }>()

		for (const range of deletedRanges(tr)) {
			const found = new Map<string, MarkerAttrs>()

			range.doc.nodesBetween(range.from, range.to, (node, pos) => {
				const end = pos + node.nodeSize
				const opens = pos >= range.from && pos < range.to

				if (node.isLeaf) {
					if (pos >= range.from && end <= range.to) for (const marker of survivorsOf(node)) found.set(markerKey(marker), marker)
				} else if (node.isBlock && opens) {
					// A block whose start token goes loses its attributes (a join keeps the first node's).
					markersOf(node, ["before", "end"], found)
				} else if (node.isBlock && end - 1 >= range.from && end - 1 < range.to) {
					markersOf(node, ["end"], found)
				}

				return true
			})

			if (found.size === 0) continue

			let pos = range.from

			for (let later = range.step; later < tr.steps.length; later++) pos = tr.mapping.maps[later]?.map(pos, -1) ?? pos

			for (const [key, marker] of found) candidates.set(key, { marker, pos })
		}

		if (candidates.size === 0) return

		const present = new Set<string>()
		const finalSize = tr.doc.content.size

		for (let k = 0; k < ranges.length; k += 2) {
			const from = Math.max(0, Math.min(finalSize, (ranges[k] ?? 0) - 1))
			const to = Math.max(from, Math.min(finalSize, (ranges[k + 1] ?? 0) + 1))

			tr.doc.nodesBetween(from, to, node => {
				keysIn(node, present)

				return true
			})
		}

		for (const marker of tr.doc.attrs["end"] as MarkerList) present.add(markerKey(marker))

		for (const [key, { marker, pos }] of candidates) {
			if (present.has(key)) continue

			const at = toState(index, pos, -1)
			const list = lost.get(at)

			if (list === undefined) lost.set(at, [marker])
			else list.push(marker)
		}
	})

	// 1. Marks. New content never carries a source revision: text typed inside someone's tracked insertion
	// is plain text (the insertion splits). Atoms at paragraph level carry wrapper marks only (a marker
	// written inside a w:r also its run).
	if (insertsRevisions) {
		for (let k = 0; k < inserted.length; k += 2) {
			const from = inserted[k] ?? 0
			const to = inserted[k + 1] ?? 0

			if (to > from) for (const type of revisionTypes) repair.removeMark(from, to, type)
		}
	}

	const spans: [number, number][] = []

	for (let k = 0; k < changed.length; k += 2) {
		const from = Math.max(0, Math.min(size, changed[k] ?? 0))

		spans.push([from, Math.max(from, Math.min(size, changed[k + 1] ?? 0))])
	}

	spans.sort((a, b) => a[0] - b[0])

	const holders = new Map<number, number[]>()
	const listFixes: { pos: number; attrs: ListAttrs }[] = []
	const markerFixes: { pos: number; name: string; list: MarkerList }[] = []
	const strayMarks: { pos: number; mark: Mark }[] = []
	const listed = new Set<string>()
	const seen = new Set<number>()

	for (const [from, to] of spans) {
		state.doc.nodesBetween(from, to, (node, pos) => {
			if (seen.has(pos)) return !node.isTextblock && !node.isLeaf

			seen.add(pos)

			if (node.isInline) {
				if (inlineLevel(node) === "paragraph") {
					for (const mark of node.marks) {
						if (wrapperTypes.includes(mark.type)) continue
						if (mark.type === docxSchema.marks.run && node.type === nodes.marker) continue

						strayMarks.push({ pos, mark })
					}
				}

				return false
			}

			// A split copies a block's marker lists into both halves: the first keeps them.
			for (const name of ["before", "end"]) {
				const list: unknown = node.attrs[name]

				if (!Array.isArray(list) || list.length === 0) continue

				const kept = (list as MarkerList).filter(marker => {
					const key = markerKey(marker)

					if (listed.has(key)) return false

					listed.add(key)

					return true
				})

				if (kept.length !== list.length) markerFixes.push({ pos, name, list: kept.length === 0 ? NO_MARKERS : kept })
			}

			const sectPr = node.attrs["sectPr"]

			if ((node.type === nodes.paragraph || node.type === nodes.locked_block) && typeof sectPr === "number" && sectPr !== NO_SOURCE) {
				const list = holders.get(sectPr)

				if (list === undefined) holders.set(sectPr, [pos])
				else list.push(pos)
			}

			if (node.type === nodes.paragraph) {
				const attrs: ListAttrs = { numId: node.attrs["numId"] as number | null, ilvl: node.attrs["ilvl"] as number | null }
				const fixed = normalizeListAttrs(attrs)
				const hooked = options.normalizeList?.(fixed ?? attrs, node) ?? fixed

				if (hooked !== undefined && (hooked.numId !== attrs.numId || hooked.ilvl !== attrs.ilvl))
					listFixes.push({ pos, attrs: hooked })
			}

			return node.type !== nodes.locked_block
		})
	}

	for (const stray of strayMarks) repair.removeNodeMark(stray.pos, stray.mark)

	// 2. Attributes (no position moves).
	for (const fix of markerFixes) repair.setNodeAttribute(fix.pos, fix.name, fix.list)

	// A split paragraph's halves both carry the break; it belongs to the last (the paragraph mark).
	for (const positions of holders.values()) {
		if (positions.length < 2) continue

		positions.sort((a, b) => a - b)

		for (const pos of positions.slice(0, -1)) repair.setNodeAttribute(pos, "sectPr", NO_SOURCE)
	}

	for (const fix of listFixes) {
		repair.setNodeAttribute(fix.pos, "numId", fix.attrs.numId)
		repair.setNodeAttribute(fix.pos, "ilvl", fix.attrs.ilvl)
	}

	// 3. Inserted copies of source identities become new content, from the end.
	for (const fix of duplicateFixes(state.doc, inserted).reverse()) applyDuplicateFix(repair, fix)

	// 4. Lost markers last, from the end so earlier positions stay valid.
	for (const at of [...lost.keys()].sort((a, b) => b - a)) reinsert(repair, repair.mapping.map(at, -1), lost.get(at) ?? [])

	if (!repair.docChanged) return null

	return markInternal(repair)
}

function carries(fragment: Fragment, test: (mark: Mark) => boolean): boolean {
	let found = false

	fragment.descendants(node => {
		if (!found && node.marks.some(test)) found = true

		return !found
	})

	return found
}

export const docxInvariantsKey = new PluginKey("docxInvariants")

// The structural invariants of an editing document (see invariantViolation and repairTransaction), the
// read-only gate, and paste sanitising.
export function docxInvariants(options: InvariantOptions = {}): Plugin {
	return new Plugin({
		key: docxInvariantsKey,
		filterTransaction(tr) {
			if (tr.docChanged && !isInternal(tr) && options.readOnly?.() === true) return false

			return invariantViolation(tr) === undefined
		},
		appendTransaction(transactions, _old, state) {
			return repairTransaction(transactions, state, options)
		},
		props: {
			transformPasted(slice, view) {
				// A drag keeps its slice until the drop knows whether it moves (see handleDrop).
				return view.dragging === null ? freshSlice(slice) : slice
			},
			handleDrop(view, event, slice, moved) {
				// A move keeps identity (the original is deleted); a drop from outside was parsed from the
				// clipboard and has none. A copy inside the document goes in as new content.
				if (moved || view.dragging === null) return false

				const coords = view.posAtCoords({ left: Number(Reflect.get(event, "clientX")), top: Number(Reflect.get(event, "clientY")) })

				if (coords === null) return false

				const fresh = freshSlice(slice)
				const doc = view.state.doc
				const at = dropPoint(doc, coords.pos, fresh) ?? coords.pos

				view.dispatch(view.state.tr.replaceRange(at, at, fresh).setMeta("uiEvent", "drop"))

				return true
			}
		}
	})
}
