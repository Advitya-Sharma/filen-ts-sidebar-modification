import { NodeKind, XmlError, decodeUtf8, encodeUtf8, escapeAttribute, parseXml, type XmlDocument } from "../../xml/index"
import type { StoryDoc } from "../edit/build"
import type { DocxIds } from "../edit/context"
import { partReader } from "../model/reader"
import type { StoryPart } from "../model/index"
import { Ns } from "../model/namespaces"
import type { Out } from "./markup"
import type { StoryDraft } from "./story"
import { DocxWriteError } from "./values"

// A story container rewritten inside source bytes another story copies (a text box in a body paragraph):
// source bytes [from, to) of the part are written as `bytes` instead.
export interface NestedSplice {
	readonly from: number
	readonly to: number
	readonly bytes: Uint8Array
}

// The rewritten containers inside source bytes [from, to) of `part`, ascending and apart (WriteEnvironment.nested).
export type NestedStories = (part: StoryPart, from: number, to: number) => readonly NestedSplice[] | undefined

// Source bytes [from, to) of a story's part as written, with the rewritten containers inside them spliced
// in. `trusted`: the range is whole source nodes (see Out.copy); with a splice it is new markup to gate 2.
export function copySource(out: Out, part: StoryPart, from: number, to: number, trusted: boolean, nested: NestedStories | undefined): void {
	const bytes = part.document.bytes
	const splices = nested?.(part, from, to)

	if (splices === undefined || splices.length === 0) {
		if (trusted) out.copy(bytes, from, to)
		else out.bytes(bytes, from, to)

		return
	}

	let at = from

	for (const splice of splices) {
		if (splice.from < at || splice.to < splice.from || splice.to > to)
			throw new DocxWriteError("unsupported", `Nested stories of ${part.name} overlap`)

		out.bytes(bytes, at, splice.from)
		out.bytes(splice.bytes, 0, splice.bytes.length)
		at = splice.to
	}

	out.bytes(bytes, at, to)
}

// ---- text boxes and their copies ---------------------------------------------------------------------------

// Word writes a DrawingML text box inside mc:AlternateContent twice: in its mc:Choice (wps:txbx) and again in
// the VML of its mc:Fallback (v:textbox; [MS-ODRAWXML] 2.2.7.3: a wps shape's fallback is a w:pict), the same
// w:txbxContent (Word 16 output, e.g. the govuk forms of the corpus). A text box is edited in the branch the
// model reads and written into every branch, so a reader of either shows the same text.

export type TextboxProblem = "linkedTextbox" | "textboxFallback"

// Where a fallback copy's renumbered ids come from (a save: its own revision count, the editor's other ids).
export type NestedIds = Pick<DocxIds, "next">

const WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

function isElement(document: XmlDocument, node: number, local: string): boolean {
	return document.kinds[node] === NodeKind.Element && document.localName(node) === local
}

// The w:txbxContent elements of `branch` not inside another one, in document order.
function textboxesOf(part: StoryPart, branch: number): number[] {
	const document = part.document
	const reader = part.reader
	const end = document.end(branch)
	const out: number[] = []

	for (let node = branch + 1; node < document.count && document.start(node) < end; node++) {
		if (reader.group(node) !== Ns.W || !isElement(document, node, "txbxContent")) continue

		out.push(node)

		// Its subtree is the text box's own.
		const close = document.end(node)

		while (node + 1 < document.count && document.start(node + 1) < close) node++
	}

	return out
}

// Ids of wps:txbx text boxes whose text flows on into others (wps:linkedTxbx names them), per part.
const linkedIds = new WeakMap<XmlDocument, ReadonlySet<string>>()

function linkedTextboxes(part: StoryPart): ReadonlySet<string> {
	const document = part.document
	let found = linkedIds.get(document)

	if (found === undefined) {
		const ids = new Set<string>()

		for (let node = 0; node < document.count; node++) {
			if (isElement(document, node, "linkedTxbx") && part.reader.group(node) === Ns.Wps) {
				const id = part.reader.attr(node, Ns.None, "id")

				if (id !== undefined) ids.add(id)
			}
		}

		found = ids
		linkedIds.set(document, found)
	}

	return found
}

// Whether a w:txbxContent's text flows into other boxes: a wps:txbx another shape links to, or a VML text box
// with a next box (mso-next-textbox).
function linked(part: StoryPart, content: number): boolean {
	const document = part.document
	const holder = document.parent(content)

	if (holder < 0) return false

	if (isElement(document, holder, "txbx") && part.reader.namespaces.elementUri(holder) === WPS_URI) {
		const id = part.reader.attr(holder, Ns.None, "id")

		return id !== undefined && linkedTextboxes(part).has(id)
	}

	if (isElement(document, holder, "textbox") && part.reader.group(holder) === Ns.V) {
		return /mso-next-textbox\s*:/i.test(part.reader.attr(holder, Ns.None, "style") ?? "")
	}

	return false
}

// The other copies (w:txbxContent nodes) of text box content `content` a save must write it into, or why it
// cannot be edited: its text flows on into other boxes, or the branches around it do not match up (one
// mc:AlternateContent around it at most, the same count of text boxes in every branch that has any).
export function textboxCopies(part: StoryPart, content: number): readonly number[] | TextboxProblem {
	const document = part.document
	const reader = part.reader
	let branch = -1
	let alternates = 0

	for (let at = document.parent(content); at >= 0; at = document.parent(at)) {
		if (reader.group(at) !== Ns.MC) continue

		const local = document.localName(at)

		if (local === "AlternateContent") alternates++
		else if ((local === "Choice" || local === "Fallback") && branch < 0) branch = at
	}

	if (linked(part, content)) return "linkedTextbox"
	if (alternates === 0) return []
	if (alternates > 1 || branch < 0) return "textboxFallback"

	const own = textboxesOf(part, branch)
	const index = own.indexOf(content)

	if (index < 0) return "textboxFallback"

	const copies: number[] = []
	const alternate = document.parent(branch)

	for (let other = document.firstChildElement(alternate); other >= 0; other = document.nextSiblingElement(other)) {
		if (other === branch || reader.group(other) !== Ns.MC) continue

		const theirs = textboxesOf(part, other)

		// A branch with no text box at all (a picture of the shape) holds no copy of the text.
		if (theirs.length === 0) continue

		const copy = theirs[index]

		if (theirs.length !== own.length || copy === undefined || linked(part, copy)) return "textboxFallback"

		copies.push(copy)
	}

	return copies
}

// ---- the copy in the fallback -------------------------------------------------------------------------------

const W14_URI = "http://schemas.microsoft.com/office/word/2010/wordml"
// Elements whose w:id is unique per document (gate 2 counts them): the copy gets ids of its own, as Word's
// fallback copies of a text box's bookmarks do (govuk IHM forms of the corpus: the same bookmark, another id).
const RENUMBERED: Readonly<Record<string, "bookmark" | "revision">> = {
	bookmarkStart: "bookmark",
	bookmarkEnd: "bookmark",
	ins: "revision",
	del: "revision",
	moveFrom: "revision",
	moveTo: "revision",
	rPrChange: "revision",
	pPrChange: "revision",
	sectPrChange: "revision",
	tblPrChange: "revision",
	trPrChange: "revision",
	tcPrChange: "revision",
	tblGridChange: "revision",
	tblPrExChange: "revision",
	numberingChange: "revision",
	cellIns: "revision",
	cellDel: "revision",
	cellMerge: "revision",
	moveFromRangeStart: "revision",
	moveFromRangeEnd: "revision",
	moveToRangeStart: "revision",
	moveToRangeEnd: "revision"
}
// References to a comment or note: one per anchor, so the copy has none (the mc:Choice keeps them).
const DROPPED = new Set(["commentRangeStart", "commentRangeEnd", "commentReference", "footnoteReference", "endnoteReference"])

// The ids the copy `node` of a text box gives its bookmarks (by name) and permission ranges (in order), for the
// rewritten copy to keep.
function copyIds(part: StoryPart, node: number): { readonly bookmarks: Map<string, string>; readonly permissions: string[] } {
	const document = part.document
	const reader = part.reader
	const bookmarks = new Map<string, string>()
	const permissions: string[] = []
	const end = document.end(node)

	for (let at = node + 1; at < document.count && document.start(at) < end; at++) {
		if (document.kinds[at] !== NodeKind.Element || reader.group(at) !== Ns.W) continue

		const local = document.localName(at)
		const id = reader.attr(at, Ns.W, "id")

		if (id === undefined) continue
		if (local === "bookmarkStart") {
			const name = reader.attr(at, Ns.W, "name")

			if (name !== undefined && !bookmarks.has(name)) bookmarks.set(name, id)
		} else if (local === "permStart") {
			permissions.push(id)
		}
	}

	return { bookmarks, permissions }
}

// What a text box's content is written as in the other branches of its mc:AlternateContent: the same markup
// with the ids that must be unique renumbered (bookmarks, permission ranges, revisions, drawings; a bookmark or
// permission range the copy being replaced (`copy`) had keeps its id there), references to comments and notes
// left out, and no paragraph ids (w14:paraId and w14:textId repeated would be duplicates).
export function fallbackCopy(part: StoryPart, bytes: Uint8Array, ids: NestedIds | undefined, copy?: number): Uint8Array {
	const kept = copy === undefined ? { bookmarks: new Map<string, string>(), permissions: [] } : copyIds(part, copy)
	let permissions = 0
	const source = part.document
	const root = source.documentElement()
	const open = decodeUtf8(source.bytes, source.start(root), source.openEnd(root))
	const head = encodeUtf8(source.isSelfClosing(root) ? `${open.replace(/\s*\/>$/, "")}>` : open)
	const close = encodeUtf8(`</${source.name(root)}>`)
	const wrapped = new Uint8Array(head.length + bytes.length + close.length)

	wrapped.set(head, 0)
	wrapped.set(bytes, head.length)
	wrapped.set(close, head.length + bytes.length)

	let document: XmlDocument

	try {
		document = parseXml(wrapped)
	} catch (error) {
		if (error instanceof XmlError) throw new DocxWriteError("unsupported", `A rewritten text box of ${part.name} does not parse`)

		throw error
	}

	const reader = partReader(document)
	const w14 = reader.namespaces.prefixFor(document.documentElement(), W14_URI)
	const paragraphIds = w14 === undefined || w14 === "" ? undefined : new Set([`${w14}:paraId`, `${w14}:textId`])
	const edits: { from: number; to: number; text: string }[] = []
	const renamed = new Map<string, number>()
	const fresh = (space: "bookmark" | "permission" | "revision" | "docPr", old: string, keep?: string): string => {
		const key = `${space}:${old}`
		let id = renamed.get(key)

		if (id === undefined && keep !== undefined && /^-?\d+$/.test(keep)) id = Number(keep)

		if (id === undefined) {
			if (ids === undefined) throw new DocxWriteError("unsupported", `A text box of ${part.name} has ids its copy cannot renumber`)

			id = ids.next(space)
		}

		renamed.set(key, id)

		return String(id)
	}
	const attribute = (node: number, name: string, change: (value: string) => string | null): void => {
		const spans = document.attributeSpans(node)

		for (let k = 0; k < spans.length; k += 4) {
			if (decodeUtf8(document.bytes, spans[k] ?? 0, spans[k + 1] ?? 0) !== name) continue

			const value = change(decodeUtf8(document.bytes, spans[k + 2] ?? 0, spans[k + 3] ?? 0))

			if (value === null) {
				let from = spans[k] ?? 0

				while (from > document.start(node) && /\s/.test(String.fromCharCode(document.bytes[from - 1] ?? 0))) from--

				edits.push({ from, to: (spans[k + 3] ?? 0) + 1, text: "" })
			} else {
				edits.push({ from: spans[k + 2] ?? 0, to: spans[k + 3] ?? 0, text: escapeAttribute(value) })
			}
		}
	}
	const box = document.firstChildElement(document.documentElement())
	const end = box < 0 ? 0 : document.end(box)

	for (let node = box; node >= 0 && node < document.count && document.start(node) < end; node++) {
		if (document.kinds[node] !== NodeKind.Element) continue

		const group = reader.group(node)
		const local = document.localName(node)

		if (group === Ns.W && DROPPED.has(local)) {
			edits.push({ from: document.start(node), to: document.end(node), text: "" })

			const after = document.end(node)

			while (node + 1 < document.count && document.start(node + 1) < after) node++

			continue
		}

		if (paragraphIds !== undefined && group === Ns.W && (local === "p" || local === "tr")) {
			for (const name of paragraphIds) attribute(node, name, () => null)
		}

		const space =
			group === Ns.W
				? local === "permStart" || local === "permEnd"
					? "permission"
					: RENUMBERED[local]
				: group === Ns.WP && local === "docPr"
					? "docPr"
					: undefined

		if (space !== undefined) {
			const name = space === "docPr" ? "id" : document.name(node).includes(":") ? `${document.prefix(node)}:id` : "id"
			const keep =
				local === "bookmarkStart"
					? kept.bookmarks.get(reader.attr(node, Ns.W, "name") ?? "")
					: local === "permStart"
						? kept.permissions[permissions++]
						: undefined

			attribute(node, name, value => fresh(space, value, keep))
		}
	}

	edits.sort((a, b) => a.from - b.from)

	const out: Uint8Array[] = []
	let at = document.start(box)

	for (const edit of edits) {
		out.push(wrapped.subarray(at, edit.from), encodeUtf8(edit.text))
		at = edit.to
	}

	out.push(wrapped.subarray(at, end))

	const length = out.reduce((sum, piece) => sum + piece.length, 0)
	const joined = new Uint8Array(length)
	let offset = 0

	for (const piece of out) {
		joined.set(piece, offset)
		offset += piece.length
	}

	return joined
}

// ---- the plan -----------------------------------------------------------------------------------------------

// How a save writes stories that sit inside other stories: the sessions in the order they are written
// (inner containers first), which are written into the copies of their outer story rather than into their
// part, and env.nested fed with each inner story's draft.
export interface NestedPlan {
	readonly order: readonly number[]
	inner(index: number): boolean
	record(index: number, draft: StoryDraft): void
	readonly nested: NestedStories
}

// A rewritten text box as a save spliced it: its part, and what each of its container (the branch the model
// reads) and its copies (source ranges) became.
export interface NestedTextbox {
	readonly part: StoryPart
	readonly splices: readonly NestedSplice[]
	// The container's own.
	readonly bytes: Uint8Array
	// How many of the splices the save wrote (none when nothing it writes holds the text box any more).
	used: number
}

// The text boxes the last save writing a part rewrote in it (for gate 2's text box checks).
const lastWritten = new WeakMap<StoryPart, NestedTextbox[]>()

export function writtenTextboxes(part: StoryPart): readonly NestedTextbox[] {
	return lastWritten.get(part) ?? []
}

// Text boxes written with the stories holding them: each text box session is written first, and its draft
// takes the place of its container, and of the container's copies in the other branches of the
// mc:AlternateContent around it, wherever the story holding it copies those bytes. A changed text box whose
// story this save does not write cannot be written (its copies would stay as they were): the save is refused.
export function planNested(stories: readonly StoryDoc[], ids?: NestedIds): NestedPlan | undefined {
	const byStory = new Map<number, number>()

	for (const story of stories) lastWritten.delete(story.part)

	stories.forEach((story, index) => byStory.set(story.story.index, index))

	const inner = new Set<number>()
	const depth = new Map<number, number>()
	// Text boxes whose holding story this save does not write: written only as they were.
	const unheld = new Set<number>()

	stories.forEach((story, index) => {
		if (story.story.kind !== "textbox") return

		let level = 0

		for (let outer = story.story.parent; outer >= 0; level++) {
			const holder = byStory.get(outer)

			if (holder === undefined) {
				unheld.add(index)

				break
			}

			const next = stories[holder]?.story

			outer = next?.kind === "textbox" ? next.parent : -1
		}

		inner.add(index)
		depth.set(index, level)
	})

	if (inner.size === 0) return undefined

	const written: NestedTextbox[] = []
	const order = [...stories.keys()].sort((a, b) => (depth.get(b) ?? -1) - (depth.get(a) ?? -1) || a - b)

	return {
		order,
		inner: index => inner.has(index),
		record(index, draft) {
			const story = stories[index]

			if (story === undefined) return

			const part = story.part
			const document = part.document
			const container = part.node(story.story.item)
			const from = document.start(container)
			const to = document.end(container)

			// Written as it was: nothing to splice.
			if (draft.out.equals(document.bytes, from, to)) return
			if (unheld.has(index)) throw new DocxWriteError("unsupported", "A text box is written only with the story that holds it")

			const copies = textboxCopies(part, container)

			if (typeof copies === "string")
				throw new DocxWriteError("unsupported", `A text box of ${part.name} cannot be written (${copies})`)

			const bytes = draft.out.finish()
			const textbox: NestedTextbox = {
				part,
				splices: [
					{ from, to, bytes },
					...copies.map(node => ({
						from: document.start(node),
						to: document.end(node),
						bytes: fallbackCopy(part, bytes, ids, node)
					}))
				].sort((a, b) => a.from - b.from),
				bytes,
				used: 0
			}
			const list = lastWritten.get(part)

			written.push(textbox)

			if (list === undefined) lastWritten.set(part, [textbox])
			else list.push(textbox)
		},
		nested(part, from, to) {
			let out: NestedSplice[] | undefined

			for (const textbox of written) {
				if (textbox.part !== part) continue

				for (const splice of textbox.splices) {
					if (splice.to <= from || splice.from >= to) continue
					if (splice.from < from || splice.to > to)
						throw new DocxWriteError("unsupported", `A rewritten text box of ${part.name} is copied in part`)

					;(out ??= []).push(splice)
					textbox.used++
				}
			}

			if (out === undefined) return undefined

			out.sort((a, b) => a.from - b.from || b.to - a.to)

			// Only the outermost: a text box rewritten inside another rewritten one is in that one's bytes.
			const outermost: NestedSplice[] = []

			for (const splice of out) {
				const last = outermost[outermost.length - 1]

				if (last !== undefined && splice.from < last.to) continue

				outermost.push(splice)
			}

			return outermost
		}
	}
}
