import type { Node } from "prosemirror-model"
import type { StoryDoc } from "../../edit/build"
import type { DocxEditContext } from "../../edit/context"
import { NO_SOURCE, docxSchema } from "../../edit/schema"
import type { DocxEditorSession } from "../../edit/session"
import { Item, type StoryPart } from "../../model/index"
import { W_NAMESPACE } from "../source"
import { DocxWriteError } from "../values"
import type { PartContributor, SaveParts, WrittenStories } from "./index"

// Footnotes and endnotes parts. A new note's w:footnote / w:endnote (its id) is appended to its part (the
// document's, or one the save makes from a template holding Word's separator notes); a source note whose
// every reference the edits deleted leaves its part, as Word deletes a note with its reference. Separator
// notes are never touched.

const nodes = docxSchema.nodes

type NoteKind = "footnote" | "endnote"

function prefixOf(part: StoryPart): string {
	const prefix = part.reader.namespaces.prefixFor(part.document.documentElement(), W_NAMESPACE)

	if (prefix === undefined) throw new DocxWriteError("unsupported", `${part.name} does not bind WordprocessingML`)

	return prefix
}

function qualified(prefix: string, local: string): string {
	return prefix === "" ? local : `${prefix}:${local}`
}

// A new note's container, appended to its part.
function container(story: StoryDoc) {
	const own = story.story
	const kind = own.kind

	if ((kind !== "footnote" && kind !== "endnote") || own.item >= 0 || own.id === undefined) return undefined
	if (!/^\d{1,10}$/.test(own.id) || Number(own.id) < 1) throw new DocxWriteError("missingSource", `A new ${kind} without an id`)

	const document = story.part.document
	const prefix = prefixOf(story.part)
	const tag = qualified(prefix, kind)

	return {
		open: `<${tag} ${qualified(prefix, "id")}="${own.id}">`,
		close: `</${tag}>`,
		at: document.contentEnd(document.documentElement())
	}
}

type References = Readonly<Record<NoteKind, ReadonlyMap<string, readonly number[]>>>

// The note references among a story part's items, by kind and note id (read once per part).
const REFERENCES = new WeakMap<StoryPart, References>()

function referencesOf(part: StoryPart): References {
	let found = REFERENCES.get(part)

	if (found === undefined) {
		const footnote = new Map<string, number[]>()
		const endnote = new Map<string, number[]>()

		for (let item = 0; item < part.count; item++) {
			const kind = part.kinds[item]
			const map = kind === Item.FootnoteReference ? footnote : kind === Item.EndnoteReference ? endnote : undefined
			const id = map === undefined ? undefined : part.id(item)

			if (map === undefined || id === undefined) continue

			const list = map.get(id)

			if (list === undefined) map.set(id, [item])
			else list.push(item)
		}

		found = { footnote, endnote }
		REFERENCES.set(part, found)
	}

	return found
}

// What a story document holds of note references: those its note_ref nodes make ("kind:id"), and the source
// ranges written as their bytes (atoms, locked blocks), which keep whatever reference is inside them.
interface HeldReferences {
	readonly ids: ReadonlySet<string>
	readonly ranges: readonly number[]
}

const HELD = new WeakMap<Node, HeldReferences>()

function heldBy(doc: Node, part: StoryPart): HeldReferences {
	let held = HELD.get(doc)

	if (held === undefined) {
		const ids = new Set<string>()
		const ranges: number[] = []

		doc.descendants(node => {
			if (node.type === nodes.note_ref) {
				ids.add(`${String(node.attrs["kind"])}:${String(node.attrs["id"])}`)

				return false
			}

			const own: unknown = node.attrs["src"]

			if ((node.isAtom || node.type === nodes.locked_block) && typeof own === "number" && own !== NO_SOURCE) {
				ranges.push(own, part.end(own))

				return false
			}

			return true
		})

		held = { ids, ranges }
		HELD.set(doc, held)
	}

	return held
}

function holds(held: HeldReferences, items: readonly number[]): boolean {
	for (const item of items) {
		for (let k = 0; k < held.ranges.length; k += 2) if (item >= (held.ranges[k] ?? 0) && item < (held.ranges[k + 1] ?? 0)) return true
	}

	return false
}

// Source notes the edits took every reference of (in the stories the save writes), by part: the save cuts them
// out. A reference a paste or an undo put back, anywhere the save writes, keeps its note.
export function orphanedNotes(
	context: DocxEditContext,
	sessions: readonly DocxEditorSession[],
	docs: readonly Node[]
): Map<StoryPart, number[]> {
	const out = new Map<StoryPart, number[]>()
	const model = context.model

	if (model.footnotes.size === 0 && model.endnotes.size === 0) return out

	// The stories that may hold references, written, by part.
	const written = new Map<number, { doc: Node; part: StoryPart }[]>()

	sessions.forEach((session, k) => {
		const story = session.story.story

		if (story.kind === "footnote" || story.kind === "endnote" || story.kind === "comment" || story.item < 0) return

		const list = written.get(story.part) ?? []

		list.push({ doc: docs[k] ?? session.state.doc, part: session.story.part })
		written.set(story.part, list)
	})

	const referencedAnywhere = (key: string): boolean => {
		for (const list of written.values()) for (const entry of list) if (heldBy(entry.doc, entry.part).ids.has(key)) return true

		return false
	}

	for (const kind of ["footnote", "endnote"] as const) {
		for (const [id, story] of kind === "footnote" ? model.footnotes : model.endnotes) {
			if (story.noteType !== "normal" || story.item < 0) continue

			const key = `${kind}:${id}`
			let found = false
			let referenced = false

			for (const [index, part] of model.parts.entries()) {
				const items = referencesOf(part)[kind].get(id)

				if (items === undefined) continue

				found = true

				// A part none of whose stories is written keeps its references; a written story keeps those in
				// bytes it copies.
				const entries = written.get(index)

				if (entries === undefined || entries.some(entry => holds(heldBy(entry.doc, entry.part), items))) {
					referenced = true

					break
				}
			}

			if (!found || referenced || referencedAnywhere(key)) continue

			const part = model.parts[story.part]

			if (part === undefined) continue

			const list = out.get(part)

			if (list === undefined) out.set(part, [story.item])
			else list.push(story.item)
		}
	}

	return out
}

export const noteParts: PartContributor = {
	container,
	removed: (parts: SaveParts, written: Omit<WrittenStories, "stories">) => orphanedNotes(parts.context, written.sessions, written.docs)
}
