import { NodeKind, XmlError, decodeUtf8, encodeUtf8, parseXml, type XmlDocument } from "../../xml/index"
import type { Relationships } from "../../opc/index"
import type { DocxEditContext } from "../edit/context"
import type { DocxEditorSession } from "../edit/session"
import { Ns } from "../model/namespaces"
import { partReader, type PartReader } from "../model/reader"
import type { StoryPart } from "../model/index"
import { newControlId } from "../forms/controls"
import { planNested } from "../write/nested"
import { cloneRevisionIds } from "../write/newStory"
import { DocxWriteError, StoryWriter, type WriteEnvironment } from "../write/story"

// A header or footer copied into a part of its own, as Word copies the one a section showed when the section
// is unlinked from the previous one: everything it holds (pictures, drawings, fields, content controls,
// content kept as bytes) with the relationships they use. What must be unique in a document is made new
// (drawing ids, revision ids, content control ids) or left out (paragraph ids, which Word makes again; bookmarks, comment anchors
// and permission ranges, which stay with the original).

// What a part made from another holds: its bytes and the relationships (ids as in the bytes) it keeps.
export interface StoryTemplate {
	readonly bytes: Uint8Array
	readonly relationships: Relationships
}

// Left out of a copy (the original keeps them).
const LEFT_OUT: ReadonlySet<string> = new Set([
	"bookmarkStart",
	"bookmarkEnd",
	"commentRangeStart",
	"commentRangeEnd",
	"commentReference",
	"permStart",
	"permEnd"
])

// Revision elements whose w:id a copy makes new; ranges keep their start and end paired.
const REVISIONS: ReadonlySet<string> = new Set([
	"ins",
	"del",
	"moveFrom",
	"moveTo",
	"rPrChange",
	"pPrChange",
	"sectPrChange",
	"tblPrChange",
	"trPrChange",
	"tcPrChange",
	"tblGridChange",
	"tblPrExChange",
	"numberingChange",
	"cellIns",
	"cellDel",
	"cellMerge"
])
const RANGES = /^(moveFrom|moveTo|customXmlInsRange|customXmlDelRange|customXmlMoveFromRange|customXmlMoveToRange)(Start|End)$/

// Placeholders the writer leaves where a save still has to place something (stories/sections.ts, the layout's
// pictures and links): a story holding them cannot be written alone.
const PLACEHOLDERS = /filenStory\d|filenMedia\d|filenPlannedLink\d/

interface Edit {
	readonly from: number
	readonly to: number
	readonly text: string
}

// The story's root element as a save would write it now: its source bytes where neither it nor a text box in
// it (`textboxes`, their sessions) is edited, else the writer's output, the text boxes written into it as a save
// writes them (undefined when it holds something only a save can place: a new picture, list or link).
function storyBytes(
	context: DocxEditContext,
	part: StoryPart,
	item: number,
	session: DocxEditorSession | undefined,
	textboxes: readonly DocxEditorSession[]
): Uint8Array | undefined {
	const document = part.document
	const node = part.node(item)
	const edited = (candidate: DocxEditorSession): boolean => candidate.state.doc !== candidate.story.doc

	if (session === undefined || (!edited(session) && !textboxes.some(edited)))
		return document.bytes.subarray(document.start(node), document.end(node))

	// Ids written here are made new below; these only keep them apart.
	let revision = Math.max(0, context.model.maxIds.revision) + 1_000_000
	const sessions = [session, ...textboxes]
	const nested = planNested(
		sessions.map(candidate => candidate.story),
		{ next: space => (space === "revision" ? ++revision : context.ids.next(space)) }
	)
	const env: WriteEnvironment = {
		context,
		nextRevisionId: () => ++revision,
		numId: () => {
			throw new DocxWriteError("unsupported", "A new list in a copied story")
		},
		hyperlink: () => {
			throw new DocxWriteError("unsupported", "A new hyperlink in a copied story")
		},
		...(nested === undefined ? {} : { nested: nested.nested })
	}

	try {
		// Text boxes first, innermost first: each takes the place of its container in what holds it.
		for (const index of nested?.order ?? []) {
			const inner = sessions[index]

			if (inner !== undefined && nested?.inner(index) === true)
				nested.record(index, new StoryWriter(env, inner.story).draftStorySync(inner.state.doc))
		}

		const bytes = new StoryWriter(env, session.story).draftStorySync(session.state.doc).out.finish()

		return PLACEHOLDERS.test(decodeUtf8(bytes)) ? undefined : bytes
	} catch (error) {
		if (error instanceof DocxWriteError) return undefined

		throw error
	}
}

// The value span of attribute {group}local of `node` with the span of the whole attribute (its leading space
// included), or undefined.
function attributeSpan(
	reader: PartReader,
	document: XmlDocument,
	node: number,
	group: Ns,
	local: string
): { readonly from: number; readonly to: number; readonly value: [number, number] } | undefined {
	const value = new Int32Array(2)

	if (!reader.find(node, group, local, value)) return undefined

	const spans = document.attributeSpans(node)
	const bytes = document.bytes

	for (let k = 0; k < spans.length; k += 4) {
		if (spans[k + 2] !== value[0]) continue

		let from = spans[k] ?? 0

		while (from > 0 && (bytes[from - 1] === 0x20 || bytes[from - 1] === 0x09 || bytes[from - 1] === 0x0a || bytes[from - 1] === 0x0d))
			from--

		return { from, to: (spans[k + 3] ?? 0) + 1, value: [value[0] ?? 0, value[1] ?? 0] }
	}

	return undefined
}

// What a copy changes in the story's markup, in document order.
function cloneEdits(context: DocxEditContext, document: XmlDocument): Edit[] {
	const reader = partReader(document)
	const edits: Edit[] = []
	const ranges = new Map<string, number>()
	let skipTo = -1

	for (let node = 0; node < document.count; node++) {
		if (document.kinds[node] !== NodeKind.Element || document.start(node) < skipTo) continue

		const group = reader.group(node)
		const local = document.localName(node)

		if (group === Ns.WP && local === "docPr") {
			const id = attributeSpan(reader, document, node, Ns.None, "id")

			if (id !== undefined) edits.push({ from: id.value[0], to: id.value[1], text: String(context.ids.next("docPr")) })

			continue
		}

		if (group !== Ns.W) continue

		// A content control's w:id is the document's alone (Word gives a pasted control a new one).
		if (local === "id" && reader.is(document.parent(node), Ns.W, "sdtPr")) {
			const value = attributeSpan(reader, document, node, Ns.W, "val")

			if (value !== undefined) edits.push({ from: value.value[0], to: value.value[1], text: newControlId() })

			continue
		}

		if (LEFT_OUT.has(local)) {
			edits.push({ from: document.start(node), to: document.end(node), text: "" })
			skipTo = document.end(node)

			continue
		}

		if (local === "p" || local === "tr") {
			for (const name of ["paraId", "textId"]) {
				const attribute = attributeSpan(reader, document, node, Ns.W14, name)

				if (attribute !== undefined) edits.push({ from: attribute.from, to: attribute.to, text: "" })
			}

			continue
		}

		const range = RANGES.exec(local)

		if (range === null && !REVISIONS.has(local)) continue

		const id = attributeSpan(reader, document, node, Ns.W, "id")

		if (id === undefined) continue

		let next: number

		if (range === null) {
			next = cloneRevisionIds(context, 1)
		} else {
			// A range's end takes the id its start got.
			const key = `${range[1] ?? ""}:${decodeUtf8(document.bytes, id.value[0], id.value[1])}`
			const known = ranges.get(key)

			next = known ?? cloneRevisionIds(context, 1)
			ranges.set(key, next)
		}

		edits.push({ from: id.value[0], to: id.value[1], text: String(next) })
	}

	// Attribute edits of one start tag come in attribute order; elements left out never overlap them.
	return edits.sort((a, b) => a.from - b.from)
}

function spliced(bytes: Uint8Array, edits: readonly Edit[]): Uint8Array {
	const pieces: Uint8Array[] = []
	let at = 0
	let length = 0

	for (const edit of edits) {
		if (edit.from < at) continue

		const kept = bytes.subarray(at, edit.from)
		const text = encodeUtf8(edit.text)

		pieces.push(kept, text)
		length += kept.length + text.length
		at = edit.to
	}

	const rest = bytes.subarray(at)

	pieces.push(rest)
	length += rest.length

	const out = new Uint8Array(length)
	let offset = 0

	for (const piece of pieces) {
		out.set(piece, offset)
		offset += piece.length
	}

	return out
}

// The part a copy of header or footer story `item` of `part` makes: the part's prolog and the story as it is
// now (`session`'s document when one edits it, with the sessions of the text boxes in it, `textboxes`), made new
// where a document needs it unique. Undefined when the story cannot be written alone (see storyBytes) or its
// markup does not read.
export function cloneStory(
	context: DocxEditContext,
	part: StoryPart,
	item: number,
	session: DocxEditorSession | undefined,
	textboxes: readonly DocxEditorSession[] = []
): StoryTemplate | undefined {
	const story = storyBytes(context, part, item, session, textboxes)

	if (story === undefined) return undefined

	const source = part.document
	const prolog = source.bytes.subarray(0, source.start(source.documentElement()))
	const whole = new Uint8Array(prolog.length + story.length)

	whole.set(prolog, 0)
	whole.set(story, prolog.length)

	let document: XmlDocument

	try {
		document = parseXml(whole)
	} catch (error) {
		if (error instanceof XmlError) return undefined

		throw error
	}

	return { bytes: spliced(whole, cloneEdits(context, document)), relationships: part.relationships }
}
