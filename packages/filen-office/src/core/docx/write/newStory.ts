import type { Node } from "prosemirror-model"
import { NodeKind, decodeUtf8, type XmlDocument } from "../../xml/index"
import type { Relationship } from "../../opc/index"
import type { DocxEditContext, NewStoryKind } from "../edit/context"
import { NO_SOURCE, docxSchema } from "../edit/schema"
import { Ns } from "../model/namespaces"
import { openTagWithout, Out } from "./markup"
import type { StoryWriting } from "./story"
import { DocxWriteError, paraId as checkParaId } from "./values"

// The container of a story the source does not have (a new comment, note or header).
export interface NewStoryContainer {
	// Its start and end tags (<w:comment w:id="4" ...>, </w:comment>), with the prefixes the part binds.
	readonly open: string
	readonly close: string
	// Where it goes in its part's bytes: the start of a child of the root, or the end of the root's content.
	// In a part the save creates, `at` and `to` may instead span the root, which the container then is (a
	// header).
	readonly at: number
	readonly to?: number
}

const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

// A part a new story creates: its content type and the type of the main part's relationship to it.
export const NEW_STORY_PARTS: Readonly<Record<NewStoryKind, { readonly contentType: string; readonly relationship: string }>> = {
	comment: { contentType: `${WML}.comments+xml`, relationship: `${RELATIONSHIPS}/comments` },
	footnote: { contentType: `${WML}.footnotes+xml`, relationship: `${RELATIONSHIPS}/footnotes` },
	endnote: { contentType: `${WML}.endnotes+xml`, relationship: `${RELATIONSHIPS}/endnotes` },
	header: { contentType: `${WML}.header+xml`, relationship: `${RELATIONSHIPS}/header` },
	footer: { contentType: `${WML}.footer+xml`, relationship: `${RELATIONSHIPS}/footer` }
}

// The comments and notes a save's writers wrote new references to (SaveParts.state key), as referenceKey
// gives them: an added story nothing refers to is left out of the save.
export const WRITTEN_REFERENCES: object = {}

export function referenceKey(kind: string, id: string): string {
	return `${kind}:${id}`
}

// Parts a copy of a story part gets copies of, by relationship type (its last segment), as Word copies them
// with a header or footer: a chart (and what it holds), an embedded or linked object, an ActiveX control, a
// diagram. Pictures and other media stay shared.
const OWNED_BY_COPY: ReadonlySet<string> = new Set([
	"chart",
	"chartUserShapes",
	"chartStyle",
	"chartColorStyle",
	"themeOverride",
	"oleObject",
	"package",
	"control",
	"activeXControlBinary",
	"diagramData",
	"diagramLayout",
	"diagramQuickStyle",
	"diagramColors",
	"diagramDrawing"
])

export function ownedByCopy(relationship: Relationship, pkg: { has(name: string): boolean }): boolean {
	const type = relationship.type

	return (
		!relationship.external &&
		relationship.partName !== undefined &&
		OWNED_BY_COPY.has(type.slice(type.lastIndexOf("/") + 1)) &&
		pkg.has(relationship.partName)
	)
}

// Relationship ids a copied note names the parts it owns by (stories/clone.ts cloneNote): the save copies each
// part and writes the id of the notes part's relationship to the copy in their place (write/save.ts).
export const COPIED_RELATIONSHIP = "filenCopiedPart"

// A name for a copy of part `name` next to it, numbered as Word numbers its parts (chart1.xml -> chart2.xml).
export function copiedPartName(name: string, free: (candidate: string) => boolean): string {
	const slash = name.lastIndexOf("/")
	const dot = name.lastIndexOf(".")
	const end = dot > slash ? dot : name.length
	const stem = /^(.*?)(\d*)$/.exec(name.slice(0, end))?.[1] ?? name.slice(0, end)

	for (let n = 1; ; n++) {
		const candidate = `${stem}${n}${name.slice(end)}`

		if (free(candidate)) return candidate
	}
}

// Revision ids copies of stories took (stories/clone.ts): above the source's, so a save numbering its own
// revisions from the highest (save.ts) never repeats one.
const CLONED_REVISIONS = new WeakMap<DocxEditContext, number>()

// `count` new revision ids for a copy; the first of them.
export function cloneRevisionIds(context: DocxEditContext, count: number): number {
	const last = Math.max(context.model.maxIds.revision, CLONED_REVISIONS.get(context) ?? -1)

	CLONED_REVISIONS.set(context, last + count)

	return last + 1
}

// The highest revision id a copy took, or -1.
export function clonedRevisions(context: DocxEditContext): number {
	return CLONED_REVISIONS.get(context) ?? -1
}

// Whether byte `at` of a part is where a new child of its root can go: the start of a child element, or the
// end of the root's content.
export function isInsertionPoint(document: XmlDocument, at: number): boolean {
	const root = document.documentElement()

	if (document.isSelfClosing(root)) return false
	if (at === document.contentEnd(root)) return true

	const starts = document.starts
	let low = 0
	let high = document.count

	while (low < high) {
		const middle = (low + high) >> 1

		if ((starts[middle] ?? 0) < at) low = middle + 1
		else high = middle
	}

	return low < document.count && starts[low] === at && document.parent(low) === root && document.kinds[low] === NodeKind.Element
}

// A new story: its container around its blocks, every one written new or as a copy.
export function* writeNewStory(w: StoryWriting, out: Out, doc: Node, container: NewStoryContainer): Generator<void, void> {
	out.text(container.open)
	yield* w.writeChildren(out, doc, NO_SOURCE, (child, at) => w.writeBlock(at, child))
	out.text(container.close)
}

// The w14:paraId a feature gave a paragraph (its para_id mark), checked; undefined when it has none.
export function paragraphId(node: Node): string | undefined {
	const id: unknown = docxSchema.marks.para_id.isInSet(node.marks)?.attrs["id"]

	if (id === undefined) return undefined

	checkParaId(id)

	return id as string
}

// Whether source paragraph `source` already has w14:paraId `id`.
export function sameParagraphId(w: StoryWriting, source: number, id: string): boolean {
	return w.part.reader.attr(w.part.node(source), Ns.W14, "paraId")?.toUpperCase() === id.toUpperCase()
}

function w14(w: StoryWriting): string {
	if (w.w14 === "") throw new DocxWriteError("unsupported", `A paragraph id in ${w.part.name}, which does not declare w14`)

	return w.w14
}

// The start tag of a paragraph not written as its source (new, or a copy of source paragraph `source`),
// with the w14:paraId it was given. A copy never repeats its source's id while the source's own paragraph
// (its first occurrence) is written with it; one given another id leaves it to the copy (a comment's split,
// its id moved to its last paragraph).
export function newParagraphTag(w: StoryWriting, source: number, id: string | undefined): string {
	const p = w.names.el("p")

	if (id === undefined) return `<${p}>`

	if (source !== NO_SOURCE && sameParagraphId(w, source, id)) {
		const owner = w.ownerOf(source)
		const ownerId = owner === undefined ? undefined : paragraphId(owner)

		if (owner !== undefined && (ownerId === undefined || sameParagraphId(w, source, ownerId))) return `<${p}>`
	}

	return `<${p} ${w14(w)}:paraId="${id}">`
}

// The start tag of source paragraph `source` with a new w14:paraId (and without its w14:textId, which
// hashed the old paragraph).
export function writeParagraphIdTag(w: StoryWriting, out: Out, source: number, id: string): void {
	const prefix = w14(w)
	const tag = new Out()

	openTagWithout(tag, w.part.document, w.part.node(source), new Set([`${prefix}:paraId`, `${prefix}:textId`]))

	const markup = decodeUtf8(tag.finish())

	out.text(`${markup.slice(0, -1)} ${prefix}:paraId="${id}">`)
}
