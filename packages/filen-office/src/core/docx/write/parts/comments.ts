import type { Node } from "prosemirror-model"
import { relativeTarget } from "../../../opc/index"
import { XmlEditor, encodeUtf8, escapeAttribute, type XmlDocument } from "../../../xml/index"
import type { StoryDoc } from "../../edit/build"
import type { DocxEditContext } from "../../edit/context"
import { docxSchema } from "../../edit/schema"
import type { DocxEditorSession } from "../../edit/session"
import { COMMENT_NAMESPACES, COMMENT_PARTS, readCommentsIds, type CommentPartKind } from "../../model/comments"
import type { CommentInfo, StoryPart } from "../../model/index"
import { Ns } from "../../model/namespaces"
import { partReader } from "../../model/reader"
import { commentRecords, orphanedComments, type DocxCommentRecord } from "../../review/comments"
import { W_NAMESPACE } from "../source"
import { DocxWriteError, paraId as checkParaId } from "../values"
import type { PartContributor, RootDeclarations, SaveParts, WrittenStories } from "./index"

// The comments part family. A new comment's w:comment (id, author, date, initials from its own story's
// record) is appended to the comments part; commentsExtended gets an entry per new comment (its last
// paragraph's id, the comment it replies to, whether it is resolved) and the resolved state of source
// comments that changed; commentsIds, commentsExtensible and people get entries for new comments where the
// document has those parts (made only when missing: commentsExtended, which threads need). A deleted source
// comment (by the command, or with all of its anchors) leaves every part of the family; its replies answer
// their thread's first comment.

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const W15_URIS = [COMMENT_PARTS.commentsExtended.uri, "http://schemas.microsoft.com/office/word/2010/11/wordml"]

function prefixOf(document: XmlDocument, uris: readonly string[]): string | undefined {
	const root = document.documentElement()
	const namespaces = partReader(document).namespaces

	for (const uri of uris) {
		const prefix = namespaces.prefixFor(root, uri)

		if (prefix !== undefined) return prefix
	}

	return undefined
}

function qualified(prefix: string, local: string): string {
	return prefix === "" ? local : `${prefix}:${local}`
}

function attributes(prefix: string, values: readonly (readonly [string, string | undefined])[]): string {
	return values
		.filter((entry): entry is readonly [string, string] => entry[1] !== undefined)
		.map(([name, value]) => ` ${qualified(prefix, name)}="${escapeAttribute(value)}"`)
		.join("")
}

// The id of a new comment's story (its w:id), when `story` is one.
function newCommentId(story: StoryDoc): string | undefined {
	return story.story.kind === "comment" && story.story.item < 0 ? story.story.id : undefined
}

function lastParagraphId(doc: Node): string | undefined {
	const last = doc.lastChild
	const id: unknown = last === null ? undefined : docxSchema.marks.para_id.isInSet(last.marks)?.attrs["id"]

	return typeof id === "string" ? id : undefined
}

// What a save knows of the comments: the records of every story it writes, each comment's own story.
interface Facts {
	readonly records: Map<string, DocxCommentRecord>
	// New comments the save writes, by id, with their documents.
	readonly added: Map<string, Node>
	// Sessions of comment stories by comment id, with their documents as written.
	readonly stories: Map<string, Node>
	readonly sources: Map<string, CommentInfo>
}

function factsOf(context: DocxEditContext, written: WrittenStories): Facts {
	const records = new Map<string, DocxCommentRecord>()
	const stories = new Map<string, Node>()
	const sources = new Map(context.model.comments.map(comment => [comment.id, comment]))
	const writtenSessions = new Set<DocxEditorSession>(written.stories.map(story => story.session))
	const added = new Map<string, Node>()

	written.sessions.forEach((session, k) => {
		const doc = written.docs[k] ?? session.state.doc
		const story = session.story.story

		if (story.kind === "comment") {
			if (story.id === undefined) return

			stories.set(story.id, doc)

			if (story.item < 0 && writtenSessions.has(session)) added.set(story.id, doc)

			return
		}

		for (const [id, record] of Object.entries(commentRecords(doc))) records.set(id, record)
	})

	return { records, added, stories, sources }
}

function paraIdOf(facts: Facts, id: string): string | undefined {
	const doc = facts.stories.get(id)
	const own = doc === undefined ? undefined : lastParagraphId(doc)

	return own ?? facts.sources.get(id)?.paraId
}

// Random durable ids of the new comments, kept per document so saves agree.
const DURABLE_IDS = new WeakMap<DocxEditContext, Map<string, string>>()

function durableId(context: DocxEditContext, id: string, taken: ReadonlySet<string>): string {
	let ids = DURABLE_IDS.get(context)

	if (ids === undefined) {
		ids = new Map()
		DURABLE_IDS.set(context, ids)
	}

	let out = ids.get(id)

	if (out === undefined) {
		const used = new Set([...taken, ...ids.values()])

		out = context.ids.paraId(value => used.has(value.toString(16).toUpperCase().padStart(8, "0")))
		ids.set(id, out)
	}

	return out
}

// A part of the family the package has, with an editor over it.
interface FamilyPart {
	readonly name: string
	readonly document: XmlDocument
	readonly editor: XmlEditor
	readonly prefix: string
	// The namespace declaration the root gets when the part is written (it did not bind the namespace).
	readonly declaration: readonly [string, string] | undefined
	changed: boolean
}

async function familyPart(parts: SaveParts, kind: CommentPartKind): Promise<FamilyPart | undefined> {
	const name = parts.pkg.part(kind)
	const document = name === undefined ? undefined : await parts.xml(name)

	if (name === undefined || document === undefined) return undefined

	const info = COMMENT_PARTS[kind]
	const prefix = prefixOf(document, kind === "commentsExtended" || kind === "people" ? W15_URIS : [info.uri])

	if (prefix !== undefined) return { name, document, editor: new XmlEditor(document), prefix, declaration: undefined, changed: false }

	// A part that does not bind the namespace (an older writer's): its root declares it, under a prefix the
	// root does not use.
	const bound = partReader(document).namespaces.inScope(document.documentElement())
	let declared: string = info.prefix

	for (let n = 1; bound.has(declared); n++) declared = `${info.prefix}${String(n)}`

	return {
		name,
		document,
		editor: new XmlEditor(document),
		prefix: declared,
		declaration: [`xmlns:${declared}`, info.uri],
		changed: false
	}
}

// The root's children `local` of namespace group `ns`.
function children(document: XmlDocument, ns: Ns, local: string): number[] {
	const reader = partReader(document)
	const out: number[] = []

	for (let child = document.firstChildElement(document.documentElement()); child >= 0; child = document.nextSiblingElement(child)) {
		if (reader.is(child, ns, local)) out.push(child)
	}

	return out
}

function attribute(document: XmlDocument, node: number, ns: Ns, local: string): string | undefined {
	return partReader(document).attr(node, ns, local)?.toUpperCase()
}

// Writes what the family part holds now: the package's part edited, or a new part (with a relationship from
// the main part, unless one names a part the package lacks) when it has entries.
async function finish(parts: SaveParts, kind: CommentPartKind, part: FamilyPart | undefined, appended: readonly string[]): Promise<void> {
	const info = COMMENT_PARTS[kind]

	if (part !== undefined) {
		if (!part.changed && appended.length === 0) return

		const root = part.document.documentElement()

		if (part.declaration !== undefined) part.editor.setAttribute(root, part.declaration[0], part.declaration[1])
		if (appended.length > 0) part.editor.append(root, appended.join(""))

		parts.replace(part.name, part.editor.serialize())

		return
	}

	if (appended.length === 0) return

	const named = parts.pkg.part(kind)
	const name = named ?? parts.freeName(info.name)
	const root = qualified(info.prefix, info.root)

	parts.add(
		name,
		encodeUtf8(`${XML_DECLARATION}<${root} xmlns:${info.prefix}="${info.uri}">${appended.join("")}</${root}>`),
		info.contentType
	)

	if (named !== undefined) return

	await parts.loadRelationships(parts.pkg.mainPart)
	parts.relationship(parts.pkg.mainPart, { type: info.relationship, target: relativeTarget(parts.pkg.mainPart, name) })
}

async function contribute(parts: SaveParts, written: WrittenStories): Promise<void> {
	const context = parts.context
	const facts = factsOf(context, written)
	const sources = facts.sources
	// Source comments whose resolved state the edits changed, and those deleted.
	const resolved = new Map<string, boolean>()
	const deleted = new Set<string>()

	for (const [id, record] of facts.records) {
		const source = sources.get(id)

		if (source === undefined) continue
		if (record.deleted === true) deleted.add(id)
		else if (record.done !== undefined && record.done !== source.done) resolved.set(id, record.done)
	}

	const cut = deletedComments(parts, written)

	if (facts.added.size === 0 && resolved.size === 0 && deleted.size === 0 && cut.size === 0) return

	// A deleted comment leaves the comments part (the save cuts its w:comment out: removed below), and its
	// entries leave the other parts of the family.
	const removed = new Set<string>()

	for (const [part, items] of cut) {
		for (const story of context.model.stories)
			if (story.part === part.index && items.includes(story.item) && story.id !== undefined) removed.add(story.id)
	}

	for (const id of removed) resolved.delete(id)

	const removedParaIds = new Set(
		[...removed].map(id => sources.get(id)?.paraId?.toUpperCase()).filter((id): id is string => id !== undefined)
	)

	await extended(parts, facts, resolved, removedParaIds)

	const idsPart = await familyPart(parts, "commentsIds")
	const durable = new Map<string, string>()

	if (idsPart !== undefined) {
		const existing = readCommentsIds(idsPart.document)
		const taken = new Set(existing.values())
		const appended: string[] = []
		const removedDurable = new Set<string>()

		for (const node of children(idsPart.document, Ns.W16Cid, "commentId")) {
			const paraId = attribute(idsPart.document, node, Ns.W16Cid, "paraId")

			if (paraId !== undefined && removedParaIds.has(paraId)) {
				idsPart.editor.remove(node)
				idsPart.changed = true

				const id = existing.get(paraId)

				if (id !== undefined) removedDurable.add(id)
			}
		}

		for (const [id, doc] of facts.added) {
			const paraId = lastParagraphId(doc)

			if (paraId === undefined) continue

			const value = durableId(context, id, taken)

			durable.set(id, value)
			appended.push(
				`<${qualified(idsPart.prefix, "commentId")}${attributes(idsPart.prefix, [
					["paraId", paraId],
					["durableId", value]
				])}/>`
			)
		}

		await finish(parts, "commentsIds", idsPart, appended)

		const extensible = await familyPart(parts, "commentsExtensible")

		if (extensible !== undefined) {
			const added: string[] = []

			for (const node of children(extensible.document, Ns.W16Cex, "commentExtensible")) {
				const id = attribute(extensible.document, node, Ns.W16Cex, "durableId")

				if (id !== undefined && removedDurable.has(id)) {
					extensible.editor.remove(node)
					extensible.changed = true
				}
			}

			for (const [id, value] of durable) {
				const record = commentRecords(facts.added.get(id) as Node)[id]
				const date = record?.dateUtc ?? record?.date

				added.push(
					`<${qualified(extensible.prefix, "commentExtensible")}${attributes(extensible.prefix, [
						["durableId", value],
						["dateUtc", date]
					])}/>`
				)
			}

			await finish(parts, "commentsExtensible", extensible, added)
		}
	}

	const people = await familyPart(parts, "people")

	if (people !== undefined) {
		const reader = partReader(people.document)
		const known = new Set(children(people.document, Ns.W15, "person").map(node => reader.attr(node, Ns.W15, "author")))
		const added: string[] = []

		for (const [id, doc] of facts.added) {
			const author = commentRecords(doc)[id]?.author

			if (author === undefined || known.has(author)) continue

			known.add(author)
			added.push(`<${qualified(people.prefix, "person")}${attributes(people.prefix, [["author", author]])}/>`)
		}

		await finish(parts, "people", people, added)
	}
}

const ORPHANS = {}

// Source comments the edits deleted (the command, or all of their anchors), and the comments part's items of
// those it holds.
function deletedComments(parts: SaveParts, written: Omit<WrittenStories, "stories">): Map<StoryPart, number[]> {
	const context = parts.context
	const sources = new Set(context.model.comments.map(comment => comment.id))
	const ids = new Set<string>(parts.state(ORPHANS, () => orphanedComments(context, written.sessions, written.docs)))
	const items = new Map<StoryPart, number[]>()

	written.sessions.forEach((session, k) => {
		if (session.story.story.kind === "comment") return

		for (const [id, record] of Object.entries(commentRecords(written.docs[k] ?? session.state.doc)))
			if (record.deleted === true && sources.has(id)) ids.add(id)
	})

	if (ids.size === 0) return items

	const model = context.model
	const commentsName = parts.pkg.part("comments")
	const part = commentsName === undefined ? undefined : model.parts.find(candidate => candidate.name === commentsName)

	if (part === undefined) return items

	for (const story of model.stories) {
		if (story.kind !== "comment" || story.part !== part.index || story.item < 0 || story.id === undefined || !ids.has(story.id))
			continue

		const list = items.get(part)

		if (list === undefined) items.set(part, [story.item])
		else list.push(story.item)
	}

	return items
}

// commentsExtended: an entry per new comment that has a paragraph id, the resolved state of source comments
// that changed (an entry added where they had none), deleted comments' entries removed.
async function extended(
	parts: SaveParts,
	facts: Facts,
	resolved: ReadonlyMap<string, boolean>,
	removed: ReadonlySet<string>
): Promise<void> {
	const part = await familyPart(parts, "commentsExtended")
	const prefix = part?.prefix ?? COMMENT_PARTS.commentsExtended.prefix
	const entry = (paraId: string, parent: string | undefined, done: boolean): string =>
		`<${qualified(prefix, "commentEx")}${attributes(prefix, [
			["paraId", paraId],
			["paraIdParent", parent],
			["done", done ? "1" : "0"]
		])}/>`
	const appended: string[] = []
	const pending = new Map(resolved)

	if (part !== undefined) {
		const document = part.document
		const byParaId = new Map<string, string>()

		for (const [id] of resolved) {
			const paraId = paraIdOf(facts, id)?.toUpperCase()

			if (paraId !== undefined) byParaId.set(paraId, id)
		}

		const entries = children(document, Ns.W15, "commentEx")
		const parents = new Map<string, string>()

		if (removed.size > 0) {
			for (const node of entries) {
				const paraId = attribute(document, node, Ns.W15, "paraId")
				const parent = attribute(document, node, Ns.W15, "paraIdParent")

				if (paraId !== undefined && parent !== undefined) parents.set(paraId, parent)
			}
		}

		for (const node of entries) {
			const paraId = attribute(document, node, Ns.W15, "paraId")

			if (paraId === undefined) continue

			if (removed.has(paraId)) {
				part.editor.remove(node)
				part.changed = true

				continue
			}

			const parent = parents.get(paraId)

			// A reply to a removed comment answers its thread's first comment now (where that is kept).
			if (parent !== undefined && removed.has(parent)) {
				const root = survivingRoot(parents, parent, removed)

				if (root === undefined) part.editor.removeAttribute(node, qualified(prefix, "paraIdParent"))
				else part.editor.setAttribute(node, qualified(prefix, "paraIdParent"), root)

				part.changed = true
			}

			const id = byParaId.get(paraId)

			if (id === undefined || !pending.has(id)) continue

			part.editor.setAttribute(node, qualified(prefix, "done"), pending.get(id) === true ? "1" : "0")
			part.changed = true
			pending.delete(id)
		}
	}

	for (const [id, done] of pending) {
		const paraId = paraIdOf(facts, id)

		if (paraId === undefined) throw new DocxWriteError("missingSource", `Comment ${id} is resolved but has no paragraph id`)

		checkParaId(paraId)
		appended.push(entry(paraId, undefined, done))
	}

	for (const [id, doc] of facts.added) {
		const paraId = lastParagraphId(doc)
		const record = facts.records.get(id)

		if (paraId === undefined) {
			if (record?.parent !== undefined || record?.done === true)
				throw new DocxWriteError("missingSource", `Comment ${id} is in a thread or resolved but has no paragraph id`)

			continue
		}

		const parent = record?.parent === undefined ? undefined : paraIdOf(facts, record.parent)

		if (record?.parent !== undefined && parent === undefined)
			throw new DocxWriteError("missingSource", `Comment ${id} replies to ${record.parent}, which has no paragraph id`)

		appended.push(entry(paraId, parent, record?.done === true))
	}

	await finish(parts, "commentsExtended", part, appended)
}

// The first comment of the thread of paragraph `paraId` (it and its ancestors by `parents`) that is not
// removed, the thread's topmost; undefined when every one is.
function survivingRoot(parents: ReadonlyMap<string, string>, paraId: string, removed: ReadonlySet<string>): string | undefined {
	const seen = new Set<string>()
	let found: string | undefined

	for (let at: string | undefined = paraId; at !== undefined && !seen.has(at); at = parents.get(at)) {
		seen.add(at)

		if (!removed.has(at)) found = at
	}

	return found
}

// A new comment's w:comment, appended to its part.
function container(story: StoryDoc, doc: Node) {
	const id = newCommentId(story)

	if (id === undefined) return undefined

	const document = story.part.document
	const prefix = prefixOf(document, [W_NAMESPACE])
	const record = commentRecords(doc)[id]

	if (prefix === undefined) throw new DocxWriteError("unsupported", `${story.part.name} does not bind WordprocessingML`)
	if (record?.author === undefined) throw new DocxWriteError("missingSource", `A new comment ${id} without an author`)

	const tag = qualified(prefix, "comment")

	return {
		open: `<${tag}${attributes(prefix, [
			["id", id],
			["author", record.author],
			["date", record.date],
			["initials", record.initials]
		])}>`,
		close: `</${tag}>`,
		at: document.contentEnd(document.documentElement())
	}
}

const MARKUP_COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006"

// The comments part's root as the save must write it when it lacks w14 and a paragraph of it gets a
// w14:paraId (a 2007-era part): w14 declared and marked ignorable (with mc declared where it is not), every
// other attribute as the source has it. Undefined when nothing needs declaring.
export function commentDeclarations(part: StoryPart, written: WrittenStories): RootDeclarations | undefined {
	const document = part.document
	const reader = partReader(document)
	const root = document.documentElement()

	if (!reader.is(root, Ns.W, "comments") || prefixOf(document, [COMMENT_NAMESPACES.w14]) !== undefined) return undefined

	const ids = written.sessions.some((session, k) => {
		const doc = written.docs[k] ?? session.state.doc
		let found = false

		if (session.story.part !== part) return false

		doc.forEach(block => {
			if (docxSchema.marks.para_id.isInSet(block.marks) !== undefined) found = true
		})

		return found
	})

	if (!ids) return undefined

	const bound = reader.namespaces.inScope(root)
	const free = (base: string): string => {
		let prefix = base

		for (let n = 1; bound.has(prefix); n++) prefix = `${base}${String(n)}`

		return prefix
	}
	const w14 = free("w14")
	const mc = prefixOf(document, [MARKUP_COMPATIBILITY]) ?? free("mc")
	const ignorable = `${mc}:Ignorable`
	let open = `<${document.name(root)}`
	let marked = false

	for (const attribute of document.attributes(root)) {
		const value = attribute.name === ignorable ? `${attribute.value} ${w14}`.trim() : attribute.value

		marked ||= attribute.name === ignorable
		open += ` ${attribute.name}="${escapeAttribute(value)}"`
	}

	open += ` xmlns:${w14}="${COMMENT_NAMESPACES.w14}"`

	if (!bound.has(mc)) open += ` xmlns:${mc}="${MARKUP_COMPATIBILITY}"`
	if (!marked) open += ` ${ignorable}="${w14}"`

	return { open: `${open}>`, prefixes: new Map([[COMMENT_NAMESPACES.w14, w14]]) }
}

const contributor: PartContributor = {
	attributes: ["comments"],
	container,
	contribute,
	removed: deletedComments,
	declarations: commentDeclarations
}

export const commentParts: PartContributor = contributor
