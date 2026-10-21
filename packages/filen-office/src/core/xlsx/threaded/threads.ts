import { SHEET_MAX_COLS, SHEET_MAX_ROWS } from "../../sheet/protocol"
import { encodeUtf8, escapeAttribute, escapeText, XmlEditor, type XmlDocument } from "../../xml/index"
import { encodeXstring } from "../edit/document"
import { cellKeyOf } from "../edit/history"
import { decodeXstring } from "../model/strings"
import { DECLARATION, THREADED_NAMESPACE } from "../comments/parts"
import { cellRef, parseCellRef } from "../write/text"

// Threaded comments ([MS-XLSX] §2.1.17, §2.6.207 CT_ThreadedComment): per sheet, comments with a cell, a time, an
// author (a person of the workbook's persons part, §2.1.18) and plain text; a reply names its thread's first comment
// as parentId. Mentions (§2.6.209) point into the text by offsets.

export interface ThreadComment {
	readonly id: string
	readonly parentId: string | null
	readonly personId: string
	// dT as written (xsd:dateTime, Excel's local time).
	readonly date: string
	readonly text: string
	readonly done: boolean
	readonly mentions: boolean
	readonly node: number
}

export interface Thread {
	readonly row: number
	readonly col: number
	readonly root: ThreadComment
	readonly replies: readonly ThreadComment[]
}

export interface ThreadedPart {
	readonly document: XmlDocument
	// The prefix of the part's own elements, with its colon.
	readonly prefix: string
	// The first thread of each cell, by cellKeyOf.
	readonly threads: ReadonlyMap<number, Thread>
	// Every comment by id (upper case), with its thread's cell.
	readonly comments: ReadonlyMap<string, { readonly comment: ThreadComment; readonly key: number }>
}

export interface Person {
	readonly id: string
	readonly displayName: string
	readonly userId: string
	readonly providerId: string
}

export interface PersonsPart {
	readonly document: XmlDocument
	readonly prefix: string
	readonly persons: ReadonlyMap<string, Person>
}

function prefixOf(document: XmlDocument, node: number): string {
	const name = document.name(node)
	const colon = name.indexOf(":")

	return colon < 0 ? "" : name.slice(0, colon + 1)
}

function child(document: XmlDocument, node: number, local: string): number {
	for (let found = document.firstChildElement(node); found >= 0; found = document.nextSiblingElement(found))
		if (document.localName(found) === local) return found

	return -1
}

export function readThreaded(document: XmlDocument): ThreadedPart {
	const root = document.documentElement()

	if (document.localName(root) !== "ThreadedComments") throw new Error("not a threaded comments part")

	const roots = new Map<string, { comment: ThreadComment; key: number; row: number; col: number; replies: ThreadComment[] }>()
	const replies: { comment: ThreadComment; key: number }[] = []
	const comments = new Map<string, { comment: ThreadComment; key: number }>()

	for (let node = document.firstChildElement(root); node >= 0; node = document.nextSiblingElement(node)) {
		if (document.localName(node) !== "threadedComment") continue

		const at = parseCellRef(document.attribute(node, "ref") ?? "")
		const id = (document.attribute(node, "id") ?? "").toUpperCase()

		if (at === undefined || at.row >= SHEET_MAX_ROWS || at.col >= SHEET_MAX_COLS || id === "") continue

		const text = child(document, node, "text")
		const parent = document.attribute(node, "parentId")
		const comment: ThreadComment = {
			id,
			parentId: parent === undefined ? null : parent.toUpperCase(),
			personId: (document.attribute(node, "personId") ?? "").toUpperCase(),
			date: document.attribute(node, "dT") ?? "",
			text: text < 0 ? "" : decodeXstring(document.textContent(text)),
			done: document.attribute(node, "done") === "1" || document.attribute(node, "done") === "true",
			mentions: child(document, node, "mentions") >= 0,
			node
		}
		const key = cellKeyOf(at.row, at.col)

		comments.set(id, { comment, key })

		if (comment.parentId === null) roots.set(id, { comment, key, row: at.row, col: at.col, replies: [] })
		else replies.push({ comment, key })
	}

	for (const reply of replies) roots.get(reply.comment.parentId ?? "")?.replies.push(reply.comment)

	const threads = new Map<number, Thread>()

	for (const thread of roots.values())
		if (!threads.has(thread.key))
			threads.set(thread.key, { row: thread.row, col: thread.col, root: thread.comment, replies: thread.replies })

	return { document, prefix: prefixOf(document, root), threads, comments }
}

export function readPersons(document: XmlDocument): PersonsPart {
	const root = document.documentElement()

	if (document.localName(root) !== "personList") throw new Error("not a persons part")

	const persons = new Map<string, Person>()

	for (let node = document.firstChildElement(root); node >= 0; node = document.nextSiblingElement(node)) {
		if (document.localName(node) !== "person") continue

		const id = (document.attribute(node, "id") ?? "").toUpperCase()

		if (id === "" || persons.has(id)) continue

		persons.set(id, {
			id,
			displayName: decodeXstring(document.attribute(node, "displayName") ?? ""),
			userId: document.attribute(node, "userId") ?? "",
			providerId: document.attribute(node, "providerId") ?? ""
		})
	}

	return { document, prefix: prefixOf(document, root), persons }
}

export function emptyThreaded(main: string): Uint8Array {
	return encodeUtf8(`${DECLARATION}<ThreadedComments xmlns="${THREADED_NAMESPACE}" xmlns:x="${main}"/>`)
}

export function emptyPersons(main: string): Uint8Array {
	return encodeUtf8(`${DECLARATION}<personList xmlns="${THREADED_NAMESPACE}" xmlns:x="${main}"/>`)
}

// The person comments are written under: one with the author's name that no account provides (providerId "None",
// as Excel writes a local author), else a new one.
export function findPerson(part: PersonsPart | null, displayName: string): Person | undefined {
	for (const person of part?.persons.values() ?? []) if (person.displayName === displayName && person.providerId === "None") return person

	return undefined
}

export function addPerson(part: PersonsPart, person: Person): Uint8Array {
	const editor = new XmlEditor(part.document)
	const p = part.prefix

	editor.append(
		part.document.documentElement(),
		`<${p}person displayName="${escapeAttribute(encodeXstring(person.displayName))}" id="${person.id}" userId="${escapeAttribute(person.userId)}" providerId="${escapeAttribute(person.providerId)}"/>`
	)

	return editor.serialize()
}

export interface CommentWrite {
	readonly row: number
	readonly col: number
	readonly id: string
	readonly parentId: string | null
	readonly personId: string
	readonly date: string
	readonly text: string
}

function textMarkup(text: string, prefix: string): string {
	const encoded = encodeXstring(text)
	// As the writer of cells and shared strings marks it: only text with leading or trailing white space.
	const preserve = /^\s|\s$/.test(encoded)

	return `<${prefix}text${preserve ? ' xml:space="preserve"' : ""}>${escapeText(encoded)}</${prefix}text>`
}

// A comment added: a new thread at the end, a reply after its thread's last comment.
export function addComment(part: ThreadedPart, write: CommentWrite): Uint8Array {
	const { document, prefix } = part
	const editor = new XmlEditor(document)
	const markup =
		`<${prefix}threadedComment ref="${cellRef(write.row, write.col)}" dT="${write.date}" personId="${write.personId}" id="${write.id}"` +
		`${write.parentId === null ? "" : ` parentId="${write.parentId}"`}>${textMarkup(write.text, prefix)}</${prefix}threadedComment>`
	let after = -1
	let last = -1

	for (const { comment } of write.parentId === null ? [] : part.comments.values()) {
		if (comment.id !== write.parentId && comment.parentId !== write.parentId) continue

		const start = document.start(comment.node)

		if (start > last) {
			last = start
			after = comment.node
		}
	}

	if (after >= 0) editor.insertAfter(after, markup)
	else editor.append(document.documentElement(), markup)

	return editor.serialize()
}

export function editComment(part: ThreadedPart, id: string, text: string): Uint8Array | null {
	const found = part.comments.get(id.toUpperCase())

	if (found === undefined) return null

	const { document, prefix } = part
	const editor = new XmlEditor(document)
	const node = child(document, found.comment.node, "text")

	if (node < 0) editor.prepend(found.comment.node, textMarkup(text, prefix))
	else editor.replace(node, textMarkup(text, prefix))

	return editor.serialize()
}

// Comments removed by id; removing a thread's first comment removes its replies too.
export function removeComments(part: ThreadedPart, ids: ReadonlySet<string>): Uint8Array {
	const editor = new XmlEditor(part.document)

	for (const { comment } of part.comments.values())
		if (ids.has(comment.id) || (comment.parentId !== null && ids.has(comment.parentId))) editor.remove(comment.node)

	return editor.serialize()
}

// What Excel writes as the legacy note of a thread, for versions without threaded comments (lo__threadedComment.xlsx,
// written by Excel): the thread's text under a fixed explanation.
export function placeholderText(texts: readonly string[]): string {
	const [first = "", ...replies] = texts

	return (
		"[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if " +
		"the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    " +
		first +
		replies.map(reply => `\nReply:\n    ${reply}`).join("")
	)
}
