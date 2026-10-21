import { SHEET_MAX_COLS, SHEET_MAX_ROWS } from "../../sheet/protocol"
import { decodeUtf8, encodeUtf8, escapeAttribute, escapeText, XmlEditor, type XmlDocument } from "../../xml/index"
import type { XlsxColor, XlsxFont } from "../model/types"
import { decodeXstring } from "../model/strings"
import { encodeXstring } from "../edit/document"
import { cellKeyOf } from "../edit/history"
import { cellRef, parseCellRef } from "../write/text"
import { DECLARATION, MC_NAMESPACE, REVISION_NAMESPACE } from "./parts"

// The comments part (ECMA-376 Part 1 §18.7): a sheet's notes, each a cell reference, an author from the part's
// list and rich text. Read into an index by cell; edited by splicing (every note and attribute not edited keeps
// its bytes). The text a note shows is its runs' text (phonetic runs aside), read as ST_Xstring.

export interface NoteEntry {
	readonly row: number
	readonly col: number
	readonly authorId: number
	readonly author: string
	readonly text: string
	// The <comment> element.
	readonly node: number
	// The thread whose placeholder this note is ([MS-XLSX] §2.1.17: Excel writes one, its author "tc={id}").
	readonly thread: string | null
}

export interface CommentsPart {
	readonly document: XmlDocument
	// The prefix the root's SpreadsheetML elements carry, with its colon ("" for the default namespace).
	readonly prefix: string
	// The prefix bound to the revision namespace on the root, for xr:uid; null when none is.
	readonly revision: string | null
	readonly authors: readonly string[]
	readonly authorsNode: number
	readonly listNode: number
	// The first note of each cell, by cellKeyOf.
	readonly notes: ReadonlyMap<number, NoteEntry>
}

const THREAD_AUTHOR = /^tc=(\{[0-9A-Fa-f-]{36}\})$/

function children(document: XmlDocument, node: number): number[] {
	const out: number[] = []

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) out.push(child)

	return out
}

function child(document: XmlDocument, node: number, local: string): number {
	for (let found = document.firstChildElement(node); found >= 0; found = document.nextSiblingElement(found))
		if (document.localName(found) === local) return found

	return -1
}

// The prefix the root binds to `namespace`, null when none is (the default namespace binds "").
export function rootPrefix(document: XmlDocument, namespace: string): string | null {
	for (const attribute of document.attributes(document.documentElement())) {
		if (attribute.value !== namespace) continue
		if (attribute.name === "xmlns") return ""
		if (attribute.name.startsWith("xmlns:")) return attribute.name.slice(6)
	}

	return null
}

// The text of a rich text element (CT_Rst): its own <t>, or its runs' (phonetic runs and properties left out).
export function richText(document: XmlDocument, node: number): string {
	let out = ""

	for (const part of children(document, node)) {
		const local = document.localName(part)

		if (local === "t") out += document.textContent(part)
		else if (local === "r") {
			const t = child(document, part, "t")

			if (t >= 0) out += document.textContent(t)
		}
	}

	return decodeXstring(out)
}

export function readComments(document: XmlDocument): CommentsPart {
	const root = document.documentElement()

	if (document.localName(root) !== "comments") throw new Error("not a comments part")

	const rootName = document.name(root)
	const colon = rootName.indexOf(":")
	const prefix = colon < 0 ? "" : rootName.slice(0, colon + 1)
	const authorsNode = child(document, root, "authors")
	const listNode = child(document, root, "commentList")
	const authors = authorsNode < 0 ? [] : children(document, authorsNode).map(author => decodeXstring(document.textContent(author)))
	const notes = new Map<number, NoteEntry>()

	for (const comment of listNode < 0 ? [] : children(document, listNode)) {
		if (document.localName(comment) !== "comment") continue

		const at = parseCellRef(document.attribute(comment, "ref") ?? "")

		if (at === undefined || at.row >= SHEET_MAX_ROWS || at.col >= SHEET_MAX_COLS) continue

		const key = cellKeyOf(at.row, at.col)

		if (notes.has(key)) continue

		const authorId = Number(document.attribute(comment, "authorId") ?? "")
		const author = Number.isInteger(authorId) ? (authors[authorId] ?? "") : ""
		const text = child(document, comment, "text")

		notes.set(key, {
			row: at.row,
			col: at.col,
			authorId: Number.isInteger(authorId) ? authorId : -1,
			author,
			text: text < 0 ? "" : richText(document, text),
			node: comment,
			thread: THREAD_AUTHOR.exec(author)?.[1]?.toUpperCase() ?? null
		})
	}

	return { document, prefix, revision: rootPrefix(document, REVISION_NAMESPACE), authors, authorsNode, listNode, notes }
}

function colorMarkup(color: XlsxColor | undefined): string {
	if (color === undefined) return ""
	if (color.theme !== undefined)
		return `<color theme="${String(color.theme)}"${color.tint === undefined ? "" : ` tint="${String(color.tint)}"`}/>`
	if (color.indexed !== undefined) return `<color indexed="${String(color.indexed)}"/>`
	if (color.argb !== undefined) return `<color rgb="${(color.argb >>> 0).toString(16).toUpperCase().padStart(8, "0")}"/>`
	if (color.auto === true) return '<color auto="1"/>'

	return ""
}

// A run's properties (CT_RPrElt) for a new note: the workbook's default font, as Excel writes a new note.
export function runProperties(font: XlsxFont | undefined, prefix: string): string {
	if (font === undefined) return ""

	const p = prefix
	let out = ""

	if (font.size !== undefined) out += `<${p}sz val="${String(font.size)}"/>`

	out += colorMarkup(font.color).replace("<color", `<${p}color`)

	if (font.name !== undefined) out += `<${p}rFont val="${escapeAttribute(font.name)}"/>`
	if (font.family !== undefined) out += `<${p}family val="${String(font.family)}"/>`
	if (font.scheme !== undefined && font.scheme !== "none") out += `<${p}scheme val="${font.scheme}"/>`

	return out === "" ? "" : `<${p}rPr>${out}</${p}rPr>`
}

function textElement(text: string, prefix: string): string {
	const encoded = encodeXstring(text)
	// As the writer of cells and shared strings marks it: only text with leading or trailing white space.
	const preserve = /^\s|\s$/.test(encoded)

	return `<${prefix}t${preserve ? ' xml:space="preserve"' : ""}>${escapeText(encoded)}</${prefix}t>`
}

function run(text: string, properties: string, prefix: string): string {
	return `<${prefix}r>${properties}${textElement(text, prefix)}</${prefix}r>`
}

// The runs of a note's new text. An edited note keeps its runs' looks: a first run the new text still starts with
// (the "Author:" line Excel writes) keeps its own, the rest takes the next run's.
function textMarkup(part: CommentsPart | null, previous: number, text: string, properties: string, prefix: string): string {
	if (part === null || previous < 0) return `<${prefix}text>${run(text, properties, prefix)}</${prefix}text>`

	const { document } = part
	const runs = children(document, previous).filter(node => document.localName(node) === "r")
	const looks = runs.map(node => {
		const rPr = child(document, node, "rPr")

		return rPr < 0 ? "" : decodeUtf8(document.raw(rPr))
	})
	const first = runs[0]

	if (first === undefined) return `<${prefix}text>${run(text, "", prefix)}</${prefix}text>`

	const lead = decodeXstring(document.textContent(child(document, first, "t")))

	if (runs.length > 1 && lead !== "" && text.startsWith(lead) && text.length > lead.length)
		return `<${prefix}text>${run(lead, looks[0] ?? "", prefix)}${run(text.slice(lead.length), looks[1] ?? "", prefix)}</${prefix}text>`

	return `<${prefix}text>${run(text, looks[0] ?? "", prefix)}</${prefix}text>`
}

export interface NoteWrite {
	readonly row: number
	readonly col: number
	readonly author: string
	readonly text: string
	// A new note's xr:uid (written where the part binds the revision namespace).
	readonly uid?: string
	// Plain text without runs (a thread's placeholder, as Excel writes it).
	readonly plain?: boolean
}

// A new comments part holding nothing yet, as Excel starts one (the revision namespace bound for uids).
export function emptyComments(main: string): Uint8Array {
	return encodeUtf8(
		`${DECLARATION}<comments xmlns="${main}" xmlns:mc="${MC_NAMESPACE}" mc:Ignorable="xr" xmlns:xr="${REVISION_NAMESPACE}"><authors/><commentList/></comments>`
	)
}

// Notes written: a cell with a note gets the new text (its author kept, as Excel keeps a note's author), the others a
// new note, placed in row then column order. `font`: the run properties of new text.
export function writeNotes(part: CommentsPart, writes: readonly NoteWrite[], font: XlsxFont | undefined): Uint8Array {
	const { document, prefix } = part
	const editor = new XmlEditor(document)
	const authors = [...part.authors]
	const newAuthors: string[] = []
	const properties = runProperties(font, prefix)
	const added: { key: number; markup: string }[] = []

	for (const write of writes) {
		const key = cellKeyOf(write.row, write.col)
		const existing = part.notes.get(key)

		if (existing !== undefined) {
			const text = child(document, existing.node, "text")
			const markup =
				write.plain === true
					? `<${prefix}text>${textElement(write.text, prefix)}</${prefix}text>`
					: textMarkup(part, text, write.text, properties, prefix)

			if (text < 0) editor.append(existing.node, markup)
			else editor.replace(text, markup)

			continue
		}

		let authorId = authors.indexOf(write.author)

		if (authorId < 0) {
			authorId = authors.length
			authors.push(write.author)
			newAuthors.push(write.author)
		}

		const uid = write.uid !== undefined && part.revision !== null ? ` ${part.revision}:uid="${write.uid}"` : ""

		added.push({
			key,
			markup:
				`<${prefix}comment ref="${cellRef(write.row, write.col)}" authorId="${String(authorId)}" shapeId="0"${uid}>` +
				`${write.plain === true ? `<${prefix}text>${textElement(write.text, prefix)}</${prefix}text>` : textMarkup(null, -1, write.text, properties, prefix)}</${prefix}comment>`
		})
	}

	const authorsMarkup = newAuthors.map(author => `<${prefix}author>${escapeText(encodeXstring(author))}</${prefix}author>`).join("")

	added.sort((a, b) => a.key - b.key)

	if (part.authorsNode >= 0) {
		if (authorsMarkup !== "") editor.append(part.authorsNode, authorsMarkup)
	} else if (authorsMarkup !== "" || part.listNode < 0) {
		// CT_Comments: authors, then commentList.
		const list = part.listNode < 0 ? `<${prefix}commentList>${added.map(entry => entry.markup).join("")}</${prefix}commentList>` : ""

		editor.prepend(document.documentElement(), `<${prefix}authors>${authorsMarkup}</${prefix}authors>${list}`)

		if (part.listNode < 0) return editor.serialize()
	}

	if (part.listNode < 0) {
		editor.insertAfter(part.authorsNode, `<${prefix}commentList>${added.map(entry => entry.markup).join("")}</${prefix}commentList>`)
	} else if (added.length > 0) {
		const existing = children(document, part.listNode).flatMap(node => {
			const at = parseCellRef(document.attribute(node, "ref") ?? "")

			return at === undefined ? [] : [{ node, key: cellKeyOf(at.row, at.col) }]
		})
		let k = 0

		for (const entry of added) {
			while (k < existing.length && (existing[k]?.key ?? 0) < entry.key) k++

			const before = existing[k]

			if (before === undefined) editor.append(part.listNode, entry.markup)
			else editor.insertBefore(before.node, entry.markup)
		}
	}

	return editor.serialize()
}

// Every note (all of a cell's, should a part hold more than one) of the cells in `keys` removed.
export function removeNotes(part: CommentsPart, keys: ReadonlySet<number>): Uint8Array {
	const { document } = part
	const editor = new XmlEditor(document)

	for (const comment of part.listNode < 0 ? [] : children(document, part.listNode)) {
		const at = parseCellRef(document.attribute(comment, "ref") ?? "")

		if (at !== undefined && keys.has(cellKeyOf(at.row, at.col))) editor.remove(comment)
	}

	return editor.serialize()
}
