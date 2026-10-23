import { PARAGRAPH_MARK_PROPERTIES, PARAGRAPH_PROPERTIES, RUN_PROPERTIES } from "../../docx/write/order.generated"
import { decodeUtf8, encodeUtf8, escapeAttribute, NodeKind, parseXml, type XmlDocument } from "../../xml/index"
import type { LossyCounts } from "./report"

// The ODT translation as a standalone .docx. The translation is the viewer's and the ODT writer's (which reads it back
// against the source), so it keeps what serves them; written out as a file it is made to stand on its own:
// - the translator's origin markers (w:rsidR = a content.xml offset on runs) are left out: internal offsets are no
//   revision session ids;
// - its flow-with-text compatibility setting (ODF 1.3 Part 3 §20.267, for this engine's layout) is left out:
//   w:compatSetting's w:name is an enumeration (§17.15.3.4) without it; documents with anchored drawings report it;
// - property children in schema order (CT_RPr, CT_PPr, CT_ParaRPr: Part 4 wml.xsd);
// - font family lists (ODF svg:font-family) as their first family: ST_String of w:rFonts is at most 31 characters;
// - revision ids unique (one id per annotation; the translator gives one change region's id to each of its
//   spans);
// - a hyperlink inside a tracked change written as a tracked change inside the hyperlink (CT_RunTrackChange holds
//   run content, not w:hyperlink: Part 4 wml.xsd EG_ContentRunContent);
// - bookmark names of at most 40 characters (ST_String maxLength 40 of w:bookmarkStart w:name), their
//   REF/PAGEREF fields and anchors renamed with them;
// - the translator's field codes that name ODF elements, variables or entries the package does not have (they
//   would not update) written as their results; bookmark names as field instructions can name them;
// - comment marks inside a comment's own text moved out (a comment holds no comment of its own);
// - each header, footer, notes and comments part with the relationships it names (the translator records them all on
//   the main document part).

const STORY = /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/u
const PROPERTY_PARTS = new Set(["word/styles.xml", "word/numbering.xml"])
const SETTINGS = "word/settings.xml"
const DOCUMENT_RELS = "word/_rels/document.xml.rels"
const REVISIONS = new Set(["ins", "del", "moveFrom", "moveTo"])
const COMMENT_MARKS = new Set(["commentRangeStart", "commentRangeEnd", "commentReference"])
const FONT_ATTRIBUTES = ["w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"]
const MAX_FONT = 31
const MAX_BOOKMARK = 40

function rank(order: readonly string[]): Map<string, number> {
	return new Map(order.map((name, k) => [name, k]))
}

const RUN_RANK = rank(RUN_PROPERTIES)
const MARK_RANK = rank(PARAGRAPH_MARK_PROPERTIES)
const PARAGRAPH_RANK = rank(PARAGRAPH_PROPERTIES)

// The first family of an ODF font family list ("'Liberation Serif', 'Times New Roman'").
function firstFamily(value: string): string {
	const first = (value.split(",")[0] ?? "").trim().replace(/^(['"])(.*)\1$/u, "$2")

	return first.slice(0, MAX_FONT)
}

interface Walk {
	readonly document: XmlDocument
	readonly part: string
	// Relationship ids the part names.
	readonly named: Set<string>
	readonly state: PackageState
}

interface PackageState {
	nextRevision: number
	readonly bookmarks: Bookmarks
	fieldsAsText: number
	commentMarksMoved: number
	// Comments (by id) whose reference the comments part held.
	readonly referenced: Set<string>
}

function attributesMarkup(document: XmlDocument, node: number, changes: ReadonlyMap<string, string>): string {
	let out = `<${document.name(node)}`

	for (const attribute of document.attributes(node))
		out += ` ${attribute.name}="${escapeAttribute(changes.get(attribute.name) ?? attribute.value)}"`

	return document.isSelfClosing(node) ? `${out}/>` : `${out}>`
}

function renamed(instruction: string, bookmarks: Bookmarks): string {
	return instruction.replace(/(\b(?:REF|PAGEREF|NOTEREF)\s+|\\l\s+"?)([^\s"\\]+)/gu, (match, head: string, name: string) => {
		const to = bookmarks.renamed.get(name)

		return to === undefined ? match : `${head}${to}`
	})
}

// The translator's field codes that are not fields a .docx can update: DOCPROPERTY naming the ODF element a field was
// read from (metadata, dates, chapters, form fields), DOCVARIABLE naming an ODF variable (the package has no
// w:docVars), CITATION naming a bibliography entry (it has no sources part), and references to a bookmark the
// document does not have. Written as their shown result (a field's result is its content).
// "reported": the translation already counts it (fieldAsText, for variables); "new": counted here; undefined: a field.
function staticField(instruction: string, bookmarks: Bookmarks): "reported" | "new" | undefined {
	const words = instruction.trim().split(/\s+/u)
	const code = (words[0] ?? "").toUpperCase()

	if (code === "DOCVARIABLE") return "reported"
	if (code === "DOCPROPERTY" || code === "CITATION") return "new"
	if ((code === "REF" || code === "NOTEREF" || code === "PAGEREF") && !bookmarks.names.has((words[1] ?? "").toLowerCase())) return "new"

	return undefined
}

// The element's markup with every fix applied, or undefined when it needs none.
function rewrite(walk: Walk, node: number): string | undefined {
	const { document, state } = walk
	const local = document.localName(node)
	const prefix = document.prefix(node)
	const changes = new Map<string, string>()

	for (const attribute of document.attributes(node)) {
		if (attribute.name.startsWith("r:")) walk.named.add(attribute.value)
	}

	if (prefix === "w") {
		if (REVISIONS.has(local) && document.attribute(node, "w:id") !== undefined) changes.set("w:id", String(state.nextRevision++))

		if (local === "rFonts") {
			for (const name of FONT_ATTRIBUTES) {
				const value = document.attribute(node, name)

				if (value !== undefined && (value.length > MAX_FONT || value.includes(","))) changes.set(name, firstFamily(value))
			}
		}

		if (local === "bookmarkStart") {
			const to = state.bookmarks.renamed.get(document.attribute(node, "w:name") ?? "")

			if (to !== undefined) changes.set("w:name", to)
		}

		if (local === "hyperlink") {
			const to = state.bookmarks.renamed.get(document.attribute(node, "w:anchor") ?? "")

			if (to !== undefined) changes.set("w:anchor", to)
		}

		if (local === "fldSimple") {
			const instruction = document.attribute(node, "w:instr") ?? ""
			const to = renamed(instruction, state.bookmarks)

			const kind = staticField(to, state.bookmarks)

			if (kind !== undefined) {
				if (kind === "new") state.fieldsAsText++

				return contentOf(walk, node)
			}

			if (to !== instruction) changes.set("w:instr", to)
		}
	}

	// Children, rewritten where they need it.
	const children: { node: number; markup: string | undefined }[] = []
	let changed = changes.size > 0

	for (let child = document.firstChild(node); child >= 0; child = document.nextSibling(child)) {
		if (document.kind(child) !== NodeKind.Element) {
			children.push({ node: child, markup: undefined })

			continue
		}

		if (
			walk.part === "word/comments.xml" &&
			prefix === "w" &&
			COMMENT_MARKS.has(document.localName(child)) &&
			document.prefix(child) === "w"
		) {
			if (document.localName(child) === "commentReference") state.referenced.add(document.attribute(child, "w:id") ?? "")

			state.commentMarksMoved++
			children.push({ node: child, markup: "" })
			changed = true

			continue
		}

		const markup = rewrite(walk, child)

		if (markup !== undefined) changed = true

		children.push({ node: child, markup })
	}

	if (prefix === "w" && local === "instrText") {
		const text = document.textContent(node)
		const to = renamed(text, state.bookmarks)

		if (to !== text) return `${attributesMarkup(document, node, changes)}${escapeTextContent(to)}</${document.name(node)}>`
	}

	// Property children in schema order.
	const order =
		prefix !== "w"
			? undefined
			: local === "rPr"
				? document.localName(document.parent(node)) === "pPr"
					? MARK_RANK
					: RUN_RANK
				: local === "pPr"
					? PARAGRAPH_RANK
					: undefined

	if (order !== undefined) {
		const last = order.size - 1
		const keyOf = (child: number): number =>
			document.kind(child) !== NodeKind.Element
				? -1
				: document.prefix(child) === "w"
					? (order.get(document.localName(child)) ?? last - 0.5)
					: last - 0.5
		const elements = children.filter(child => document.kind(child.node) === NodeKind.Element)
		// A property given twice reads as its last (as this engine's reader takes it): the earlier ones go.
		const lastOf = new Map<string, number>()

		elements.forEach((child, index) => {
			if (document.prefix(child.node) === "w") lastOf.set(document.localName(child.node), index)
		})

		const kept = elements.filter(
			(child, index) => document.prefix(child.node) !== "w" || lastOf.get(document.localName(child.node)) === index
		)
		let sorted = kept.length === elements.length

		for (let k = 1; k < kept.length && sorted; k++) {
			if (keyOf((kept[k - 1] as { node: number }).node) > keyOf((kept[k] as { node: number }).node)) sorted = false
		}

		if (!sorted) {
			const inner = kept
				.map((child, index) => ({ child, index, key: keyOf(child.node) }))
				.sort((a, b) => a.key - b.key || a.index - b.index)
				.map(({ child }) => child.markup ?? decodeUtf8(document.raw(child.node)))
				.join("")

			return `${attributesMarkup(document, node, changes).replace(/\/>$/u, ">")}${inner}</${document.name(node)}>`
		}
	}

	// A hyperlink inside a tracked change: the change goes inside the hyperlink.
	if (prefix === "w" && REVISIONS.has(local) && children.some(child => isHyperlink(document, child.node))) {
		// Each piece of the change is an annotation of its own: one id each.
		let pieces = 0
		const open = (): string =>
			attributesMarkup(document, node, pieces++ === 0 ? changes : new Map([...changes, ["w:id", String(state.nextRevision++)]]))
		const close = `</${document.name(node)}>`
		let out = ""
		let run = ""
		const flush = (): void => {
			if (run !== "") out += `${open()}${run}${close}`
			run = ""
		}

		for (const child of children) {
			const markup = child.markup ?? decodeUtf8(document.raw(child.node))

			if (!isHyperlink(document, child.node)) {
				run += markup

				continue
			}

			flush()

			const link = child.node
			const linkOpen = markup.slice(0, markup.indexOf(">") + 1)
			const inner = markup.slice(linkOpen.length, markup.length - `</${document.name(link)}>`.length)

			out += linkOpen.endsWith("/>") ? linkOpen : `${linkOpen}${open()}${inner}${close}</${document.name(link)}>`
		}

		flush()

		return out
	}

	if (!changed) return undefined

	let out = attributesMarkup(document, node, changes)

	if (document.isSelfClosing(node)) return out

	for (const child of children) out += child.markup ?? decodeUtf8(document.raw(child.node))

	return `${out}</${document.name(node)}>`
}

// An element's content, rewritten.
function contentOf(walk: Walk, node: number): string {
	const { document } = walk
	let out = ""

	for (let child = document.firstChild(node); child >= 0; child = document.nextSibling(child)) {
		out += (document.kind(child) === NodeKind.Element ? rewrite(walk, child) : undefined) ?? decodeUtf8(document.raw(child))
	}

	return out
}

function isHyperlink(document: XmlDocument, node: number): boolean {
	return document.kind(node) === NodeKind.Element && document.prefix(node) === "w" && document.localName(node) === "hyperlink"
}

function escapeTextContent(text: string): string {
	return text.replace(/&/gu, "&amp;").replace(/</gu, "&lt;").replace(/>/gu, "&gt;")
}

function rewritePart(walk: Walk): Uint8Array | undefined {
	const root = walk.document.documentElement()
	const markup = root < 0 ? undefined : rewrite(walk, root)

	if (markup === undefined) return undefined

	const bytes = walk.document.bytes

	return joinBytes([bytes.subarray(0, walk.document.start(root)), encodeUtf8(markup), bytes.subarray(walk.document.end(root))])
}

function joinBytes(pieces: readonly Uint8Array[]): Uint8Array {
	let length = 0

	for (const piece of pieces) length += piece.length

	const out = new Uint8Array(length)
	let at = 0

	for (const piece of pieces) {
		out.set(piece, at)
		at += piece.length
	}

	return out
}

interface Bookmarks {
	// Names a bookmark is written under, by its name and by the name the translator's fields refer to it by.
	readonly renamed: ReadonlyMap<string, string>
	// Every bookmark's written name, lower case (Word matches names case-insensitively).
	readonly names: ReadonlySet<string>
}

// Bookmark names as a .docx holds them: at most 40 characters, spaces, quotes and backslashes as "_"
// (as the translator's REF fields name them: a field instruction's argument holds none), unique case-insensitively.
function bookmarkNames(documents: Iterable<XmlDocument>): Bookmarks {
	const all: string[] = []

	for (const document of documents) {
		const id = document.nameIdOf("w:bookmarkStart")

		if (id < 0) continue

		for (let node = 0; node < document.count; node++)
			if (document.nameId(node) === id) all.push(document.attribute(node, "w:name") ?? "")
	}

	const kept = (name: string): boolean => name.length <= MAX_BOOKMARK && !/[\s"\\]/u.test(name)
	const names = new Set(all.filter(kept).map(name => name.toLowerCase()))
	const renamed = new Map<string, string>()

	for (const name of all) {
		if (kept(name) || renamed.has(name)) continue

		const plain = name.replace(/[\s"\\]/gu, "_")
		let to = plain.slice(0, MAX_BOOKMARK)

		for (let n = 1; names.has(to.toLowerCase()); n++) {
			const suffix = `_${String(n)}`

			to = `${plain.slice(0, MAX_BOOKMARK - suffix.length)}${suffix}`
		}

		names.add(to.toLowerCase())
		renamed.set(name, to)

		if (!renamed.has(plain)) renamed.set(plain, to)
	}

	return { renamed, names }
}

interface Relationship {
	readonly id: string
	readonly markup: string
}

function relationshipsOf(bytes: Uint8Array | undefined): Relationship[] {
	if (bytes === undefined) return []

	const document = parseXml(bytes)
	const root = document.documentElement()
	const out: Relationship[] = []

	for (let node = document.firstChildElement(root); node >= 0; node = document.nextSiblingElement(node)) {
		out.push({ id: document.attribute(node, "Id") ?? "", markup: decodeUtf8(document.raw(node)) })
	}

	return out
}

function relsName(part: string): string {
	const slash = part.lastIndexOf("/")

	return `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`
}

const FLOW_WITH_TEXT = /<w:compatSetting w:name="flow-with-text" w:uri="[^"]*" w:val="false"\/>/u
const ANCHOR = encodeUtf8("<wp:anchor")

function includes(bytes: Uint8Array, needle: Uint8Array): boolean {
	for (let at = bytes.indexOf(needle[0] ?? 0); at >= 0; at = bytes.indexOf(needle[0] ?? 0, at + 1)) {
		let k = 1

		while (k < needle.length && bytes[at + k] === needle[k]) k++

		if (k === needle.length) return true
	}

	return false
}

// The translation's parts (pictures read) made a standalone package; `withoutOrigins` already applied to stories.
export function standaloneParts(parts: Map<string, Uint8Array>, lossy: LossyCounts): void {
	const documents = new Map<string, XmlDocument>()

	for (const [name, bytes] of parts) if (STORY.test(name) || PROPERTY_PARTS.has(name)) documents.set(name, parseXml(bytes))

	const state: PackageState = {
		nextRevision: 1,
		bookmarks: bookmarkNames(documents.values()),
		fieldsAsText: 0,
		commentMarksMoved: 0,
		referenced: new Set()
	}
	const named = new Map<string, Set<string>>()

	for (const [name, document] of documents) {
		const walk: Walk = { document, part: name, named: new Set(), state }
		const written = rewritePart(walk)

		if (written !== undefined) parts.set(name, written)

		named.set(name, walk.named)
	}

	lossy.add("commentRange", state.commentMarksMoved)
	lossy.add("fieldAsText", state.fieldsAsText)

	// A comment whose reference sat in another comment: referenced where its range starts, its range collapsed there.
	if (state.referenced.size > 0) {
		const main = parts.get("word/document.xml")

		if (main !== undefined) {
			let text = decodeUtf8(main)

			for (const id of state.referenced) {
				if (text.includes(`<w:commentReference w:id="${id}"/>`)) continue

				const start = `<w:commentRangeStart w:id="${id}"/>`

				text = text.includes(start)
					? text.replace(start, `${start}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`)
					: text
			}

			parts.set("word/document.xml", encodeUtf8(text))
		}
	}

	// Relationships each story part names, from the main part's.
	const relationships = relationshipsOf(parts.get(DOCUMENT_RELS))

	for (const [name, ids] of named) {
		if (name === "word/document.xml" || !STORY.test(name) || ids.size === 0 || parts.has(relsName(name))) continue

		const own = relationships.filter(relationship => ids.has(relationship.id))

		if (own.length > 0)
			parts.set(
				relsName(name),
				encodeUtf8(
					`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${own.map(relationship => relationship.markup).join("")}</Relationships>`
				)
			)
	}

	const settings = parts.get(SETTINGS)

	if (settings !== undefined) {
		const text = decodeUtf8(settings)

		if (FLOW_WITH_TEXT.test(text)) {
			parts.set(SETTINGS, encodeUtf8(text.replace(FLOW_WITH_TEXT, "")))

			if ([...parts].some(([name, bytes]) => STORY.test(name) && includes(bytes, ANCHOR))) lossy.add("floatsOnAnchorPage")
		}
	}
}
