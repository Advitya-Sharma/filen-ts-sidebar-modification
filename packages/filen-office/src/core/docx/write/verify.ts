import { zipSync } from "fflate"
import type { Mark, Node } from "prosemirror-model"
import {
	bytesSource,
	parseContentTypes,
	parseRelationships,
	partKey,
	relationshipsPartName,
	OpcPackage,
	type ByteSource,
	type ContentTypes,
	type Relationships,
	type ZipEntry
} from "../../opc/index"
import { NodeKind, assertWellFormed, decodeUtf8, encodeUtf8, parseXml, type XmlDocument } from "../../xml/index"
import { DocxDocBuilder, type StoryDoc } from "../edit/build"
import { openDocxEditContext, type DocxEditContext } from "../edit/context"
import { docxSchema } from "../edit/schema"
import { Item, buildDocxModel, type StoryPart } from "../model/index"
import { partReader } from "../model/reader"
import { Ns } from "../model/namespaces"
import { parseNumbering, type Numbering } from "../numbering/index"
import { openDocxPackage, type DocxPackage } from "../package"
import { PropsStore } from "../props/index"
import { parseStyles, type DocxStyles } from "../styles/index"
import { canonicalElement, canonicalFragment } from "./canonical"
import type { DocxChanges, DocxSavePlan, StoryChange } from "./save"
import { cellWithin, inCell, sectionWritten, type StoryOutput } from "./story"

// Gate 2: the written file is checked before it may be uploaded, in two halves.
// Parts (before the zip is written): every changed part well-formed with declared prefixes, every new part
// typed, every relationship id new markup uses defined, no id space gaining a duplicate, no style or list
// reference newly dangling, list definitions with unique ids, container copies keeping their properties,
// and the rebuilt paragraphs reading back as the editor holds them. A rewritten story is checked by what is
// new in it: copies of whole source nodes stand in as placeholders, their facts carried over from the
// source, so a one-word edit in a 1,000-page document parses a few kilobytes. The copies are not taken on
// the writer's word: each must be byte for byte the whole source nodes it claims, and the story's
// paragraphs, copied and rewritten, must be the editor's, one for one and in order.
// Zip (after): every untouched entry the original compressed bytes, every changed entry the CRC of its
// planned bytes, nothing more and nothing less.

export type VerifyFailure =
	| "invalidPackage"
	| "untouchedChanged"
	| "malformedPart"
	| "contentType"
	| "relationship"
	| "duplicateId"
	| "danglingReference"
	| "model"
	| "content"

export class DocxVerifyError extends Error {
	readonly reason: VerifyFailure

	constructor(reason: VerifyFailure, message: string) {
		super(message)
		this.name = "DocxVerifyError"
		this.reason = reason
	}
}

const CHUNK = 4 * 1024 * 1024
const W_URIS = ["http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main"]
const R_URIS = [
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	"http://purl.oclc.org/ooxml/officeDocument/relationships"
]
// Elements whose w:id must be unique, per kind.
const REVISION_ELEMENTS = new Set([
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
	"cellMerge",
	"moveFromRangeStart",
	"moveToRangeStart"
])
const PLACEHOLDER = encodeUtf8("<_/>")
const STYLES_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
const NUMBERING_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"

function fail(reason: VerifyFailure, message: string): never {
	throw new DocxVerifyError(reason, message)
}

function describe(error: unknown): string {
	return error instanceof Error ? error.message : String(error)
}

// ---- facts ------------------------------------------------------------------------------------------------

interface Facts {
	// Ids per id space, counted: paragraph ids, bookmarks, comment anchors, revisions, drawing ids.
	readonly ids: Map<string, number>
	// Style and list references, counted.
	readonly styles: Map<string, number>
	readonly lists: Map<number, number>
}

function emptyFacts(): Facts {
	return { ids: new Map(), styles: new Map(), lists: new Map() }
}

function bump<K>(map: Map<K, number>, key: K, by: number): void {
	const count = (map.get(key) ?? 0) + by

	if (count === 0) map.delete(key)
	else map.set(key, count)
}

function addFacts(into: Facts, from: Facts, sign: number): void {
	for (const [key, count] of from.ids) bump(into.ids, key, sign * count)
	for (const [key, count] of from.styles) bump(into.styles, key, sign * count)
	for (const [key, count] of from.lists) bump(into.lists, key, sign * count)
}

// Facts of the elements starting in nodes [first, end) of a part.
function collectFacts(document: XmlDocument, first: number, end: number, into: Facts, sign: number = 1): void {
	const reader = partReader(document)

	for (let node = first; node < end; node++) {
		if (document.kinds[node] !== NodeKind.Element) continue

		const group = reader.group(node)

		if (group === Ns.WP) {
			if (document.localName(node) === "docPr") {
				const id = reader.attr(node, Ns.None, "id")

				if (id !== undefined) bump(into.ids, `docPr:${id}`, sign)
			}

			continue
		}

		if (group !== Ns.W) continue

		const local = document.localName(node)

		switch (local) {
			case "p":
			case "tr": {
				const paraId = reader.attr(node, Ns.W14, "paraId")

				if (paraId !== undefined) bump(into.ids, `paraId:${paraId.toUpperCase()}`, sign)

				break
			}
			case "bookmarkStart":
			case "commentRangeStart":
			case "commentReference": {
				const id = reader.attr(node, Ns.W, "id")

				if (id !== undefined) bump(into.ids, `${local}:${id}`, sign)

				break
			}
			case "pStyle":
			case "rStyle":
			case "tblStyle": {
				const value = reader.attr(node, Ns.W, "val")

				if (value !== undefined) bump(into.styles, value, sign)

				break
			}
			case "numId": {
				const value = Number(reader.attr(node, Ns.W, "val"))

				if (Number.isInteger(value) && value > 0) bump(into.lists, value, sign)

				break
			}
			default:
				if (REVISION_ELEMENTS.has(local)) {
					const id = reader.attr(node, Ns.W, "id")

					if (id !== undefined) bump(into.ids, `revision:${id}`, sign)
				}
		}
	}
}

// A source part's facts are read once per document (saves repeat).
const factsCache = new WeakMap<XmlDocument, Facts>()

function partFacts(document: XmlDocument): Facts {
	let facts = factsCache.get(document)

	if (facts === undefined) {
		facts = emptyFacts()
		collectFacts(document, 0, document.count, facts)
		factsCache.set(document, facts)
	}

	return facts
}

// The facts of every story part of a model, read once.
const modelCache = new WeakMap<object, Facts>()

function modelFacts(model: { readonly parts: readonly StoryPart[] }): Facts {
	let facts = modelCache.get(model)

	if (facts === undefined) {
		facts = emptyFacts()

		for (const part of model.parts) collectFacts(part.document, 0, part.document.count, facts)

		modelCache.set(model, facts)
	}

	return facts
}

// The first node starting at or after byte `offset`.
function nodeAt(document: XmlDocument, offset: number): number {
	const starts = document.starts
	let low = 0
	let high = document.count

	while (low < high) {
		const middle = (low + high) >> 1

		if ((starts[middle] ?? 0) < offset) low = middle + 1
		else high = middle
	}

	return low
}

function rangeFacts(document: XmlDocument, from: number, to: number, into: Facts, sign: number): void {
	collectFacts(document, nodeAt(document, from), nodeAt(document, to), into, sign)
}

// Takes out of `into` the facts of the source container [from, to) that its output does not copy: the gaps
// between the copied source ranges, plus once more each range copied again.
function replacedFacts(document: XmlDocument, from: number, to: number, trusted: readonly number[], into: Facts): void {
	const ranges: [number, number][] = []

	for (let k = 0; k < trusted.length; k += 4) ranges.push([trusted[k + 2] ?? 0, trusted[k + 3] ?? 0])

	ranges.sort((a, b) => a[0] - b[0] || b[1] - a[1])

	let at = from

	for (const [start, end] of ranges) {
		if (start < at) {
			// Copied before: it counts once more.
			rangeFacts(document, start, end, into, 1)

			continue
		}

		rangeFacts(document, at, start, into, -1)
		at = end
	}

	rangeFacts(document, at, to, into, -1)
}

// ---- relationships ------------------------------------------------------------------------------------------

// Relationship ids `document` uses in attributes of the relationships namespace that `relationships` does
// not define. Start tags are searched for the bound prefixes first; only the few that hold one have their
// attributes read.
function danglingRelationships(relationships: Relationships, document: XmlDocument): string[] {
	const reader = partReader(document)
	const prefixes: Uint8Array[] = []

	for (const bindings of reader.namespaces.declarations().values()) {
		for (const [prefix, uri] of bindings) {
			if (prefix !== "" && R_URIS.includes(uri)) prefixes.push(encodeUtf8(`${prefix}:`))
		}
	}

	if (prefixes.length === 0) return []

	const bytes = document.bytes
	const out: string[] = []

	for (let node = 0; node < document.count; node++) {
		if (document.kinds[node] !== NodeKind.Element) continue

		const start = document.start(node)
		const end = document.openEnd(node)
		let found = false

		for (let i = start + 1; i < end && !found; i++) {
			if (bytes[i] !== 32 && bytes[i] !== 9 && bytes[i] !== 10 && bytes[i] !== 13) continue

			for (const prefix of prefixes) {
				let match = true

				for (let k = 0; k < prefix.length && match; k++) match = bytes[i + 1 + k] === prefix[k]

				if (match) found = true
			}
		}

		if (!found) continue

		for (const attribute of document.attributes(node)) {
			const colon = attribute.name.indexOf(":")
			const prefix = colon < 0 ? "" : attribute.name.slice(0, colon)

			if (
				prefix !== "" &&
				R_URIS.includes(reader.namespaces.uri(node, prefix) ?? "") &&
				relationships.byId(attribute.value) === undefined
			) {
				out.push(attribute.value)
			}
		}
	}

	return out
}

// ---- parts --------------------------------------------------------------------------------------------------

// Every namespace declared on the root: copies keep their meaning wherever they go in the part.
function flat(document: XmlDocument): boolean {
	const root = document.documentElement()

	for (const node of partReader(document).namespaces.declarations().keys()) if (node !== root) return false

	return true
}

function rootTags(document: XmlDocument): { open: Uint8Array; close: Uint8Array } {
	const root = document.documentElement()

	return { open: document.bytes.subarray(document.start(root), document.openEnd(root)), close: encodeUtf8(`</${document.name(root)}>`) }
}

function join(chunks: readonly Uint8Array[]): Uint8Array {
	let size = 0

	for (const chunk of chunks) size += chunk.length

	const out = new Uint8Array(size)
	let at = 0

	for (const chunk of chunks) {
		out.set(chunk, at)
		at += chunk.length
	}

	return out
}

// The new markup of rewritten stories, each source copy replaced by a placeholder, inside the part's root.
function skeleton(document: XmlDocument, outputs: readonly StoryOutput[]): Uint8Array {
	const { open, close } = rootTags(document)
	const chunks: Uint8Array[] = [open]

	for (const output of outputs) {
		let at = 0

		for (let k = 0; k < output.trusted.length; k += 4) {
			const from = output.trusted[k] ?? 0

			chunks.push(output.bytes.subarray(at, from), PLACEHOLDER)
			at = output.trusted[k + 1] ?? from
		}

		chunks.push(output.bytes.subarray(at))
	}

	chunks.push(close)

	return join(chunks)
}

function parseChecked(name: string, bytes: Uint8Array): XmlDocument {
	try {
		const document = parseXml(bytes)

		assertWellFormed(document)

		return document
	} catch (error) {
		return fail("malformedPart", `${name}: ${describe(error)}`)
	}
}

export interface VerifyPartsInput {
	readonly original: DocxPackage
	readonly context: DocxEditContext
	readonly changes: DocxChanges
}

export async function verifyDocxParts(input: VerifyPartsInput): Promise<void> {
	const { original, context, changes } = input
	const model = context.model
	const changed = [...changes.replace, ...changes.add]
	const byKey = new Map(changed.map(change => [partKey(change.name), change]))
	const typesChange = byKey.get(partKey("[Content_Types].xml"))
	const originalTypes = await original.opc.contentTypes()
	let types: ContentTypes = originalTypes

	if (typesChange !== undefined) {
		parseChecked(typesChange.name, typesChange.bytes)
		types = parseContentTypes(typesChange.bytes)

		for (const entry of original.opc.zip.entries) {
			if (originalTypes.typeOf(entry.name) !== undefined && types.typeOf(entry.name) === undefined)
				fail("contentType", `${entry.name} lost its content type`)
		}
	}

	for (const change of changes.add) {
		if (types.typeOf(change.name) === undefined) fail("contentType", `${change.name} has no content type`)
	}

	const relationshipsOf = async (part: string): Promise<Relationships> => {
		const rels = byKey.get(partKey(relationshipsPartName(part)))

		return rels === undefined ? original.relationships(part) : parseRelationships(rels.bytes, part)
	}
	const exists = (part: string): boolean => original.has(part) || byKey.has(partKey(part))
	const storyParts = new Map<string, StoryPart>(model.parts.map(part => [partKey(part.name), part]))
	const storiesByPart = new Map<string, StoryChange[]>()

	for (const story of changes.stories) {
		const key = partKey(story.part.name)
		const list = storiesByPart.get(key)

		if (list === undefined) storiesByPart.set(key, [story])
		else list.push(story)
	}

	const rewritten: Rewritten[] = []

	// Facts of every story part: the source's, adjusted below for what changed.
	const before = modelFacts(model)
	// What the changes add (positive) and take away (negative).
	const after = emptyFacts()

	let stylesDocument: XmlDocument | undefined
	let numberingDocument: XmlDocument | undefined
	let numberingName: string | undefined

	for (const change of changed) {
		const key = partKey(change.name)
		const name = change.name

		if (key === partKey("[Content_Types].xml")) continue

		if (name.endsWith(".rels")) {
			parseChecked(name, change.bytes)

			const source = name.replace(/_rels\/([^/]+)\.rels$/, "$1")
			const relationships = parseRelationships(change.bytes, source)

			for (const relationship of relationships.all) {
				if (!relationship.external && relationship.partName !== undefined && !exists(relationship.partName)) {
					fail("relationship", `${name}: ${relationship.id} points at a missing part`)
				}
			}

			continue
		}

		const part = storyParts.get(key)
		const stories = storiesByPart.get(key)
		const outputs = stories?.map(story => story.output)
		let skeletonDocument: XmlDocument | undefined

		if (part !== undefined && stories !== undefined && outputs !== undefined) {
			checkAssembly(name, part, stories, change.bytes)
			skeletonDocument = parseChecked(name, skeleton(part.document, outputs))
			checkParagraphs(name, part, stories, skeletonDocument, rewritten)
		}

		if (part !== undefined && outputs !== undefined && skeletonDocument !== undefined && flat(part.document)) {
			// The new markup of the rewritten stories; the rest of the part is the source's.
			const document = skeletonDocument
			const relationships = await relationshipsOf(name)

			addFacts(after, partFacts(document), 1)

			for (const story of changes.stories) {
				if (story.part !== part) continue

				const container = part.node(story.session.story.story.item)

				replacedFacts(part.document, part.document.start(container), part.document.end(container), story.output.trusted, after)

				checkContainers(part, story.output)
			}

			checkRelationships(name, danglingRelationships(relationships, document), await original.relationships(name), part.document)
		} else {
			const document = parseChecked(name, change.bytes)

			if (part !== undefined) {
				addFacts(after, partFacts(part.document), -1)
				addFacts(after, partFacts(document), 1)
				checkRelationships(
					name,
					danglingRelationships(await relationshipsOf(name), document),
					await original.relationships(name),
					part.document
				)
			}

			const type = types.typeOf(name)

			if (type === STYLES_TYPE) stylesDocument = document
			if (type === NUMBERING_TYPE) {
				numberingDocument = document
				numberingName = name
			}
		}
	}

	for (const [key, change] of after.ids) {
		const was = before.ids.get(key) ?? 0

		if (change > 0 && was + change > 1) fail("duplicateId", `${key} appears ${was + change} times`)
	}

	const store = new PropsStore()
	const styles: DocxStyles = stylesDocument === undefined ? context.styles : parseStyles(stylesDocument, store)
	const numbering: Numbering | undefined =
		numberingDocument === undefined ? context.numbering : parseNumbering(numberingDocument, store, styles)

	if (numberingDocument !== undefined) {
		const known =
			numberingName !== undefined && original.has(numberingName)
				? numberingDuplicates(await original.xml(numberingName))
				: new Set<string>()

		for (const id of numberingDuplicates(numberingDocument)) if (!known.has(id)) fail("duplicateId", `numbering: ${id} appears twice`)
	}

	for (const [style, change] of after.styles) {
		if (change > 0 && styles.byId(style) === undefined && !before.styles.has(style))
			fail("danglingReference", `Style ${style} is not defined`)
	}

	for (const [list, change] of after.lists) {
		if (change > 0 && numbering?.has(list) !== true && !before.lists.has(list)) fail("danglingReference", `List ${list} is not defined`)
	}

	await readBack(input, rewritten, relationshipsOf)
}

// Relationship ids the new markup uses must exist, unless the source used them without defining them too.
function checkRelationships(name: string, dangling: readonly string[], sourceRelationships: Relationships, source: XmlDocument): void {
	if (dangling.length === 0) return

	const before = new Set(danglingRelationships(sourceRelationships, source))

	for (const id of dangling) if (!before.has(id)) fail("relationship", `${name}: relationship ${id} does not exist`)
}

// w:num and w:abstractNum ids that appear more than once in a numbering part.
function numberingDuplicates(document: XmlDocument): Set<string> {
	const reader = partReader(document)
	const root = document.documentElement()
	const seen = new Set<string>()
	const duplicates = new Set<string>()

	for (let child = document.firstChildElement(root); child >= 0; child = document.nextSiblingElement(child)) {
		if (reader.group(child) !== Ns.W) continue

		const local = document.localName(child)
		const id =
			local === "num"
				? reader.attr(child, Ns.W, "numId")
				: local === "abstractNum"
					? reader.attr(child, Ns.W, "abstractNumId")
					: undefined

		if (id === undefined) continue

		const key = `${local} ${id}`

		if (seen.has(key)) duplicates.add(key)

		seen.add(key)
	}

	return duplicates
}

// A container written as a copy keeps its source's properties, without the children carrying ids.
function checkContainers(part: StoryPart, output: StoryOutput): void {
	const { open, close } = rootTags(part.document)

	for (const container of output.containers) {
		const expected = container.properties.map(node => canonicalElement(part.document, node, false, new Set(container.strip))).join("\n")
		const document = parseChecked(part.name, join([open, output.bytes.subarray(container.from, container.to), close]))
		const actual = canonicalFragment(document, document.documentElement())

		if (actual !== expected) fail("content", `A copied container of ${part.name} lost its properties`)
	}
}

// The rebuilt paragraphs, read back by the model in a document of their own, must be what the editor holds.
// The rewritten paragraphs of one part, as written (the skeleton's markup) and as the editor holds them.
interface Rewritten {
	readonly part: StoryPart
	readonly markup: Uint8Array[]
	readonly nodes: Node[]
	// Per node, whether it is written with its section break (see sectionWritten).
	readonly sections: boolean[]
}

function sameBytes(a: Uint8Array, aFrom: number, b: Uint8Array, bFrom: number, length: number): boolean {
	for (let i = 0; i < length; i++) if (a[aFrom + i] !== b[bFrom + i]) return false

	return true
}

// Source bytes [from, to) are whole sibling nodes.
function wholeNodes(document: XmlDocument, from: number, to: number): boolean {
	let node = nodeAt(document, from)

	if (node >= document.count || document.start(node) !== from) return false

	for (;;) {
		const end = document.end(node)

		if (end === to) return true
		if (end > to) return false

		node = document.nextSibling(node)

		if (node < 0 || document.start(node) !== end) return false
	}
}

// The written part is the source outside the rewritten containers, then each container as written, and
// every copy in it is the source's bytes of the whole nodes it claims, inside that container.
function checkAssembly(name: string, part: StoryPart, stories: readonly StoryChange[], bytes: Uint8Array): void {
	const document = part.document
	const source = document.bytes
	const ordered = stories
		.map(story => ({ story, container: part.node(story.session.story.story.item) }))
		.sort((a, b) => document.start(a.container) - document.start(b.container))
	let from = 0
	let at = 0
	const same = (length: number, other: Uint8Array, otherFrom: number): void => {
		if (at + length > bytes.length || !sameBytes(bytes, at, other, otherFrom, length))
			fail("content", `${name} is not what was checked`)

		at += length
	}

	for (const { story, container } of ordered) {
		const start = document.start(container)
		const end = document.end(container)
		const output = story.output
		const trusted = output.trusted
		let written = 0

		same(start - from, source, from)

		for (let k = 0; k < trusted.length; k += 4) {
			const outFrom = trusted[k] ?? 0
			const outTo = trusted[k + 1] ?? 0
			const sourceFrom = trusted[k + 2] ?? 0
			const sourceTo = trusted[k + 3] ?? 0

			if (
				outTo - outFrom !== sourceTo - sourceFrom ||
				sourceFrom < start ||
				sourceTo > end ||
				!wholeNodes(document, sourceFrom, sourceTo)
			) {
				fail("content", `${name}: a copy is not whole source nodes of its story`)
			}

			same(outFrom - written, output.bytes, written)
			same(outTo - outFrom, source, sourceFrom)
			written = outTo
		}

		same(output.bytes.length - written, output.bytes, written)
		from = end
	}

	same(source.length - from, source, from)

	if (at !== bytes.length) fail("content", `${name} is not what was checked`)
}

// Paragraph items starting in source bytes [from, to), outside other paragraphs (text boxes are stories of
// their own).
function sourceParagraphs(part: StoryPart, from: number, to: number, out: number[]): void {
	let low = 0
	let high = part.count

	while (low < high) {
		const middle = (low + high) >> 1

		if (part.start(middle) < from) low = middle + 1
		else high = middle
	}

	for (let item = low; item < part.count && part.start(item) < to;) {
		if (part.kinds[item] === Item.Paragraph) {
			out.push(item)
			item = part.end(item)
		} else {
			item++
		}
	}
}

// A paragraph the writer copied must be its source as the editor first built it (markers before it are
// written apart).
function sameAsSource(original: Node | undefined, node: Node): boolean {
	if (original === undefined) return false
	if (original.eq(node)) return true
	if (!original.content.eq(node.content) || original.marks.length !== node.marks.length) return false
	if (original.marks.some((mark, index) => !mark.eq(node.marks[index] as Mark))) return false

	return Object.keys(original.attrs).every(attr => attr === "before" || original.attrs[attr] === node.attrs[attr])
}

const originalsCache = new WeakMap<StoryDoc, Map<number, Node>>()

function originals(story: StoryDoc): Map<number, Node> {
	let map = originalsCache.get(story)

	if (map === undefined) {
		const found = new Map<number, Node>()

		story.doc.descendants(node => {
			if (node.type === nodes.paragraph) {
				const own = node.attrs["src"] as number

				if (!found.has(own)) found.set(own, node)

				return false
			}

			return !node.isInline
		})
		map = found
		originalsCache.set(story, map)
	}

	return map
}

// Every story's paragraphs, in order: the editor's (a locked block stands for the source paragraphs in its
// bytes) against the written story's (each copy for the source paragraphs in it, each new w:p as rewritten).
// A copied paragraph must be the editor's paragraph there, unchanged; a rewritten one is read back below.
function checkParagraphs(name: string, part: StoryPart, stories: readonly StoryChange[], written: XmlDocument, out: Rewritten[]): void {
	const reader = partReader(written)
	const rewritten: Rewritten = { part, markup: [], nodes: [], sections: [] }
	let container = written.firstChildElement(written.documentElement())

	for (const story of stories) {
		if (container < 0) fail("content", `${name}: a rewritten story is missing`)

		const expected: { readonly node: Node | null; readonly item: number; readonly cell: boolean; readonly section: boolean }[] = []
		// Table cells holding the blocks visited, by their end position.
		const cells: number[] = []

		story.session.state.doc.descendants((node, pos) => {
			while (cells.length > 0 && (cells[cells.length - 1] ?? 0) <= pos) cells.pop()

			if (node.type === nodes.table_cell) cells.push(pos + node.nodeSize)

			if (node.type === nodes.paragraph) {
				const item = node.attrs["src"] as number

				const cell = cells.length > 0

				expected.push({ node, item, cell, section: sectionWritten(part, node.attrs["sectPr"] as number, cell) })

				return false
			}

			if (node.type === nodes.locked_block) {
				const item = node.attrs["src"] as number
				const items: number[] = []

				sourceParagraphs(part, part.start(item), part.byteEnd(item), items)

				for (const paragraph of items)
					expected.push({
						node: null,
						item: paragraph,
						cell: cells.length > 0 || cellWithin(part, paragraph, item),
						section: false
					})

				return false
			}

			return !node.isInline
		})

		const actual: { readonly element: number; readonly item: number }[] = []
		const trusted = story.output.trusted
		const end = written.end(container)
		let copy = 0
		let skipTo = -1

		for (let node = container + 1; node < written.count && written.start(node) < end; node++) {
			if (written.kinds[node] !== NodeKind.Element) continue

			const placeholder = written.name(node) === "_"

			if (written.start(node) < skipTo) {
				if (placeholder) copy++

				continue
			}

			if (placeholder) {
				const items: number[] = []

				sourceParagraphs(part, trusted[copy * 4 + 2] ?? 0, trusted[copy * 4 + 3] ?? 0, items)

				for (const item of items) actual.push({ element: -1, item })

				copy++
			} else if (reader.group(node) === Ns.W && written.localName(node) === "p") {
				actual.push({ element: node, item: -1 })
				skipTo = written.end(node)
			}
		}

		if (copy * 4 !== trusted.length) fail("content", `${name}: the copies of a story do not add up`)
		if (actual.length !== expected.length) {
			fail("content", `${name}: a story was written with ${actual.length} paragraphs, the editor has ${expected.length}`)
		}

		const sources = originals(story.session.story)

		actual.forEach((paragraph, index) => {
			const editor = expected[index]

			if (editor === undefined) return

			if (paragraph.item >= 0) {
				const same =
					editor.node === null
						? editor.item === paragraph.item
						: editor.item === paragraph.item && sameAsSource(sources.get(paragraph.item), editor.node)

				if (!same) fail("content", `${name}: a copied paragraph is not the editor's paragraph there`)

				// Its bytes keep a section break only where its source had it: in a table cell or out of one.
				if ((part.paragraphs.sectPr[part.aux[paragraph.item] ?? -1] ?? -1) >= 0 && inCell(part, paragraph.item) !== editor.cell)
					fail("content", `${name}: a copied section break moved into or out of a table cell`)
			} else {
				if (editor.node === null) fail("content", `${name}: a locked block was rewritten`)

				rewritten.markup.push(written.bytes.subarray(written.start(paragraph.element), written.end(paragraph.element)))
				rewritten.nodes.push(editor.node)
				rewritten.sections.push(editor.section)
			}
		})

		container = written.nextSiblingElement(container)
	}

	out.push(rewritten)
}

async function readBack(
	input: VerifyPartsInput,
	rewritten: readonly Rewritten[],
	relationshipsOf: (part: string) => Promise<Relationships>
): Promise<void> {
	for (const story of rewritten) {
		const paragraphs = story.nodes

		if (paragraphs.length === 0) continue

		const document = story.part.document
		const reader = partReader(document)
		const root = document.documentElement()
		const prefix = W_URIS.map(uri => reader.namespaces.prefixFor(root, uri)).find(found => found !== undefined) ?? ""
		const qualify = (local: string): string => (prefix === "" ? local : `${prefix}:${local}`)
		const open = decodeUtf8(document.bytes, document.start(root), document.openEnd(root)).replace(
			/^<[^\s>/]+/,
			`<${qualify("document")}`
		)
		const body = join([
			encodeUtf8(
				`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${open.endsWith("/>") ? open.slice(0, -2) + ">" : open}<${qualify("body")}>`
			),
			...story.markup,
			encodeUtf8(`</${qualify("body")}></${qualify("document")}>`)
		])
		const relationships = await relationshipsOf(story.part.name)
		const files: Record<string, Uint8Array> = {
			"[Content_Types].xml": encodeUtf8(
				'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
					'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
					'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
			),
			"_rels/.rels": encodeUtf8(
				'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
			),
			"word/document.xml": body,
			"word/_rels/document.xml.rels": encodeUtf8(
				'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					relationships.all
						.map(
							relationship =>
								`<Relationship Id="${escape(relationship.id)}" Type="${escape(relationship.type)}" Target="${escape(relationship.target)}"${relationship.external ? ' TargetMode="External"' : ""}/>`
						)
						.join("") +
					"</Relationships>"
			)
		}
		const opened = await openDocxPackage(bytesSource(zipSync(files, { level: 0 })))

		if (!opened.ok) fail("model", `The rewritten ${story.part.name} does not read back: ${opened.message}`)

		const model = await buildDocxModel(opened.package)

		if (model.malformed || model.body === undefined) fail("model", `The rewritten ${story.part.name} has malformed markup`)

		const context = await openDocxEditContext(opened.package, model)
		const readBlocks = textblocks(new DocxDocBuilder(context).build(model.body).doc)

		if (readBlocks.length !== paragraphs.length)
			fail("content", `${paragraphs.length} rewritten paragraphs read back as ${readBlocks.length}`)

		paragraphs.forEach((paragraph, index) => {
			const other = readBlocks[index]

			if (other === undefined || other.type !== nodes.paragraph) return

			const expected = paragraphSignature(paragraph, input.changes.numIds, story.sections[index])
			const actual = paragraphSignature(other, input.changes.numIds)

			if (expected !== actual)
				fail("content", `A rewritten paragraph of ${story.part.name} reads back differently:\n${actual}\n!=\n${expected}`)
		})
	}
}

function escape(value: string): string {
	return value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;")
}

// ---- zip ----------------------------------------------------------------------------------------------------

async function sameRaw(a: OpcPackage, entryA: ZipEntry, b: OpcPackage, entryB: ZipEntry): Promise<boolean> {
	if (
		entryA.method !== entryB.method ||
		entryA.crc32 !== entryB.crc32 ||
		entryA.compressedSize !== entryB.compressedSize ||
		entryA.uncompressedSize !== entryB.uncompressedSize
	) {
		return false
	}

	const layoutA = await a.zip.layout(entryA)
	const layoutB = await b.zip.layout(entryB)

	for (let offset = 0; offset < entryA.compressedSize; offset += CHUNK) {
		const length = Math.min(CHUNK, entryA.compressedSize - offset)
		const left = await a.zip.source.read(layoutA.dataOffset + offset, length)
		const right = await b.zip.source.read(layoutB.dataOffset + offset, length)

		for (let i = 0; i < length; i++) if (left[i] !== right[i]) return false
	}

	return true
}

// The zip half: `written` (random access to the produced file) holds every original entry as its original
// compressed bytes, every changed part with the CRC of its planned bytes, and nothing else.
export async function verifyDocxZip(original: DocxPackage, written: ByteSource, save: DocxSavePlan): Promise<void> {
	let output: OpcPackage

	try {
		output = await OpcPackage.open(written)
	} catch (error) {
		return fail("invalidPackage", `The written file does not open: ${describe(error)}`)
	}

	const entries = new Map(output.zip.entries.map(entry => [entry.name, entry]))
	const seen = new Set<string>()

	for (const entry of original.opc.zip.entries) {
		const key = partKey(entry.name)
		const copy = entries.get(entry.name)

		seen.add(entry.name)

		if (copy === undefined) fail("untouchedChanged", `${entry.name} is missing`)

		const crc = save.crcs.get(key)

		if (crc !== undefined) {
			if (copy.crc32 !== crc) fail("untouchedChanged", `${entry.name} is not what was written`)

			continue
		}

		if (!(await sameRaw(original.opc, entry, output, copy))) fail("untouchedChanged", `${entry.name} changed`)
	}

	for (const entry of output.zip.entries) {
		if (seen.has(entry.name)) continue

		const crc = save.crcs.get(partKey(entry.name))

		if (crc === undefined) fail("untouchedChanged", `Unexpected entry ${entry.name}`)
		if (entry.crc32 !== crc) fail("untouchedChanged", `${entry.name} is not what was written`)
	}

	const names = new Set(output.zip.entries.map(entry => partKey(entry.name)))

	for (const key of save.crcs.keys()) if (!names.has(key)) fail("untouchedChanged", `${key} was not written`)
}

const nodes = docxSchema.nodes
const marks = docxSchema.marks
const IGNORED_MARKS = new Set([marks.run, marks.plain])
const IDENTITY_MARKS = new Set([marks.insertion, marks.deletion, marks.move_from, marks.move_to, marks.content_control, marks.wrapper])

function sorted(value: unknown): unknown {
	if (value === null || typeof value !== "object") return value
	if (Array.isArray(value)) return value.map(sorted)

	return Object.fromEntries(
		Object.entries(value as Record<string, unknown>)
			.filter(([, entry]) => entry !== undefined)
			.sort(([a], [b]) => (a < b ? -1 : 1))
			.map(([key, entry]) => [key, sorted(entry)])
	)
}

function markSignature(mark: Mark): string {
	if (IDENTITY_MARKS.has(mark.type)) return mark.type.name
	// A tracked formatting change split over several runs gets new ids: who and when is what counts.
	if (mark.type === marks.format_change) return `format_change:${String(mark.attrs["author"])}:${String(mark.attrs["date"])}`
	if (mark.type === marks.link) return `link:${String(mark.attrs["anchor"] || null)}:${String(mark.attrs["target"] || null)}`

	return `${mark.type.name}${JSON.stringify(sorted(mark.attrs))}`
}

// The paragraphs and locked blocks of a document, in order.
export function textblocks(doc: Node): Node[] {
	const out: Node[] = []

	doc.descendants(node => {
		if (node.type === nodes.paragraph || node.type === nodes.locked_block) {
			out.push(node)

			return false
		}

		return !node.isInline
	})

	return out
}

// What a reader sees and edits of a paragraph, without source items: properties, then text and atoms with
// their marks (text with the same marks joined).
// Text as XML can carry it: a lone surrogate (half a pair an undo can bring back) is written as U+FFFD.
function xmlText(text: string): string {
	let out: string | undefined

	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i)

		if (code < 0xd800 || code > 0xdfff) continue

		const next = text.charCodeAt(i + 1)

		if (code <= 0xdbff && next >= 0xdc00 && next <= 0xdfff) {
			i++

			continue
		}

		out = (out ?? text).slice(0, i) + "\ufffd" + (out ?? text).slice(i + 1)
	}

	return out ?? text
}

// `section`: whether the paragraph is a section break as written (its own w:sectPr by default).
export function paragraphSignature(
	node: Node,
	numIds: ReadonlyMap<number, number>,
	section = (node.attrs["sectPr"] as number) >= 0
): string {
	const numId = node.attrs["numId"] as number | null
	const parts: string[] = [
		JSON.stringify([
			node.attrs["styleId"],
			numId !== null && numId < 0 ? (numIds.get(numId) ?? numId) : numId,
			node.attrs["ilvl"],
			node.attrs["align"],
			sorted(node.attrs["indent"]),
			section
		])
	]
	let lastText = -1

	node.forEach(child => {
		if (child.type === nodes.proof_err || child.type === nodes.rendered_break) return

		const markPart = child.marks
			.filter(mark => !IGNORED_MARKS.has(mark.type))
			.map(markSignature)
			.join(",")

		if (child.isText) {
			const previous = parts[lastText]
			const text = xmlText(child.text ?? "")

			if (lastText === parts.length - 1 && previous !== undefined && previous.startsWith(`t[${markPart}]`)) {
				parts[lastText] = previous + text
			} else {
				parts.push(`t[${markPart}]${text}`)
				lastText = parts.length - 1
			}

			return
		}

		const attributes = { ...child.attrs }

		for (const name of ["src", "end", "node", "markers", "index", "simple"]) delete attributes[name]

		parts.push(`${child.type.name}${JSON.stringify(sorted(attributes))}[${markPart}]`)
	})

	return parts.join("\n")
}
