import { OpcError, Relationships, ZipError, partKey } from "../../opc/index"
import { XmlError, type XmlDocument } from "../../xml/index"
import type { DocxPackage } from "../package"
import { ModelTables, buildPart, type BuildCounts, type MaxIds } from "./build"
import {
	readCommentsExtended,
	readCommentsExtensible,
	readCommentsIds,
	readPeople,
	type CommentExtension,
	type CommentInfo,
	type Person
} from "./comments"
import type { DrawingInfo, VmlInfo } from "./drawings"
import { Item } from "./kinds"
import { El, Ns, STRICT_WORDPROCESSINGML } from "./namespaces"
import type { StoryPart } from "./part"
import { partReader, type PartReader } from "./reader"
import type { DocxSettings } from "./settings"
import type { Story } from "./story"

// Story indices of a section's headers or footers by w:type; -1 when there is none.
export interface HeaderFooterSet {
	readonly default: number
	readonly first: number
	readonly even: number
}

const NO_HEADERS: HeaderFooterSet = { default: -1, first: -1, even: -1 }

// A body section. Its blocks are the body items [start, end); `item` is the paragraph carrying the
// section's sectPr, the body's final SectionProperties item, or -1 when the body ends without one.
export interface Section {
	readonly index: number
	readonly item: number
	readonly sectPr: number
	readonly start: number
	readonly end: number
	// Wrappers (block-level sdt, customXml, AlternateContent, revision wrappers) that are open at `start`
	// because the previous section ended inside them, outermost first; layout restores their context.
	readonly openWrappers: readonly number[]
	readonly headers: HeaderFooterSet
	readonly footers: HeaderFooterSet
	// With references a section leaves out inherited from the previous section (ECMA-376 §17.10.5).
	readonly effectiveHeaders: HeaderFooterSet
	readonly effectiveFooters: HeaderFooterSet
}

export interface CommentAnchorTable {
	readonly count: number
	readonly ids: ReadonlyMap<string, number>
	readonly startPart: Int32Array
	readonly startItem: Int32Array
	readonly endPart: Int32Array
	readonly endItem: Int32Array
	readonly referencePart: Int32Array
	readonly referenceItem: Int32Array
}

export interface MediaRef {
	readonly relationshipId: string
	// The target as written; external targets are never fetched.
	readonly target: string
	readonly external: boolean
	readonly partName: string | undefined
	readonly contentType: string | undefined
	// Uncompressed size from the zip directory.
	readonly size: number
}

export interface ModelInit {
	readonly package: DocxPackage
	readonly parts: readonly StoryPart[]
	readonly tables: ModelTables
	readonly sections: readonly Section[]
	readonly comments: readonly CommentInfo[]
	readonly people: readonly Person[]
	readonly background: number
	readonly strict: boolean
	readonly missingParts: readonly string[]
	readonly malformedParts: readonly string[]
	readonly rejectedParts: readonly string[]
}

// The read model of one document: story parts as item streams, stories, sections, notes and comments.
// Nothing is decoded that layout has not asked for; media stay in the zip until requested.
export class DocxModel {
	readonly package: DocxPackage
	readonly parts: readonly StoryPart[]
	readonly stories: readonly Story[]
	readonly body: Story | undefined
	readonly sections: readonly Section[]
	readonly footnotes: ReadonlyMap<string, Story>
	readonly endnotes: ReadonlyMap<string, Story>
	readonly comments: readonly CommentInfo[]
	readonly people: readonly Person[]
	readonly drawings: readonly DrawingInfo[]
	readonly vml: readonly VmlInfo[]
	readonly commentAnchors: CommentAnchorTable
	readonly strings: readonly string[]
	readonly maxIds: Readonly<MaxIds>
	readonly counts: Readonly<BuildCounts>
	readonly settings: DocxSettings
	// w:background of the main document, -1 when absent.
	readonly background: number
	readonly strict: boolean
	// Parts referenced by relationships that the package does not contain.
	readonly missingParts: readonly string[]
	// Parts (or their .rels) that could not be parsed; they are left out (a .rels as if empty).
	readonly malformedParts: readonly string[]
	// Parts referenced in a role their root element does not fit (a header that is the main document).
	readonly rejectedParts: readonly string[]

	constructor(init: ModelInit) {
		const tables = init.tables
		const anchors = tables.anchors

		this.package = init.package
		this.parts = init.parts
		this.stories = tables.stories
		this.body = tables.stories.find(story => story.kind === "body")
		this.sections = init.sections
		this.footnotes = new Map(
			tables.stories.filter(story => story.kind === "footnote" && story.id !== undefined).map(story => [story.id ?? "", story])
		)
		this.endnotes = new Map(
			tables.stories.filter(story => story.kind === "endnote" && story.id !== undefined).map(story => [story.id ?? "", story])
		)
		this.comments = init.comments
		this.people = init.people
		this.drawings = tables.drawings
		this.vml = tables.vml
		this.commentAnchors = {
			count: anchors.count,
			ids: anchors.ids,
			startPart: anchors.startPart.trim(),
			startItem: anchors.startItem.trim(),
			endPart: anchors.endPart.trim(),
			endItem: anchors.endItem.trim(),
			referencePart: anchors.referencePart.trim(),
			referenceItem: anchors.referenceItem.trim()
		}
		this.strings = tables.strings.values
		this.maxIds = tables.maxIds
		this.counts = tables.counts
		this.settings = init.package.settings
		this.background = init.background
		this.strict = init.strict
		this.missingParts = init.missingParts
		this.malformedParts = init.malformedParts
		this.rejectedParts = init.rejectedParts
	}

	// Malformed markup was found and kept opaque: such a document is view-only.
	get malformed(): boolean {
		return this.counts.malformed > 0 || this.malformedParts.length > 0
	}

	partOf(story: Story): StoryPart {
		const part = this.parts[story.part]

		if (part === undefined) throw new RangeError(`No part ${story.part}`)

		return part
	}

	string(id: number): string | undefined {
		return id < 0 ? undefined : this.strings[id]
	}

	// The story a Body/Header/Footer/Note/Comment/Textbox item roots.
	storyOf(part: StoryPart, item: number): Story | undefined {
		const kind = part.kinds[item]

		if (kind === undefined || kind < Item.Body || kind > Item.Textbox) return undefined

		return this.stories[part.aux[item] ?? -1]
	}

	mediaRef(part: number, relationshipId: string): MediaRef | undefined {
		const relationship = this.parts[part]?.relationships.byId(relationshipId)

		if (relationship === undefined) return undefined

		if (relationship.external || relationship.partName === undefined) {
			return { relationshipId, target: relationship.target, external: true, partName: undefined, contentType: undefined, size: 0 }
		}

		const entry = this.package.opc.entry(relationship.partName)

		return {
			relationshipId,
			target: relationship.target,
			external: false,
			partName: entry?.name ?? relationship.partName,
			contentType: this.package.contentTypes.typeOf(relationship.partName),
			size: entry?.uncompressedSize ?? 0
		}
	}

	// The bytes of an internal media part, inflated on each call (callers cache); undefined for external
	// or missing targets, which are never fetched.
	async media(ref: MediaRef): Promise<Uint8Array | undefined> {
		if (ref.external || ref.partName === undefined || !this.package.has(ref.partName)) return undefined

		return this.package.opc.read(ref.partName)
	}

	chartPart(drawing: DrawingInfo): string | undefined {
		return drawing.chartId === undefined ? undefined : this.mediaRef(drawing.part, drawing.chartId)?.partName
	}

	// The cached drawing Word stores for a SmartArt diagram (dsp:drawing), found through the data part's
	// dsp:dataModelExt@relId, which names a relationship of the part holding the drawing.
	async smartArtDrawingPart(drawing: DrawingInfo): Promise<string | undefined> {
		const dataId = drawing.diagram?.data
		const data = dataId === undefined ? undefined : this.mediaRef(drawing.part, dataId)

		if (data?.partName === undefined) return undefined

		let document: XmlDocument

		try {
			document = await this.package.xml(data.partName)
		} catch (error) {
			if (isDocumentError(error)) return undefined

			throw error
		}

		const reader = partReader(document)

		for (let node = 0; node < document.count; node++) {
			if (document.localName(node) !== "dataModelExt") continue

			const relId = reader.attr(node, Ns.None, "relId")

			if (relId !== undefined) return this.mediaRef(drawing.part, relId)?.partName
		}

		return undefined
	}
}

interface LoadedPart {
	readonly name: string
	readonly document: XmlDocument
	readonly relationships: Relationships
}

function headerFooterReferences(reader: PartReader, sectPr: number, relationships: Relationships, header: boolean): Map<string, string> {
	const document = reader.document
	const out = new Map<string, string>()
	const code = header ? El.HeaderReference : El.FooterReference

	for (let child = document.firstChildElement(sectPr); child >= 0; child = document.nextSiblingElement(child)) {
		if (reader.code(child) !== code) continue

		const type = reader.attr(child, Ns.W, "type") ?? "default"
		const id = reader.attr(child, Ns.R, "id")
		const target = id === undefined ? undefined : relationships.byId(id)

		if (target !== undefined && !target.external && target.partName !== undefined && !out.has(type)) out.set(type, target.partName)
	}

	return out
}

function lastParagraphId(part: StoryPart, story: Story): string | undefined {
	let last = -1

	for (let item = story.item + 1; item < part.end(story.item); item = part.end(item)) {
		if (part.kinds[item] === Item.Paragraph) last = item
	}

	return last < 0 ? undefined : part.paragraphId(last)?.toUpperCase()
}

function isDocumentError(error: unknown): boolean {
	return error instanceof XmlError || error instanceof ZipError || error instanceof OpcError
}

// The root element each referenced story part must have.
type Role = "header" | "footer" | "footnotes" | "endnotes" | "comments"

const ROOTS: Readonly<Record<Role, El>> = {
	header: El.Header,
	footer: El.Footer,
	footnotes: El.Footnotes,
	endnotes: El.Endnotes,
	comments: El.Comments
}

function openWrappers(part: StoryPart, start: number, end: number, body: number): number[] {
	const out: number[] = []

	if (start >= end) return out

	for (let ancestor = part.parent(start); ancestor > body; ancestor = part.parent(ancestor)) out.push(ancestor)

	return out.reverse()
}

// Reads every story part the main document needs: the body, the headers and footers its sections refer
// to, notes and comments. Parts are fetched in parallel and built one after the other. Malformed markup
// never throws: unreadable parts are listed in malformedParts, unreadable elements become opaque items.
export async function buildDocxModel(pkg: DocxPackage): Promise<DocxModel> {
	const tables = new ModelTables()
	const parts: StoryPart[] = []
	const indexByKey = new Map<string, number>()
	const missingParts: string[] = []
	const malformedParts: string[] = []
	const rejectedParts: string[] = []
	const load = async (name: string): Promise<LoadedPart | undefined> => {
		if (!pkg.has(name)) {
			missingParts.push(name)

			return undefined
		}

		let document: XmlDocument
		let relationships: Relationships

		try {
			document = await pkg.xml(name)
		} catch (error) {
			if (!isDocumentError(error)) throw error

			malformedParts.push(name)

			return undefined
		}

		try {
			relationships = await pkg.relationships(name)
		} catch (error) {
			if (!isDocumentError(error)) throw error

			malformedParts.push(`${name}.rels`)
			relationships = new Relationships([])
		}

		return { name, document, relationships }
	}
	const add = (loaded: LoadedPart, reader: PartReader): StoryPart => {
		const index = parts.length
		const part = buildPart(reader, tables, index, loaded.name, loaded.relationships)

		parts.push(part)
		indexByKey.set(partKey(loaded.name), index)

		return part
	}
	const optional = async (kind: string): Promise<XmlDocument | undefined> => {
		const name = pkg.part(kind)

		if (name === undefined) return undefined

		try {
			return await pkg.xml(name)
		} catch (error) {
			if (!isDocumentError(error)) throw error

			malformedParts.push(name)

			return undefined
		}
	}

	const main = await load(pkg.mainPart)
	const mainReader = main === undefined ? undefined : partReader(main.document)
	const body = main === undefined || mainReader === undefined ? undefined : add(main, mainReader)
	const seeds = tables.sections.slice()
	const references = seeds.map(seed => ({
		headers:
			main === undefined || mainReader === undefined
				? new Map<string, string>()
				: headerFooterReferences(mainReader, seed.node, main.relationships, true),
		footers:
			main === undefined || mainReader === undefined
				? new Map<string, string>()
				: headerFooterReferences(mainReader, seed.node, main.relationships, false)
	}))
	const wanted: { name: string; role: Role }[] = []
	const seen = new Set<string>([partKey(pkg.mainPart)])
	const want = (name: string | undefined, role: Role): void => {
		if (name === undefined) return

		const key = partKey(name)

		if (key === partKey(pkg.mainPart)) {
			rejectedParts.push(name)
		} else if (!seen.has(key)) {
			seen.add(key)
			wanted.push({ name, role })
		}
	}

	for (const reference of references) {
		for (const name of reference.headers.values()) want(name, "header")
		for (const name of reference.footers.values()) want(name, "footer")
	}

	want(pkg.part("footnotes"), "footnotes")
	want(pkg.part("endnotes"), "endnotes")
	want(pkg.part("comments"), "comments")

	const [loaded, extended, ids, extensible, people] = await Promise.all([
		Promise.all(wanted.map(entry => load(entry.name))),
		optional("commentsExtended"),
		optional("commentsIds"),
		optional("commentsExtensible"),
		optional("people")
	])

	loaded.forEach((part, k) => {
		const role = wanted[k]?.role

		if (part === undefined || role === undefined) return

		const reader = partReader(part.document)

		if (reader.code(part.document.documentElement()) !== ROOTS[role]) {
			rejectedParts.push(part.name)

			return
		}

		add(part, reader)
	})

	const storyOfPart = (name: string | undefined, kind: Item): number => {
		if (name === undefined) return -1

		const index = indexByKey.get(partKey(name))
		const part = index === undefined ? undefined : parts[index]

		if (part === undefined || part.count === 0 || part.kinds[0] !== kind) return -1

		return part.aux[0] ?? -1
	}
	const set = (map: Map<string, string>, kind: Item): HeaderFooterSet => ({
		default: storyOfPart(map.get("default"), kind),
		first: storyOfPart(map.get("first"), kind),
		even: storyOfPart(map.get("even"), kind)
	})
	const inherit = (own: HeaderFooterSet, previous: HeaderFooterSet): HeaderFooterSet => ({
		default: own.default >= 0 ? own.default : previous.default,
		first: own.first >= 0 ? own.first : previous.first,
		even: own.even >= 0 ? own.even : previous.even
	})
	const sections: Section[] = []
	const bodyStory = tables.stories.find(story => story.kind === "body")

	if (bodyStory !== undefined && body !== undefined) {
		const bodyEnd = body.end(bodyStory.item)
		let start = bodyStory.item + 1
		let previousHeaders = NO_HEADERS
		let previousFooters = NO_HEADERS

		seeds.forEach((seed, index) => {
			const reference = references[index]
			const headers = reference === undefined ? NO_HEADERS : set(reference.headers, Item.Header)
			const footers = reference === undefined ? NO_HEADERS : set(reference.footers, Item.Footer)
			const end = body.end(seed.item)
			const effectiveHeaders = inherit(headers, previousHeaders)
			const effectiveFooters = inherit(footers, previousFooters)

			sections.push({
				index,
				item: seed.item,
				sectPr: seed.node,
				start,
				end,
				openWrappers: openWrappers(body, start, end, bodyStory.item),
				headers,
				footers,
				effectiveHeaders,
				effectiveFooters
			})
			start = end
			previousHeaders = effectiveHeaders
			previousFooters = effectiveFooters
		})

		if (start < bodyEnd || sections.length === 0) {
			sections.push({
				index: sections.length,
				item: -1,
				sectPr: -1,
				start,
				end: bodyEnd,
				openWrappers: openWrappers(body, start, bodyEnd, bodyStory.item),
				headers: NO_HEADERS,
				footers: NO_HEADERS,
				effectiveHeaders: previousHeaders,
				effectiveFooters: previousFooters
			})
		}
	}

	const extensions = extended === undefined ? new Map<string, CommentExtension>() : readCommentsExtended(extended)
	const durableIds = ids === undefined ? new Map<string, string>() : readCommentsIds(ids)
	const utcDates = extensible === undefined ? new Map<string, string>() : readCommentsExtensible(extensible)
	const commentStories = tables.stories.filter(story => story.kind === "comment")
	const paraIds = commentStories.map(story => {
		const part = parts[story.part]

		return part === undefined ? undefined : lastParagraphId(part, story)
	})
	const commentIdByParaId = new Map<string, string>()

	commentStories.forEach((story, k) => {
		const paraId = paraIds[k]

		if (paraId !== undefined && story.id !== undefined) commentIdByParaId.set(paraId, story.id)
	})

	const comments: CommentInfo[] = commentStories.map((story, k) => {
		const part = parts[story.part]
		const reader = part?.reader
		const node = part?.node(story.item) ?? -1
		const paraId = paraIds[k]
		const extension = paraId === undefined ? undefined : extensions.get(paraId)
		const durableId = paraId === undefined ? undefined : durableIds.get(paraId)
		const parentParaId = extension?.parentParaId

		return {
			id: story.id ?? "",
			story: story.index,
			author: reader?.attr(node, Ns.W, "author"),
			initials: reader?.attr(node, Ns.W, "initials"),
			date: reader?.attr(node, Ns.W, "date"),
			paraId,
			parentId: parentParaId === undefined ? undefined : commentIdByParaId.get(parentParaId),
			done: extension?.done ?? false,
			durableId,
			dateUtc: durableId === undefined ? undefined : utcDates.get(durableId),
			anchor: story.id === undefined ? -1 : (tables.anchors.ids.get(story.id) ?? -1)
		}
	})

	const background = (() => {
		if (main === undefined || mainReader === undefined) return -1

		const document = main.document
		const root = document.documentElement()

		for (let child = document.firstChildElement(root); child >= 0; child = document.nextSiblingElement(child)) {
			if (mainReader.code(child) === El.Background) return child
		}

		return -1
	})()

	return new DocxModel({
		package: pkg,
		parts,
		tables,
		sections,
		comments,
		people: people === undefined ? [] : readPeople(people),
		background,
		strict:
			pkg.facts.strict ||
			(main !== undefined &&
				mainReader !== undefined &&
				mainReader.namespaces.elementUri(main.document.documentElement()) === STRICT_WORDPROCESSINGML),
		missingParts,
		malformedParts,
		rejectedParts
	})
}
