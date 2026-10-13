import type { Node } from "prosemirror-model"
import { TextSelection, type Transaction } from "prosemirror-state"
import { decodeUtf8, encodeUtf8, escapeAttribute, parseXml, type XmlDocument } from "../../xml/index"
import type { DocxNoteOptions } from "../edit/commands/notes"
import type { DocxPageSetup } from "../edit/commands/sections"
import type { DocxEditContext } from "../edit/context"
import { markInternal } from "../edit/invariants/index"
import { NO_MARKERS, NO_SOURCE, docxSchema } from "../edit/schema"
import { NO_HEADERS, inheritHeaders } from "../model/document"
import { Ns } from "../model/namespaces"
import { Item, RevisionFlag, partReader, type HeaderFooterSet, type Section, type Story, type StoryPart } from "../model/index"
import type { NotePositionProps, PageMargins, SectionProps } from "../props/index"
import type { DocxEditorSession } from "../edit/session"
import type { PagedEditor } from "../paged/editor"
import { resolveChanges, storyReviewIndex } from "../review/accept"
import {
	paragraphChange,
	revisionStamp,
	setParagraphChange,
	splitParagraphMark,
	withParagraphChange,
	type RevisionStamp
} from "../review/track"
import { resolveSection } from "../styles/index"
import { Names, Out, freshContainer, spliceChildren } from "../write/markup"
import { R_NAMESPACE, W_NAMESPACE, sectionWritten } from "../write/source"

// Section properties as edited. A section is ended by a paragraph carrying a section break (its `sectPr`
// attribute: the source paragraph item whose w:sectPr it writes, or a key below NO_SOURCE for a break the
// editor made) or by the body's own final properties (the doc's `sectPr`). What commands change of a section
// is a SectionEdit: on its paragraph the section_change node mark, for the final section the doc's
// `sectionChanges` (with the document-wide w:evenAndOddHeaders). Edits undo with the body; the writer splices
// them into the source's w:sectPr (or writes a new one), and layout lays out the very markup the writer makes.

export type DocxHeaderKind = "header" | "footer"
// ST_HdrFtr.
export type DocxHeaderType = "default" | "first" | "even"

export const HEADER_TYPES: readonly DocxHeaderType[] = ["default", "first", "even"]

// A section's own header (or footer) references as edited, by w:type: a story index, or null for none of
// the type (the section shows the previous section's: "link to previous"). A type left out is the source's.
export type HeaderReferenceEdits = { readonly [T in DocxHeaderType]?: number | null }

// A w:sectPrChange the editor recorded: who, when, and the section's properties before (as the edit then in
// force, null for the source's own).
export interface SectionChangeRecord extends RevisionStamp {
	readonly before: SectionEdit | null
}

// Twips throughout.
export interface SectionEdit {
	// The section: the sectPr attribute of the paragraph ending it, or the doc's for the final section.
	readonly section: number
	// The source item whose w:sectPr the properties start from (the section's own, the one a new break was
	// copied from), or NO_SOURCE for none.
	readonly base: number
	// ST_SectionMark; null: none (next page).
	readonly type?: string | null
	readonly pageSize?: { readonly width: number; readonly height: number; readonly landscape: boolean }
	readonly pageMargins?: Required<PageMargins>
	readonly columns?: { readonly count: number; readonly space: number; readonly separator: boolean }
	// `start` null: numbering continues.
	readonly pageNumbering?: { readonly format: string | null; readonly start: number | null }
	// null: none. `start` is the first number shown (written as Word writes it: one less).
	readonly lineNumbering?: {
		readonly countBy: number
		readonly start: number
		readonly distance?: number
		readonly restart?: string
	} | null
	readonly titlePage?: boolean
	readonly footnotes?: NotePositionProps | null
	readonly endnotes?: NotePositionProps | null
	readonly headers?: HeaderReferenceEdits
	readonly footers?: HeaderReferenceEdits
	// The tracked change (w:sectPrChange): undefined keeps the source's, null writes none, a record a new one.
	readonly change?: SectionChangeRecord | null
	// The source's own w:sectPrChange was rejected: the properties it recorded stand (references stay).
	readonly rejected?: boolean
}

// The doc's sectionChanges attribute.
export interface DocSectionChanges {
	readonly final?: SectionEdit
	readonly evenAndOddHeaders?: boolean
}

const nodes = docxSchema.nodes
const marks = docxSchema.marks
const NO_CHANGES: DocSectionChanges = Object.freeze({})

// CT_SectPr (wml.xsd): the header and footer references, EG_SectPrContents, then w:sectPrChange.
export const SECTION_PROPERTIES: readonly string[] = [
	"headerReference",
	"footerReference",
	"footnotePr",
	"endnotePr",
	"type",
	"pgSz",
	"pgMar",
	"paperSrc",
	"pgBorders",
	"lnNumType",
	"pgNumType",
	"cols",
	"formProt",
	"vAlign",
	"noEndnote",
	"titlePg",
	"textDirection",
	"bidi",
	"rtlGutter",
	"docGrid",
	"printerSettings",
	"sectPrChange"
]

// CT_FtnProps / CT_EdnProps.
const NOTE_PROPERTIES: readonly string[] = ["pos", "numFmt", "numStart", "numRestart"]

function isEdit(value: unknown): value is SectionEdit {
	return typeof value === "object" && value !== null && typeof (value as SectionEdit).section === "number"
}

// The edit of the section a paragraph's break ends (its section_change mark naming that break), if any.
export function sectionEditOf(node: Node): SectionEdit | undefined {
	const value: unknown = marks.section_change.isInSet(node.marks)?.attrs["value"]

	return isEdit(value) && value.section === node.attrs["sectPr"] ? value : undefined
}

export function docSectionChanges(doc: Node): DocSectionChanges {
	const value: unknown = doc.attrs["sectionChanges"]

	return typeof value === "object" && value !== null ? (value as DocSectionChanges) : NO_CHANGES
}

// The edit of the body's final section, if any.
export function finalSectionEdit(doc: Node): SectionEdit | undefined {
	const final = docSectionChanges(doc).final

	return final !== undefined && final.section === doc.attrs["sectPr"] ? final : undefined
}

// Whether the document shows even pages' headers and footers, as edited.
export function evenAndOddHeaders(context: DocxEditContext, doc: Node | undefined): boolean {
	const value = doc === undefined ? undefined : docSectionChanges(doc).evenAndOddHeaders

	return value ?? context.model.settings.evenAndOddHeaders
}

// ---- new breaks ---------------------------------------------------------------------------------------------

const LAST_KEYS = new WeakMap<DocxEditContext, number>()

// The key of a new section break: below NO_SOURCE, never handed out twice in a document.
export function newSectionKey(context: DocxEditContext): number {
	const key = (LAST_KEYS.get(context) ?? NO_SOURCE) - 1

	LAST_KEYS.set(context, key)

	return key
}

// ---- the body's sections ------------------------------------------------------------------------------------

// Where a section ends: the break's key and the edit of its properties.
export interface SectionEnd {
	readonly key: number
	readonly edit: SectionEdit | undefined
}

// A section of the body as edited: what ends it, and the paragraph doing so (`pos` -1: the body's own).
export interface BodySection extends SectionEnd {
	readonly pos: number
	readonly node: Node | undefined
}

// Whether a block (not a content control) ends a section: a paragraph's break written outside table cells.
export function endsSection(part: StoryPart, node: Node): boolean {
	const sectPr = node.attrs["sectPr"]

	return (
		(node.type === nodes.paragraph || node.type === nodes.locked_block) &&
		typeof sectPr === "number" &&
		sectionWritten(part, sectPr, false)
	)
}

export function sectionEndOf(node: Node): SectionEnd {
	return { key: node.attrs["sectPr"] as number, edit: node.type === nodes.paragraph ? sectionEditOf(node) : undefined }
}

export function finalSectionEnd(doc: Node): SectionEnd {
	return { key: doc.attrs["sectPr"] as number, edit: finalSectionEdit(doc) }
}

const BODY_SECTIONS = new WeakMap<Node, readonly BodySection[]>()

// The sections of a body document, in order: those its paragraph breaks end, and the last (the body's own
// properties, or default ones while blocks follow the last break), as layout counts them.
export function bodySections(doc: Node, part: StoryPart): readonly BodySection[] {
	const known = BODY_SECTIONS.get(doc)

	if (known !== undefined) return known

	const out: BodySection[] = []
	let lastEnds = false
	const visit = (node: Node, pos: number): void => {
		if (node.type === nodes.sdt) {
			let at = pos + 1

			node.forEach(child => {
				visit(child, at)
				at += child.nodeSize
			})

			return
		}

		lastEnds = endsSection(part, node)

		if (lastEnds) out.push({ ...sectionEndOf(node), pos, node })
	}
	let pos = 0

	doc.forEach(block => {
		visit(block, pos)
		pos += block.nodeSize
	})

	const end = doc.attrs["end"]
	const trailing = !lastEnds || (Array.isArray(end) && end.length > 0)

	if ((doc.attrs["sectPr"] as number) >= 0 || trailing || out.length === 0)
		out.push({ ...finalSectionEnd(doc), pos: -1, node: undefined })

	BODY_SECTIONS.set(doc, out)

	return out
}

// The index of the section holding body position `pos`.
export function sectionIndexAt(doc: Node, part: StoryPart, pos: number): number {
	const sections = bodySections(doc, part)

	for (let k = 0; k < sections.length; k++) {
		const section = sections[k] as BodySection

		if (section.pos < 0 || pos <= section.pos + (section.node?.nodeSize ?? 0)) return k
	}

	return Math.max(0, sections.length - 1)
}

// ---- the source -----------------------------------------------------------------------------------------------

// The w:sectPr element of source item `item` (a paragraph carrying a break, or the body's final properties),
// or -1.
export function sectionNode(part: StoryPart, item: number): number {
	if (item < 0) return -1

	const kind = part.kinds[item]

	if (kind === Item.Paragraph) return part.paragraphs.sectPr[part.aux[item] ?? -1] ?? -1
	if (kind === Item.SectionProperties) return part.node(item)

	return -1
}

// The section of the source document a key names: the opened section whose break it is.
function openedSection(opened: readonly Section[], key: number): Section | undefined {
	return opened.find(section => section.item === key)
}

// The section's own references as edited.
export function ownReferences(source: HeaderFooterSet, edits: HeaderReferenceEdits | undefined): HeaderFooterSet {
	if (edits === undefined) return source

	const pick = (type: DocxHeaderType): number => {
		const value = edits[type]

		return value === undefined ? source[type] : value === null ? -1 : value
	}

	return { default: pick("default"), first: pick("first"), even: pick("even") }
}

// ---- markup -------------------------------------------------------------------------------------------------

function markup(write: (out: Out) => void): string {
	const out = new Out()

	write(out)

	return decodeUtf8(out.finish())
}

function onOff(names: Names, local: string, on: boolean | undefined): string {
	return on === true ? names.empty(local) : ""
}

function noteProperties(names: Names, local: string, props: NotePositionProps | null): string {
	if (props === null) return ""

	const children = new Map<string, string>()

	if (props.position !== undefined) children.set("pos", names.val("pos", props.position))
	if (props.numFormat !== undefined) children.set("numFmt", names.val("numFmt", props.numFormat))
	if (props.numStart !== undefined) children.set("numStart", names.val("numStart", String(props.numStart)))
	if (props.numRestart !== undefined) children.set("numRestart", names.val("numRestart", props.numRestart))

	return freshContainer(names, local, children, NOTE_PROPERTIES)
}

// The r:id attribute name (with a declaration of its own where the part binds no prefix to it).
function relationshipAttribute(names: Names, id: string): string {
	return names.r === "" ? `xmlns:r="${R_NAMESPACE}" r:id="${escapeAttribute(id)}"` : `${names.r}:id="${escapeAttribute(id)}"`
}

// A section's header (or footer) references: those an edit names written new, the others its source's as
// they are.
function references(
	input: SectionMarkupInput,
	node: number,
	local: "headerReference" | "footerReference",
	edits: HeaderReferenceEdits | undefined
): string {
	const { names, part } = input
	const document = part.document
	let out = ""

	for (const type of ["even", "default", "first"] as const) {
		const story = edits?.[type]

		if (story === undefined) {
			for (let child = node < 0 ? -1 : document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
				if (part.reader.is(child, Ns.W, local) && (part.reader.attr(child, Ns.W, "type") ?? "default") === type)
					out += decodeUtf8(document.bytes, document.start(child), document.end(child))
			}
		} else if (story !== null && story >= 0) {
			out += `<${names.el(local)} ${names.attr("type")}="${type}" ${relationshipAttribute(names, input.relationship(story))}/>`
		}
	}

	return out
}

export interface SectionMarkupInput {
	readonly part: StoryPart
	readonly names: Names
	// The opened sections (source references by story).
	readonly opened: readonly Section[]
	// The relationship id of the main part to a header or footer story's part.
	readonly relationship: (story: number) => string
	// Parsed properties of a source w:sectPr.
	readonly props: (node: number) => SectionProps
	// The w:id of a w:sectPrChange written.
	readonly revisionId: () => number
}

// The children a section edit replaces, removes ("") or adds, by local name.
// `previous`: as a w:sectPrChange records the properties before (no references, no change of its own).
function changedChildren(input: SectionMarkupInput, edit: SectionEdit, previous: boolean): Map<string, string> {
	const { names, part } = input
	const changes = new Map<string, string>()
	const node = sectionNode(part, edit.base)
	const source = node < 0 ? undefined : input.props(node)

	if (edit.rejected === true && node >= 0) {
		const document = part.document
		const recorded = children(part, node, "sectPrChange").flatMap(change => children(part, change, "sectPr"))[0]

		for (const local of SECTION_PROPERTIES) {
			if (local === "headerReference" || local === "footerReference" || local === "sectPrChange" || recorded === undefined) continue

			changes.set(
				local,
				children(part, recorded, local)
					.map(child => decodeUtf8(document.bytes, document.start(child), document.end(child)))
					.join("")
			)
		}
	}

	if (previous) {
		changes.set("headerReference", "")
		changes.set("footerReference", "")
	} else if (edit.headers !== undefined || edit.footers !== undefined) {
		changes.set("headerReference", references(input, node, "headerReference", edit.headers))
		changes.set("footerReference", references(input, node, "footerReference", edit.footers))
	}

	if (edit.footnotes !== undefined) changes.set("footnotePr", noteProperties(names, "footnotePr", edit.footnotes))
	if (edit.endnotes !== undefined) changes.set("endnotePr", noteProperties(names, "endnotePr", edit.endnotes))
	if (edit.type !== undefined) changes.set("type", edit.type === null || edit.type === "nextPage" ? "" : names.val("type", edit.type))

	if (edit.pageSize !== undefined) {
		const { width, height, landscape } = edit.pageSize
		const size = source?.pageSize
		// Word's paper code stays with the size it names.
		const code = size?.code !== undefined && size.width === width && size.height === height ? String(size.code) : undefined

		changes.set(
			"pgSz",
			names.empty("pgSz", [
				["w", String(width)],
				["h", String(height)],
				["orient", landscape ? "landscape" : undefined],
				["code", code]
			])
		)
	}

	if (edit.pageMargins !== undefined) {
		const margins = edit.pageMargins

		changes.set(
			"pgMar",
			names.empty(
				"pgMar",
				(["top", "right", "bottom", "left", "header", "footer", "gutter"] as const).map(side => [side, String(margins[side])])
			)
		)
	}

	if (edit.lineNumbering !== undefined) {
		const lines = edit.lineNumbering

		changes.set(
			"lnNumType",
			lines === null
				? ""
				: names.empty("lnNumType", [
						["countBy", String(lines.countBy)],
						// Word numbers from start + 1 (MS-OI29500 2.1.215).
						["start", lines.start > 1 ? String(lines.start - 1) : undefined],
						["distance", lines.distance === undefined ? undefined : String(lines.distance)],
						["restart", lines.restart]
					])
		)
	}

	if (edit.pageNumbering !== undefined) {
		const numbering = edit.pageNumbering
		const chapters = source?.pageNumbering
		const attributes: [string, string | undefined][] = [
			["fmt", numbering.format ?? undefined],
			["start", numbering.start === null ? undefined : String(numbering.start)],
			["chapStyle", chapters?.chapterStyle === undefined ? undefined : String(chapters.chapterStyle)],
			["chapSep", chapters?.chapterSeparator]
		]

		changes.set("pgNumType", attributes.every(([, value]) => value === undefined) ? "" : names.empty("pgNumType", attributes))
	}

	if (edit.columns !== undefined) {
		const { count, space, separator } = edit.columns

		changes.set(
			"cols",
			names.empty("cols", [
				["num", count > 1 ? String(count) : undefined],
				["space", String(space)],
				["sep", separator && count > 1 ? "1" : undefined]
			])
		)
	}

	if (edit.titlePage !== undefined) changes.set("titlePg", onOff(names, "titlePg", edit.titlePage))

	if (previous) {
		changes.set("sectPrChange", "")
	} else if (edit.change !== undefined) {
		const change = edit.change

		if (change === null) {
			changes.set("sectPrChange", "")
		} else {
			// The properties before, without references (ECMA-376 17.13.5.32).
			const before = sectionMarkupOf(input, change.before ?? { section: edit.section, base: edit.base }, true)
			const tag = names.empty("sectPrChange", [
				["id", String(input.revisionId())],
				["author", change.author],
				["date", change.date === "" ? undefined : change.date]
			])

			changes.set("sectPrChange", `${tag.slice(0, -2)}>${before}</${names.el("sectPrChange")}>`)
		}
	}

	return changes
}

// The WordprocessingML children of `node` named `local`.
function children(part: StoryPart, node: number, local: string): number[] {
	const document = part.document
	const out: number[] = []

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child))
		if (part.reader.is(child, Ns.W, local)) out.push(child)

	return out
}

function sectionMarkupOf(input: SectionMarkupInput, edit: SectionEdit, previous: boolean): string {
	const changes = changedChildren(input, edit, previous)
	const node = sectionNode(input.part, edit.base)

	if (node >= 0) return markup(out => spliceChildren(out, input.part.reader, node, changes, SECTION_PROPERTIES))

	return freshContainer(input.names, "sectPr", changes, SECTION_PROPERTIES) || input.names.empty("sectPr")
}

// The w:sectPr a section edit makes: the base's with what it changed spliced in (in schema order), or a new one.
export function sectionMarkup(input: SectionMarkupInput, edit: SectionEdit): string {
	return sectionMarkupOf(input, edit, false)
}

// ---- sections as layout reads them --------------------------------------------------------------------------

// Each part's names and a root start tag binding what its root binds (and r), for the markup layout parses.
const PART_ROOTS = new WeakMap<StoryPart, { readonly names: Names; readonly open: string; readonly close: string }>()

function partRoot(part: StoryPart): { readonly names: Names; readonly open: string; readonly close: string } {
	let found = PART_ROOTS.get(part)

	if (found === undefined) {
		const document = part.document
		const root = document.documentElement()
		const namespaces = partReader(document).namespaces
		const declared = namespaces.declarations().get(root) ?? new Map<string, string>()
		const w = namespaces.prefixFor(root, W_NAMESPACE) ?? "w"
		const r = namespaces.prefixFor(root, R_NAMESPACE)
		let open = "<filenSection"

		for (const [prefix, uri] of declared) open += ` ${prefix === "" ? "xmlns" : `xmlns:${prefix}`}="${escapeAttribute(uri)}"`
		if (!declared.has(w)) open += ` xmlns${w === "" ? "" : `:${w}`}="${W_NAMESPACE}"`

		found = { names: new Names(w, r ?? ""), open: `${open}>`, close: "</filenSection>" }
		PART_ROOTS.set(part, found)
	}

	return found
}

// A placeholder relationship id for the markup layout parses (layout takes stories from the section itself).
function storyPlaceholder(story: number): string {
	return `filenStory${String(story)}`
}

// A section whose properties an edit changed: its w:sectPr is a node of `document`, not of the main part.
export interface EditedSection extends Section {
	readonly document?: XmlDocument
}

// The document a section's w:sectPr node is in (the main part's when undefined).
export function sectionDocument(section: Section): XmlDocument | undefined {
	return (section as EditedSection).document
}

// Builds the sections of an edited body for layout (DocumentEdits.resection): one Section per section end,
// the source's own object where nothing changed, else one whose w:sectPr is the markup the writer writes (in a
// document of its own) with its own references as edited. Each is made once per edit.
export class EditedSections {
	private readonly opened: readonly Section[]
	private readonly part: StoryPart | undefined
	private readonly props: (document: XmlDocument, node: number) => SectionProps
	private readonly made = new WeakMap<SectionEdit, EditedSection>()
	private readonly fresh = new Map<number, Section>()

	constructor(opened: readonly Section[], part: StoryPart | undefined, props: (document: XmlDocument, node: number) => SectionProps) {
		this.opened = opened
		this.part = part
		this.props = props
	}

	// The section a break (or the body's own properties) ends, before its index and inherited references.
	base(end: SectionEnd): EditedSection {
		const edit = end.edit
		const source = openedSection(this.opened, end.key)

		if (edit === undefined || this.part === undefined) return source ?? this.defaultSection(end.key)

		const known = this.made.get(edit)

		if (known !== undefined) return known

		const part = this.part
		const root = partRoot(part)
		const sectPr = sectionMarkup(
			{
				part,
				names: root.names,
				opened: this.opened,
				relationship: storyPlaceholder,
				props: node => this.props(part.document, node),
				revisionId: () => 0
			},
			edit
		)
		const document = parseXml(encodeUtf8(`${root.open}${sectPr}${root.close}`))
		const based = openedSection(this.opened, edit.base)
		const template = source ?? based ?? this.defaultSection(end.key)
		const own = (kind: "headers" | "footers"): HeaderFooterSet => ownReferences(based?.[kind] ?? NO_HEADERS, edit[kind])
		const section: EditedSection = {
			...template,
			item: end.key,
			sectPr: document.firstChildElement(document.documentElement()),
			document,
			headers: own("headers"),
			footers: own("footers")
		}

		this.made.set(edit, section)

		return section
	}

	private defaultSection(key: number): Section {
		let section = this.fresh.get(key)

		if (section === undefined) {
			const end = this.opened[this.opened.length - 1]?.end ?? 0

			section = {
				index: this.opened.length,
				item: key,
				sectPr: -1,
				start: end,
				end,
				openWrappers: [],
				headers: NO_HEADERS,
				footers: NO_HEADERS,
				effectiveHeaders: NO_HEADERS,
				effectiveFooters: NO_HEADERS
			}
			this.fresh.set(key, section)
		}

		return section
	}
}

// Whether two section bases lay out alike (the same break, properties and own references).
export function sameSection(a: Section | undefined, b: Section | undefined): boolean {
	if (a === undefined || b === undefined) return a === b

	return (
		a.item === b.item &&
		a.sectPr === b.sectPr &&
		sectionDocument(a) === sectionDocument(b) &&
		sameSet(a.headers, b.headers) &&
		sameSet(a.footers, b.footers)
	)
}

function sameSet(a: HeaderFooterSet, b: HeaderFooterSet): boolean {
	return a.default === b.default && a.first === b.first && a.even === b.even
}

// The sections `bases` make with their indices and the references each inherits (ECMA-376 17.10.5), from
// section `from` on; `previous` the section before it.
export function inheritedSections(bases: readonly Section[], from: number, previous: Section | undefined): Section[] {
	let headers = previous?.effectiveHeaders ?? NO_HEADERS
	let footers = previous?.effectiveFooters ?? NO_HEADERS
	const out: Section[] = []

	for (let k = from; k < bases.length; k++) {
		const base = bases[k] as Section

		headers = inheritHeaders(base.headers, headers)
		footers = inheritHeaders(base.footers, footers)
		out.push({ ...base, index: k, effectiveHeaders: headers, effectiveFooters: footers })
	}

	return out
}

// ---- sections as edited, with their header and footer stories --------------------------------------------------

export interface SectionState extends BodySection {
	readonly index: number
	// Own references (story indices, -1 none) and those shown (inherited where it has none).
	readonly headers: HeaderFooterSet
	readonly footers: HeaderFooterSet
	readonly effectiveHeaders: HeaderFooterSet
	readonly effectiveFooters: HeaderFooterSet
}

const STATES = new WeakMap<Node, readonly SectionState[]>()

// The body's sections with the header and footer stories each has and shows, as layout lays them out.
export function sectionStates(opened: readonly Section[], doc: Node, part: StoryPart): readonly SectionState[] {
	const known = STATES.get(doc)

	if (known !== undefined) return known

	let headers = NO_HEADERS
	let footers = NO_HEADERS
	const out = bodySections(doc, part).map((section, index): SectionState => {
		const edit = section.edit
		const base = openedSection(opened, edit?.base ?? section.key)
		const own = (kind: "headers" | "footers"): HeaderFooterSet => ownReferences(base?.[kind] ?? NO_HEADERS, edit?.[kind])
		const ownHeaders = own("headers")
		const ownFooters = own("footers")

		headers = inheritHeaders(ownHeaders, headers)
		footers = inheritHeaders(ownFooters, footers)

		return { ...section, index, headers: ownHeaders, footers: ownFooters, effectiveHeaders: headers, effectiveFooters: footers }
	})

	STATES.set(doc, out)

	return out
}

// Properties of a section as edited, parsed (from the markup the writer writes, references as placeholders).
export function sectionProps(context: DocxEditContext, part: StoryPart, end: SectionEnd): SectionProps {
	const edit = end.edit

	if (edit === undefined) {
		const node = sectionNode(part, end.key)

		return node < 0 ? {} : context.store.section(part.document, node)
	}

	const root = partRoot(part)
	const markup = sectionMarkup(
		{
			part,
			names: root.names,
			opened: context.model.sections,
			relationship: storyPlaceholder,
			props: node => context.store.section(part.document, node),
			revisionId: () => 0
		},
		edit
	)
	const document = parseXml(encodeUtf8(`${root.open}${markup}${root.close}`))

	return context.store.section(document, document.firstChildElement(document.documentElement()))
}

// The edit a section has, or the one standing for its source as it is.
export function editOf(end: SectionEnd): SectionEdit {
	return end.edit ?? { section: end.key, base: end.key }
}

// `next` recorded as a tracked change of `previous` where tracking records one (`stamp`): the first change
// keeps the properties before it; a source w:sectPrChange stays as Word keeps the earliest.
export function trackedEdit(
	previous: SectionEdit,
	next: SectionEdit,
	stamp: RevisionStamp | undefined,
	sourceChange: boolean
): SectionEdit {
	if (stamp === undefined) return next
	if (previous.change !== undefined && previous.change !== null) return { ...next, change: previous.change }
	if (sourceChange && previous.change === undefined) return next

	const { change: _change, headers: _headers, footers: _footers, ...before } = previous

	return {
		...next,
		change: {
			author: stamp.author,
			date: stamp.date,
			...(stamp.id === undefined ? {} : { id: stamp.id }),
			before: Object.keys(before).length > 2 ? before : null
		}
	}
}

// Whether a section's properties can be edited: it ends with a paragraph or the body's own properties, not a
// locked block (written as its bytes, break included).
export function editableSection(section: BodySection): boolean {
	return section.node === undefined || section.node.type === nodes.paragraph
}

// Sets the edit of section `section` in `tr` (the body's transaction): the section_change mark of its
// paragraph, or the doc's sectionChanges.
export function setSectionEdit(tr: Transaction, section: BodySection, edit: SectionEdit): void {
	if (section.pos < 0) {
		tr.setDocAttribute("sectionChanges", Object.freeze({ ...docSectionChanges(tr.doc), final: Object.freeze(edit) }))

		return
	}

	tr.addNodeMark(section.pos, marks.section_change.create({ value: Object.freeze(edit) }))
}

// ---- commands -------------------------------------------------------------------------------------------------

// The body's session, its part and sections, for commands of the paged editor.
export interface BodySections {
	readonly session: DocxEditorSession
	readonly part: StoryPart
	readonly states: readonly SectionState[]
}

export function bodySectionsOf(editor: PagedEditor): BodySections {
	const session = editor.body.session

	return {
		session,
		part: session.story.part,
		states: sectionStates(session.context.model.sections, session.state.doc, session.story.part)
	}
}

// The section the caret is in: the body's caret's, or (in a header or footer) the section of the page its caret
// is drawn on; in another story, the body's caret's.
export function caretSection(editor: PagedEditor): number {
	const { session, part, states } = bodySectionsOf(editor)
	const kind = editor.session.story.story.kind

	if (kind === "header" || kind === "footer") {
		const page = pageOfCaret(editor)
		const section = page === undefined ? undefined : editor.page(page)?.section

		if (section !== undefined) return Math.min(Math.max(0, section), states.length - 1)
	}

	return sectionIndexAt(session.state.doc, part, session.state.selection.head)
}

// The page the active story's caret is drawn on.
export function pageOfCaret(editor: PagedEditor): number | undefined {
	return editor.geometry.locate(editor.session.state.selection.head, editor.affinity)?.page
}

// The stamp section changes are recorded with: undefined when not tracking; null when tracking wants an author
// the document has none of yet (the command is refused).
export function sectionStamp(context: DocxEditContext): RevisionStamp | undefined | null {
	if (!context.review.tracking) return undefined

	return revisionStamp(context.identity) ?? null
}

// Whether the source properties of a section hold a tracked change of their own.
function sourceChanged(context: DocxEditContext, part: StoryPart, edit: SectionEdit): boolean {
	const node = sectionNode(part, edit.base)

	return node >= 0 && context.store.section(part.document, node).change !== undefined
}

// Applies `update` to sections `indices` (one body step); property changes are tracked while tracking is on
// (`tracked`), references never are. False when nothing changed or the body cannot be edited.
export function editSections(
	editor: PagedEditor,
	indices: readonly number[],
	update: (state: SectionState, edit: SectionEdit) => SectionEdit | undefined,
	tracked: boolean
): boolean {
	const { session, part, states } = bodySectionsOf(editor)
	const context = session.context
	const stamp = tracked ? sectionStamp(context) : undefined

	if (session.readOnly || stamp === null) return false

	const tr = session.state.tr
	let changed = false

	for (const index of indices) {
		const state = states[index]

		if (state === undefined || !editableSection(state)) continue

		const previous = editOf(state)
		const next = update(state, previous)

		if (next === undefined) continue

		const recorded = tracked && stamp !== undefined ? { ...stamp, id: context.ids.next("revision") } : undefined

		setSectionEdit(
			tr,
			state,
			recorded === undefined ? next : trackedEdit(previous, next, recorded, sourceChanged(context, part, previous))
		)
		changed = true
	}

	return changed && dispatchBody(editor, tr)
}

// A section change made in the body while another story is edited (a header's chip): that story's snapshot
// (its story context) is made again.
function dispatchBody(editor: PagedEditor, tr: Transaction): boolean {
	const done = editor.body.session.dispatch(markInternal(tr))

	if (done && editor.session !== editor.body.session) editor.session.invalidate()

	return done
}

function allSections(states: readonly SectionState[]): number[] {
	return states.map((_state, index) => index)
}

// Whether the caret is where Word inserts a section break: body text outside tables, frames and locked blocks.
export function canBreak(editor: PagedEditor): boolean {
	const session = editor.session

	if (session !== editor.body.session || session.readOnly) return false

	const { $from, $to } = session.state.selection

	if (!$from.parent.isTextblock || $from.parent.type !== nodes.paragraph || !$from.sameParent($to)) return false

	for (let depth = $from.depth; depth > 0; depth--) if ($from.node(depth).type === nodes.table_cell) return false

	return true
}

// A section break of `kind` at the caret: the paragraph splits there and its first half ends a new section
// with the properties (and header and footer references) of the one it was in, as Word makes it; the section
// after it starts with `kind`, shows the same headers and footers (linked to the new one) and continues the
// page numbers. While tracking, Word's record: the first half's paragraph mark (holding the new w:sectPr) is
// inserted, and what the section after it changed is a property change of that section, one change with it.
export function insertSectionBreak(editor: PagedEditor, kind: string): boolean {
	if (!canBreak(editor)) return false

	const { session, part, states } = bodySectionsOf(editor)
	const context = session.context
	const stamp = sectionStamp(context)

	if (stamp === null) return false

	const recorded = stamp === undefined ? undefined : { ...stamp, id: context.ids.next("revision") }
	const state = session.state
	const tr = state.tr.deleteSelection()
	const $pos = tr.selection.$from
	const paragraph = $pos.parent
	const start = $pos.before()
	const index = sectionIndexAt(state.doc, part, state.selection.head)
	const section = states[index]

	if (section === undefined || paragraph.type !== nodes.paragraph || !editableSection(section)) return false

	const key = newSectionKey(context)
	const current = editOf(section)
	const base = current.base
	// The later half is the paragraph as it was (its mark, its break); the first half ends with a new mark.
	const split = splitParagraphMark(paragraph, session, recorded)
	const firstMarks = paragraph.marks.filter(mark => mark.type !== marks.section_change && mark.type !== marks.paragraph_change)
	const laterMarks = paragraph.marks.filter(mark => mark.type !== marks.para_id && mark.type !== marks.paragraph_change)
	const fresh: SectionEdit = {
		...current,
		section: key,
		base,
		headers: { default: section.headers.default, first: section.headers.first, even: section.headers.even },
		footers: { default: section.footers.default, first: section.footers.first, even: section.footers.even },
		change: null
	}
	const props = sectionProps(context, part, section)
	const pageStart = props.pageNumbering?.start
	const linked: SectionEdit = {
		...current,
		type: kind === "nextPage" ? null : kind,
		headers: { default: null, first: null, even: null },
		footers: { default: null, first: null, even: null },
		...(pageStart === undefined ? {} : { pageNumbering: { format: props.pageNumbering?.format ?? null, start: null } })
	}

	tr.split($pos.pos)
	tr.setNodeMarkup(start, undefined, { ...paragraph.attrs, sectPr: key, revision: split.first.revision }, [
		...withParagraphChange(firstMarks, split.first.change),
		marks.section_change.create({ value: Object.freeze(fresh) })
	])

	const first = tr.doc.nodeAt(start)
	const later = first === null ? -1 : start + first.nodeSize
	const laterNode = later < 0 ? null : tr.doc.nodeAt(later)

	if (laterNode === null) return false

	tr.setNodeMarkup(
		later,
		undefined,
		{ ...laterNode.attrs, before: NO_MARKERS, revision: split.later.revision },
		withParagraphChange(laterMarks, split.later.change)
	)

	const next = recorded === undefined ? linked : trackedEdit(current, linked, recorded, sourceChanged(context, part, current))

	if (section.pos < 0) {
		setSectionEdit(tr, section, next)
	} else {
		// The caret's paragraph ending the section: its break is the later half's now.
		const mapped = tr.mapping.map(section.pos)
		const pos = mapped === start ? later : mapped

		if (tr.doc.nodeAt(pos)?.attrs["sectPr"] !== section.key) return false

		tr.addNodeMark(pos, marks.section_change.create({ value: Object.freeze(next) }))
	}

	tr.setSelection(TextSelection.create(tr.doc, later + 1))

	return session.dispatch(markInternal(tr).scrollIntoView())
}

// Whether the caret's section ends with a break the command can remove (while tracking, one not deleted yet).
export function canDeleteBreak(editor: PagedEditor): boolean {
	const { session, states } = bodySectionsOf(editor)

	if (session.readOnly) return false

	const index = caretSection(editor)
	const node = states[index]?.node

	return (
		index < states.length - 1 &&
		node?.type === nodes.paragraph &&
		!(session.context.review.tracking && ((node.attrs["revision"] as number) & RevisionFlag.Deleted) !== 0)
	)
}

// Removes the break ending the caret's section: its text joins the section after it, which keeps its own
// properties, as Word combines them ("the new combined section will use the formatting of the second
// section"), and the headers and footers it showed. While tracking, Word's record: the break's paragraph mark
// is deleted (accepted, the sections combine so); a break the author inserted while tracking goes for real,
// as rejecting it takes it out.
export function deleteSectionBreak(editor: PagedEditor): boolean {
	if (!canDeleteBreak(editor)) return false

	const { session, part, states } = bodySectionsOf(editor)
	const context = session.context
	const index = caretSection(editor)
	const section = states[index]

	if (section === undefined || section.node === undefined) return false

	const tr = session.state.tr
	const node = section.node

	if (context.review.tracking) {
		const stamp = sectionStamp(context)

		if (stamp === null || stamp === undefined) return false

		const change = paragraphChange(node)
		const inserted = ((node.attrs["revision"] as number) & RevisionFlag.Inserted) !== 0 ? change?.inserted : undefined

		if (inserted?.id !== undefined && inserted.author === stamp.author) {
			const outcome = resolveChanges(session, storyReviewIndex(session), { id: String(inserted.id) }, false)

			return outcome.tr !== null && session.dispatch(outcome.tr)
		}

		tr.setNodeAttribute(section.pos, "revision", (node.attrs["revision"] as number) | RevisionFlag.Deleted)
		setParagraphChange(tr, section.pos, node, { ...change, deleted: { ...stamp, id: context.ids.next("revision") } })

		return session.dispatch(markInternal(tr))
	}

	tr.setNodeMarkup(
		section.pos,
		undefined,
		{ ...node.attrs, sectPr: NO_SOURCE },
		node.marks.filter(mark => mark.type !== marks.section_change)
	)
	keepShown(tr, context, part, states, new Set([index]))

	return session.dispatch(markInternal(tr))
}

// After the breaks of sections `removed` (indices of `before`, the body's sections as they were) went, each
// section their text joined keeps showing what it showed through its link to them: their own header and footer
// references, nearest first, where it has none of a type. `tr`'s document has the breaks taken out.
export function keepShown(
	tr: Transaction,
	context: DocxEditContext,
	part: StoryPart,
	before: readonly SectionState[],
	removed: ReadonlySet<number>
): void {
	if (removed.size === 0) return

	const inherited = (own: HeaderFooterSet, from: HeaderFooterSet): HeaderReferenceEdits | undefined => {
		let out: Record<string, number> | undefined

		for (const type of HEADER_TYPES) if (own[type] < 0 && from[type] >= 0) (out ??= {})[type] = from[type]

		return out
	}
	let shift = 0

	for (let index = 0; index < before.length; index++) {
		if (removed.has(index)) {
			shift++

			continue
		}

		for (let gone = index - 1; gone >= 0 && removed.has(gone); gone--) {
			const now = sectionStates(context.model.sections, tr.doc, part)[index - shift]
			const was = before[gone] as SectionState

			if (now === undefined || now.key !== before[index]?.key || !editableSection(now)) break

			const headers = inherited(now.headers, was.headers)
			const footers = inherited(now.footers, was.footers)

			if (headers === undefined && footers === undefined) continue

			const edit = editOf(now)

			setSectionEdit(tr, now, {
				...edit,
				...(headers === undefined ? {} : { headers: { ...edit.headers, ...headers } }),
				...(footers === undefined ? {} : { footers: { ...edit.footers, ...footers } })
			})
		}
	}
}

// The twips in a measure parsed from a w:sectPr.
function twips(value: number | undefined, fallback: number): number {
	return Math.round(value ?? fallback)
}

// Page setup of the caret's section (or every section): what `setup` names changes, the rest stays.
export function pageSetup(editor: PagedEditor, setup: DocxPageSetup): boolean {
	const { states } = bodySectionsOf(editor)
	const context = editor.body.session.context
	const part = editor.body.session.story.part
	const indices = setup.scope === "document" ? allSections(states) : [caretSection(editor)]

	return editSections(
		editor,
		indices,
		(state, edit) => {
			const current = resolveSection(sectionProps(context, part, state))
			let next: SectionEdit = edit

			if (setup.size !== undefined || setup.orientation !== undefined) {
				let width = setup.size?.width ?? current.pageWidth
				let height = setup.size?.height ?? current.pageHeight
				const landscape = setup.orientation === undefined ? width > height : setup.orientation === "landscape"

				// w:orient does not turn the page: its width and height do (ECMA-376 17.6.13).
				if (landscape !== width > height) [width, height] = [height, width]

				next = { ...next, pageSize: { width: Math.round(width), height: Math.round(height), landscape } }
			}

			if (setup.margins !== undefined) {
				const margins = current.margins

				next = {
					...next,
					pageMargins: {
						top: setup.margins.top ?? twips(margins.top, 1440),
						right: setup.margins.right ?? twips(margins.right, 1440),
						bottom: setup.margins.bottom ?? twips(margins.bottom, 1440),
						left: setup.margins.left ?? twips(margins.left, 1440),
						header: setup.margins.header ?? twips(margins.header, 720),
						footer: setup.margins.footer ?? twips(margins.footer, 720),
						gutter: setup.margins.gutter ?? twips(margins.gutter, 0)
					}
				}
			}

			if (setup.columns !== undefined) {
				next = {
					...next,
					columns: {
						count: setup.columns.count,
						space: setup.columns.space ?? twips(current.columns.space, 720),
						separator: setup.columns.separator ?? current.columns.separator
					}
				}
			}

			if (setup.pageNumbers !== undefined) {
				const numbering = current.pageNumbering

				next = {
					...next,
					pageNumbering: {
						format: setup.pageNumbers.format ?? numbering.format ?? null,
						start: setup.pageNumbers.start === undefined ? (numbering.start ?? null) : setup.pageNumbers.start
					}
				}
			}

			if (setup.lineNumbers !== undefined) {
				const lines = setup.lineNumbers

				next = {
					...next,
					lineNumbering:
						lines === null
							? null
							: {
									countBy: lines.countBy,
									start: lines.start ?? 1,
									...(lines.distance === undefined ? {} : { distance: lines.distance }),
									...(lines.restart === undefined ? {} : { restart: lines.restart })
								}
				}
			}

			return next === edit ? undefined : next
		},
		true
	)
}

// Note numbering and placement of the caret's section (or every section): w:footnotePr / w:endnotePr in its
// properties, as Word's Footnote and Endnote dialog writes them.
export function noteOptions(editor: PagedEditor, options: DocxNoteOptions): boolean {
	const { states } = bodySectionsOf(editor)
	const context = editor.body.session.context
	const part = editor.body.session.story.part
	const indices = options.scope === "document" ? allSections(states) : [caretSection(editor)]
	const key = options.kind === "footnote" ? "footnotes" : "endnotes"

	return editSections(
		editor,
		indices,
		(state, edit) => {
			const current = sectionProps(context, part, state)[key] ?? {}
			const next: NotePositionProps = {
				...current,
				...(options.position === undefined ? {} : { position: options.position }),
				...(options.format === undefined ? {} : { numFormat: options.format }),
				...(options.start === undefined ? {} : { numStart: options.start }),
				...(options.restart === undefined ? {} : { numRestart: options.restart })
			}

			return { ...edit, [key]: next }
		},
		true
	)
}

export function setTitlePage(editor: PagedEditor, on: boolean): boolean {
	return editSections(editor, [caretSection(editor)], (_state, edit) => ({ ...edit, titlePage: on }), true)
}

// The document-wide switch for different even and odd headers (w:evenAndOddHeaders in settings).
export function setEvenAndOddHeaders(editor: PagedEditor, on: boolean): boolean {
	const session = editor.body.session

	if (session.readOnly) return false

	const tr = session.state.tr.setDocAttribute(
		"sectionChanges",
		Object.freeze({ ...docSectionChanges(session.state.doc), evenAndOddHeaders: on })
	)

	return dispatchBody(editor, tr)
}

// Header and footer stories the section commands added, by what the context records of them: the save asks
// whether a section refers to one (only sections do).
const SECTION_STORIES = new WeakMap<object, { readonly context: DocxEditContext; readonly index: number }>()

export function registerSectionStory(context: DocxEditContext, story: Story): void {
	const made = context.newStories.get(story)

	if (made !== undefined) SECTION_STORIES.set(made, { context, index: story.index })
}

// Whether a section of the body (`docs`, as written) refers to added story `made`.
export function sectionStoryReferenced(made: object, docs: readonly Node[]): boolean {
	const known = SECTION_STORIES.get(made)

	if (known === undefined) return false

	const model = known.context.model
	const body = model.body
	const part = body === undefined ? undefined : model.parts[body.part]
	const doc = docs.find(candidate => candidate.attrs["story"] === body?.index)

	if (part === undefined || doc === undefined) return false

	return sectionStates(model.sections, doc, part).some(state =>
		[state.headers, state.footers].some(set => set.default === known.index || set.first === known.index || set.even === known.index)
	)
}

// ---- tracked changes of section properties ----------------------------------------------------------------------

// The id a section's tracked change counts under in the review: a recorded change's own, else the source's
// paragraph (or final properties) item, as review/index.ts counts property changes.
export function sectionChangeId(state: BodySection, context: DocxEditContext, part: StoryPart): string | undefined {
	const edit = state.edit
	const change = edit?.change

	if (change !== undefined && change !== null) return change.id === undefined ? `#${String(state.key)}` : String(change.id)
	if (change === null || edit?.rejected === true) return undefined

	return sourceChanged(context, part, editOf(state)) ? `#${String(edit?.base ?? state.key)}` : undefined
}

// Accepts (or rejects) the tracked property changes of the body's sections whose break lies in [from, to)
// (the whole body when undefined) or whose change is `id`, in `tr` (positions of its document, which the steps
// added here do not move). Accepted, a change's properties stay and its record goes; rejected, the properties
// it recorded come back (references stay, as Word records none). The number resolved.
export function resolveSectionChanges(
	tr: Transaction,
	context: DocxEditContext,
	part: StoryPart,
	accept: boolean,
	scope: { readonly from: number; readonly to: number } | { readonly id: string } | undefined
): number {
	const doc = tr.doc
	let resolved = 0

	for (const state of bodySections(doc, part)) {
		const id = sectionChangeId(state, context, part)

		if (id === undefined || !editableSection(state)) continue
		if (scope !== undefined && "id" in scope && scope.id !== id) continue
		if (scope !== undefined && "from" in scope) {
			const end = state.pos < 0 ? doc.content.size : state.pos + (state.node?.nodeSize ?? 1) - 1

			if (end < scope.from || end >= Math.max(scope.to, scope.from + 1)) continue
		}

		const edit = editOf(state)
		const change = edit.change
		let next: SectionEdit

		if (change !== undefined && change !== null) {
			next = accept
				? { ...edit, change: null }
				: {
						...(change.before ?? {}),
						section: edit.section,
						base: edit.base,
						...(edit.headers === undefined ? {} : { headers: edit.headers }),
						...(edit.footers === undefined ? {} : { footers: edit.footers }),
						change: null
					}
		} else {
			next = accept ? { ...edit, change: null } : { ...edit, rejected: true, change: null }
		}

		setSectionEdit(tr, state, next)
		resolved++
	}

	return resolved
}
