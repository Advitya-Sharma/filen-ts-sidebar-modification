import { NodeKind, parseXml, type XmlDocument } from "@filen/office/core/xml"

// A LibreOffice flat ODF text document (FODT) as text and counts, to compare with what our model reads of
// the DOCX LibreOffice converted it from. ODF keeps a deletion's text out of the body, in
// text:tracked-changes, and marks an insertion's range with text:change-start/-end, so the text with all
// changes accepted and with all rejected both come from one file. List labels are not in ODF text.

export interface OdfSummary {
	// Body paragraphs, one line each (table cells and text frames included; notes, annotations, shapes and
	// pictures not).
	readonly final: string
	readonly original: string
	readonly counts: Readonly<Record<string, number>>
}

// Subtrees whose text is not body text.
const SKIPPED = new Set([
	"text:note",
	"office:annotation",
	"draw:frame",
	"draw:custom-shape",
	"draw:g",
	"draw:rect",
	"draw:line",
	"draw:ellipse",
	"draw:polygon",
	"draw:path",
	"draw:connector",
	"draw:control",
	"text:tracked-changes",
	"text:ruby-text",
	"text:table-of-content-source",
	"text:alphabetical-index-source",
	"text:illustration-index-source",
	"text:table-index-source",
	"text:object-index-source",
	"text:user-index-source",
	"text:bibliography-source",
	"text:index-title-template",
	"office:forms",
	"text:sequence-decls",
	"text:variable-decls",
	"text:user-field-decls",
	// Computed fields (fuzz/saves.ts COMPUTED_FIELDS): LibreOffice computes them again.
	"text:page-number",
	"text:page-count",
	"text:page-continuation",
	"text:date",
	"text:time",
	"text:file-name",
	"text:author-name",
	"text:author-initials",
	"text:creation-date",
	"text:creation-time",
	"text:modification-date",
	"text:modification-time",
	"text:print-date",
	"text:print-time",
	"text:editing-duration",
	"text:word-count",
	"text:character-count",
	"text:paragraph-count",
	"text:title",
	"text:subject",
	"text:keywords",
	"text:description",
	"text:creator",
	"text:initial-creator",
	"text:printed-by",
	"text:user-defined",
	"text:template-name",
	"text:chapter",
	"text:bookmark-ref",
	"text:note-ref",
	"text:sequence",
	"text:sequence-ref",
	"text:reference-ref"
])
// Indexes: their paragraphs count, their text (LibreOffice's own entries and page numbers) does not. A TOC field
// becomes a table of contents, or with \c an illustration or table index (our side leaves every TOC field's
// result out).
const MUTED = new Set([
	"text:table-of-content",
	"text:illustration-index",
	"text:table-index",
	"text:object-index",
	"text:user-index",
	// A macro button: its display text (fuzz/saves.ts COMPUTED_FIELDS).
	"text:execute-macro"
])

function children(document: XmlDocument, node: number): number[] {
	const out: number[] = []

	for (let child = document.firstChild(node); child >= 0; child = document.nextSibling(child)) out.push(child)

	return out
}

function find(document: XmlDocument, node: number, name: string): number {
	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child))
		if (document.name(child) === name) return child

	return -1
}

interface Region {
	readonly kind: string
	// A deletion's text, one entry per paragraph.
	readonly deleted: string[]
}

// Changed regions by id: what kind of change, and a deletion's text.
function regions(document: XmlDocument, text: number): Map<string, Region> {
	const out = new Map<string, Region>()
	const tracked = find(document, text, "text:tracked-changes")

	if (tracked < 0) return out

	for (let region = document.firstChildElement(tracked); region >= 0; region = document.nextSiblingElement(region)) {
		const id = document.attribute(region, "text:id") ?? document.attribute(region, "xml:id")
		const change = document.firstChildElement(region)

		if (id === undefined || change < 0) continue

		const kind = document.name(change).replace(/^text:/, "")
		const deleted: string[] = []

		// Its paragraphs, also those in lists and tables (the whitespace between elements is not text).
		const collect = (node: number): void => {
			for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
				const name = document.name(child)

				if (name === "text:p" || name === "text:h") deleted.push(paragraphText(document, child))
				else if (name !== "office:change-info") collect(child)
			}
		}

		if (kind === "deletion") collect(change)

		out.set(id, { kind, deleted })
	}

	return out
}

// A paragraph's own text (no insertion tracking: for deleted content).
function paragraphText(document: XmlDocument, node: number): string {
	let out = ""
	const visit = (at: number): void => {
		for (const child of children(document, at)) {
			const kind = document.kind(child)

			if (kind === NodeKind.Text || kind === NodeKind.CData) out += document.text(child)
			else if (kind === NodeKind.Element) {
				const name = document.name(child)

				if (SKIPPED.has(name)) continue
				if (name === "text:s") out += " ".repeat(Number(document.attribute(child, "text:c") ?? "1") || 1)
				else if (name === "text:tab") out += "\t"
				else if (name === "text:line-break") out += "\n"
				else visit(child)
			}
		}
	}

	visit(node)

	return out
}

// The draw:text-box of a text frame LibreOffice made of body content, or -1: a framed paragraph (w:framePr) or a
// floating table (w:tblpPr) becomes a frame anchored to a paragraph; text boxes are anchored to a character or
// are shapes.
function textFrame(document: XmlDocument, frame: number): number {
	const box = document.firstChildElement(frame)

	return box >= 0 && document.name(box) === "draw:text-box" && document.attribute(frame, "text:anchor-type") === "paragraph" ? box : -1
}

// Body text with every change accepted and with every change rejected, and how many paragraphs the former has.
function bodyText(document: XmlDocument, text: number): { final: string; original: string; paragraphs: number } {
	const changes = regions(document, text)
	const final: string[] = []
	const original: string[] = []
	// Open change ranges: an insertion, and a deletion LibreOffice keeps in place (one holding a table or
	// spanning cells) rather than in its changed region.
	const inserted = new Set<string>()
	const deleted = new Set<string>()
	let f = ""
	let o = ""
	let inParagraph = false
	let muted = 0
	// Text frames anchored in the paragraph being read.
	const framed: number[] = []
	const add = (value: string): void => {
		if (muted > 0) return
		if (deleted.size === 0) f += value
		if (inserted.size === 0) o += value
	}
	const visit = (node: number): void => {
		for (const child of children(document, node)) {
			const kind = document.kind(child)

			if (kind === NodeKind.Text || kind === NodeKind.CData) {
				if (inParagraph) add(document.text(child))
				continue
			}

			if (kind !== NodeKind.Element) continue

			const name = document.name(child)
			const box = name === "draw:frame" ? textFrame(document, child) : -1

			// A text frame is body text LibreOffice moved out of the flow (a framed paragraph or a floating
			// table, which ODF holds only in a frame): its paragraphs come after the one it is anchored in.
			if (box >= 0) {
				if (inParagraph) framed.push(box)
				else visit(box)
			} else if (SKIPPED.has(name)) continue
			else if (MUTED.has(name)) {
				muted++
				visit(child)
				muted--
			} else if (name === "text:p" || name === "text:h") {
				const nested = inParagraph

				if (!nested) inParagraph = true

				visit(child)

				if (!nested) {
					inParagraph = false

					// A paragraph mark inside an insertion is not in the original, inside a deletion not in the final:
					// the next paragraph runs on.
					if (deleted.size === 0) {
						final.push(f)
						f = ""
					}

					if (inserted.size === 0) {
						original.push(o)
						o = ""
					}

					for (const box of framed.splice(0)) visit(box)
				}
			} else if (name === "text:change-start") {
				const id = document.attribute(child, "text:change-id") ?? ""

				const kind = changes.get(id)?.kind

				if (kind === "insertion") inserted.add(id)
				else if (kind === "deletion") deleted.add(id)
			} else if (name === "text:change-end") {
				const id = document.attribute(child, "text:change-id") ?? ""

				inserted.delete(id)
				deleted.delete(id)
			} else if (name === "text:change") {
				const region = changes.get(document.attribute(child, "text:change-id") ?? "")

				// Muted like the text around it (a deletion in an index's entries).
				if (region?.kind === "deletion" && muted === 0) o += region.deleted.join("\n")
			} else if (name === "text:s") add(" ".repeat(Number(document.attribute(child, "text:c") ?? "1") || 1))
			else if (name === "text:tab") add("\t")
			else if (name === "text:line-break") add("\n")
			else visit(child)
		}
	}

	visit(text)

	if (f !== "") final.push(f)
	if (o !== "") original.push(o)

	return {
		final: final.join("\n"),
		original: original.join("\n"),
		paragraphs: final.length
	}
}

// Elements counted by qualified name, everywhere in the document.
function tagCounts(document: XmlDocument): Map<string, number> {
	const out = new Map<string, number>()

	for (let node = 0; node < document.count; node++) {
		if (document.kind(node) !== NodeKind.Element) continue

		const name = document.name(node)

		out.set(name, (out.get(name) ?? 0) + 1)
	}

	return out
}

// Elements named `name` of the text flow under `node`: not in notes, whose tables our reading leaves out of the
// body's (ECMA-376 Part 1 §17.11: a note is a story of its own), nor in embedded objects (a chart's data table).
function inFlow(document: XmlDocument, node: number, name: string): number {
	let count = 0

	for (let at = node + 1; at < document.count && document.start(at) < document.end(node); at++) {
		if (document.kind(at) !== NodeKind.Element) continue

		const element = document.name(at)

		if (element === "text:note" || element === "draw:object") {
			const end = document.end(at)

			while (at + 1 < document.count && document.start(at + 1) < end) at++
		} else if (element === name) count++
	}

	return count
}

const FORM_CONTROLS =
	/^form:(?:checkbox|text|textarea|listbox|combobox|date|time|formatted-text|number|button|radio|fixed-text|frame|image|file|hidden|grid|value-range|password)$/

function counts(document: XmlDocument, text: number, tags: Map<string, number>, bodyParagraphs: number): Record<string, number> {
	const tag = (name: string): number => tags.get(name) ?? 0
	const changes = regions(document, text)
	const kinds = (kind: string): number => [...changes.values()].filter(region => region.kind === kind).length
	let footnotes = 0
	let endnotes = 0
	let images = 0
	let comments = 0

	for (let node = 0; node < document.count; node++) {
		if (document.kind(node) !== NodeKind.Element) continue

		const name = document.name(node)

		if (name === "text:note") {
			if (document.attribute(node, "text:note-class") === "endnote") endnotes++
			else footnotes++
		} else if (name === "draw:frame") {
			// A picture: a frame showing an image, not an object's replacement image.
			if (
				find(document, node, "draw:image") >= 0 &&
				find(document, node, "draw:object") < 0 &&
				find(document, node, "draw:object-ole") < 0
			)
				images++
		} else if (name === "office:annotation") comments++
	}

	return {
		paragraphs: tag("text:p") + tag("text:h"),
		bodyParagraphs,
		// In the body's flow (headers repeat on every page style LibreOffice makes for a section).
		tables: inFlow(document, text, "table:table"),
		rows: inFlow(document, text, "table:table-row"),
		cells: inFlow(document, text, "table:table-cell"),
		comments,
		insertions: kinds("insertion"),
		deletions: kinds("deletion"),
		formatChanges: kinds("format-change"),
		footnotes,
		endnotes,
		headers: tag("style:header") + tag("style:header-left") + tag("style:header-first"),
		footers: tag("style:footer") + tag("style:footer-left") + tag("style:footer-first"),
		images,
		drawings: tag("draw:frame") + tag("draw:custom-shape") + tag("draw:g"),
		toc: tag("text:table-of-content"),
		masterPages: tag("style:master-page"),
		bookmarks: tag("text:bookmark") + tag("text:bookmark-start"),
		contentControls: tag("loext:content-control"),
		// LibreOffice writes MathML in its own default namespace.
		equations: tag("math") + tag("math:math"),
		forms: [...tags].reduce((sum, [name, times]) => (FORM_CONTROLS.test(name) ? sum + times : sum), 0)
	}
}

export function summarizeFodt(bytes: Uint8Array): OdfSummary {
	const document = parseXml(bytes)
	const root = document.documentElement()
	const body = find(document, root, "office:body")
	const text = body < 0 ? -1 : find(document, body, "office:text")

	if (text < 0) return { final: "", original: "", counts: {} }

	const read = bodyText(document, text)

	return {
		final: read.final,
		original: read.original,
		counts: counts(document, text, tagCounts(document), read.paragraphs)
	}
}
