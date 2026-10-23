import { elementName, OdfStyleResolver, type OdfPackage } from "../../odf/index"
import type { OdfStyleSheet } from "../../odf/styles"
import { Namespaces, assertWellFormed, decodeUtf8, encodeUtf8, parseXml, type XmlDocument } from "../../xml/index"
import { BodyFlow } from "../pages"
import { Translator, type TranslatorSettings } from "../translator"
import { PART_NAMESPACES, XML_DECLARATION } from "../wml"
import { Meaning } from "./meaning"

// The proof of a save that changed only some top-level blocks of the body, read from what it wrote there alone: each
// changed run of blocks (a window) is translated on its own and compared with the edit's meaning of it. A window
// may hold only what the document's translator reads with no state of the document's (paragraphs, headings, spans,
// spaces, tabs, line breaks, tables); translated on its own it must leave nothing a later block reads (lists,
// notes, ids, pictures, sections) and end in the state it began in, which must be the state the document had there.
// Every block outside the windows is then the source's, in the source's state, and reads as the source's
// translation does. Anything else is left to the whole file's proof.

// Elements a window's top level may hold, and what may sit inside them.
const TOP: ReadonlySet<string> = new Set(["text:p", "text:h", "table:table", "text:soft-page-break"])
const INLINE: ReadonlySet<string> = new Set(["text:span", "text:s", "text:tab", "text:line-break", "text:soft-page-break"])
const TABLE: ReadonlySet<string> = new Set([
	"table:table-column",
	"table:table-columns",
	"table:table-column-group",
	"table:table-header-columns",
	"table:table-header-rows",
	"table:table-rows",
	"table:table-row-group",
	"table:table-row",
	"table:table-cell",
	"table:covered-table-cell"
])

// Whether everything in `node` (an element of a window) is of what a window may hold.
function plain(document: XmlDocument, namespaces: Namespaces, node: number): boolean {
	const name = elementName(document, namespaces, node)
	const inside: ReadonlySet<string> =
		name === "text:p" || name === "text:h" || name === "text:span"
			? INLINE
			: name === "table:table" || TABLE.has(name)
				? name === "table:table-cell" || name === "table:covered-table-cell"
					? TOP
					: TABLE
				: INLINE

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		if (!inside.has(elementName(document, namespaces, child)) || !plain(document, namespaces, child)) return false
	}

	return true
}

// The parts of content.xml a run of its body's blocks is read within.
export interface SpanFrame {
	// The root's start tag, office:body's and office:text's (as content.xml has them), and their end tags.
	readonly open: Uint8Array
	readonly close: Uint8Array
}

// A run of top-level blocks parsed within `frame`: the document and office:text. `strict`: checked for every
// well-formedness rule (assertWellFormed), not only its structure.
function parseRun(frame: SpanFrame, bytes: Uint8Array, strict: boolean): { document: XmlDocument; namespaces: Namespaces; text: number } {
	const wrapped = new Uint8Array(frame.open.length + bytes.length + frame.close.length)

	wrapped.set(frame.open)
	wrapped.set(bytes, frame.open.length)
	wrapped.set(frame.close, frame.open.length + bytes.length)

	const document = parseXml(wrapped)

	if (strict) assertWellFormed(document)

	const body = document.firstChildElement(document.documentElement())

	return { document, namespaces: new Namespaces(document), text: document.firstChildElement(body) }
}

export interface SpanStyles {
	readonly odf: OdfPackage
	readonly sheet: OdfStyleSheet
	readonly settings: TranslatorSettings
}

// A run of top-level blocks translated on its own, from state `state` (Translator.blockState): its WordprocessingML
// blocks and the translation's styles part; undefined when the run holds what a window may not, or leaves state.
// `strict`: a written run, checked for every well-formedness rule.
export function translateRun(
	styles: SpanStyles,
	frame: SpanFrame,
	bytes: Uint8Array,
	state: string,
	strict = false
): { readonly blocks: Uint8Array; readonly styles: string } | undefined {
	const { document, namespaces, text } = parseRun(frame, bytes, strict)

	for (let node = document.firstChildElement(text); node >= 0; node = document.nextSiblingElement(node)) {
		if (!TOP.has(elementName(document, namespaces, node)) || !plain(document, namespaces, node)) return undefined
	}

	const translator = new Translator(styles.odf, styles.sheet, new OdfStyleResolver(styles.sheet), styles.settings)
	const body = new BodyFlow()

	translator.pages.body = body
	translator.document = document
	translator.namespaces = namespaces
	translator.scope = "content"

	if (translator.blockState() !== state) translator.pages.resume(state)

	for (let node = document.firstChildElement(text); node >= 0; node = document.nextSiblingElement(node)) {
		translator.block(node, body)
		translator.takeLossy()
	}

	if (!translator.leftAsIn(state)) return undefined

	body.finish("")

	return {
		blocks: body.writer.bytes(),
		styles: `${XML_DECLARATION}<w:styles ${PART_NAMESPACES}>${translator.styles.stylesXml(translator)}</w:styles>`
	}
}

// The meanings of a run's WordprocessingML blocks (no numbering, notes, comments, pictures or links: a window
// holds none), `styles` the translation's styles part.
export function runMeaning(blocks: Uint8Array, styles: XmlDocument, compact: (text: string) => string): readonly string[] | undefined {
	const document = parseXml(
		encodeUtf8(`${XML_DECLARATION}<w:document ${PART_NAMESPACES}><w:body>${decodeUtf8(blocks)}</w:body></w:document>`)
	)
	const body = document.firstChildElement(document.documentElement())
	const meaning = new Meaning({
		styles,
		numbering: undefined,
		footnotes: undefined,
		endnotes: undefined,
		comments: undefined,
		target: id => `?${id}`,
		compact
	})
	const read = meaning.body(document, body)

	return read.sections.length > 0 ? undefined : read.blocks
}
