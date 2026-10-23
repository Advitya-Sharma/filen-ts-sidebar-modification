import type { OdfFacts, OdfEncryption, OdfOpenFailure } from "../odf/index"
import type { LinkedPart } from "../opc/package"

// What a translated block approximates or leaves out. A block with any of these is shown as LibreOffice would
// show it as closely as WordprocessingML allows, and must not be re-serialized from the translation on save.
export type OdtLossyReason =
	// style:next-style-name chains beyond a first page (A → B → C), or a first page whose layout differs.
	| "masterPageChain"
	// Headers or footers of a fixed height (svg:height): WordprocessingML headers always grow.
	| "fixedHeaderHeight"
	// style:line-spacing (leading added to each line), approximated as a multiple.
	| "lineSpacingLeading"
	// fo:text-transform lowercase or capitalize.
	| "textTransform"
	// Text rotation, combining, ruby or a character effect without an equivalent.
	| "textEffect"
	// draw:chain-next-name: text flowing from one text box into another.
	| "chainedTextBox"
	// Frames anchored to another frame, or to a page other than the first, or that flow with the text
	// (style:flow-with-text true: kept on their anchor's page, not inside its text area).
	| "frameAnchor"
	// A chart, formula or embedded object shown from its replacement image or as a placeholder.
	| "embeddedObject"
	// An image whose format the viewer cannot draw, shown as a placeholder.
	| "unsupportedImage"
	// Cropping (fo:clip) or image adjustments (luminance, contrast, gamma, colour mode).
	| "imageAdjustment"
	// An index (table of contents, alphabetical, bibliography...) kept as the entries it shows: its definition
	// (the *-source element) and its link to what it indexes are left out.
	| "index"
	// A bookmark between blocks, left out (WordprocessingML's would have to move into a paragraph).
	| "bookmark"
	// Shapes drawn through their geometry: connectors, measures, groups flattened, unknown preset types.
	| "shapeApproximated"
	// Fields with no WordprocessingML equivalent (variables, sequences, conditional text, database fields):
	// their shown result is kept as text.
	| "fieldAsText"
	// text:section protection, unbalanced columns, or section-level note collection.
	| "sectionProperties"
	// More than nine list levels, or a list label WordprocessingML cannot express.
	| "listLevel"
	// Tables wider than 63 columns, repeated rows cut short, or a table that may not break between rows.
	| "tableLimit"
	// Tracked changes beyond inserted and deleted text (format changes, deleted paragraphs).
	| "trackedChange"
	// An element the translator does not know; its text content is kept.
	| "unknownElement"

export interface OdtOrigin {
	// The ODF stream the block was read from, and its byte span [start, end) there.
	readonly stream: "content.xml" | "styles.xml"
	readonly start: number
	readonly end: number
	// The WordprocessingML part written for it, and the range [first, first + count) of top-level blocks
	// (w:body, w:hdr or w:ftr children) it became; count 0 for a block that wrote nothing there (tracked-change
	// records, declarations).
	readonly part: string
	readonly first: number
	readonly count: number
	readonly lossy: readonly OdtLossyReason[]
}

export type OdtOpenFailure = OdfOpenFailure | "malformed"

export interface OdtTranslation {
	// The WordprocessingML package's parts by zip entry name; pictures stay in the ODF archive until read.
	readonly parts: ReadonlyMap<string, Uint8Array | LinkedPart>
	readonly origins: readonly OdtOrigin[]
	// The origins' byte offsets are into content.xml as stored (false: as transcoded from UTF-16).
	readonly originsExact: boolean
	readonly facts: OdfFacts
	// Reasons any block was lossy, with how many blocks each; a reason only the document's styles or note
	// settings have counts once.
	readonly lossy: ReadonlyMap<OdtLossyReason, number>
	// Embedded objects shown from a replacement or as a placeholder, by kind.
	readonly objects: ReadonlyMap<string, number>
}

export type OdtTranslateResult =
	| { readonly ok: true; readonly translation: OdtTranslation }
	| { readonly ok: false; readonly reason: OdtOpenFailure; readonly message: string; readonly encryption?: OdfEncryption }
