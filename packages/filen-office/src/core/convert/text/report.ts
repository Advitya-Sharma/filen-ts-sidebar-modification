import type { DocEncryptionScheme } from "../../doc/errors"
import type { DocNotConvertedKind } from "../../doc/toDocx/report"
import type { OdtLossyReason } from "../../odt/types"
import type { RtfPreservedKind } from "../../rtf/reader"

export type TextSourceFormat = "docx" | "odt" | "rtf" | "doc"
export type TextTargetFormat = "docx" | "odt"

// What an RTF file holds that its translation keeps in the source bytes only (RtfPreservedKind), and code pages the
// platform could not decode. A text box's (rtfTextbox) and WordArt's text and a group's text boxes (rtfShapeGroup)
// are written as paragraphs positioned as the shape (w:framePr); the shape itself stays in the source.
export type RtfLossyKind = `rtf${Capitalize<RtfPreservedKind>}` | "rtfCodepage" | "rtfTruncatedBinary"

// What ODF 1.3 cannot hold as WordprocessingML has it, or the converter writes another way. Each counts occurrences.
export type DocxToOdtLossyKind =
	// A VBA project (.docm, .dotm): an OpenDocument text holds none of this kind.
	| "macros"
	// A template's building blocks (the glossary document).
	| "glossary"
	// Custom XML parts (and the content controls' data bindings to them).
	| "customXml"
	// Fonts embedded in the package.
	| "embeddedFont"
	// Theme fonts, written as the families they name.
	| "theme"
	// A paragraph or run property without an ODF form, left out; a table, row or cell property likewise.
	| "paragraphFormat"
	| "runFormat"
	// East Asian line breaking or spacing turned off (kinsoku, auto space, word wrap...), which the writer gives no form.
	| "eastAsianTypography"
	| "tableLayout"
	// A border style the writer has no form for (double, thick, wave, art...): drawn as a single line.
	| "borderStyle"
	// A table style's conditional formatting (first row, banding...): ODF table templates have no corner cells or
	// band sizes; the table keeps its own formatting.
	| "tableStyle"
	// A block the emitter could not write as it was: written with its text and paragraph style only.
	| "blockFormat"
	// A shape, group, chart or diagram (not drawn; any text it holds is written after its paragraph), or a text box
	// whose text could not stay in its frame.
	| "drawing"
	// A complex or simple field other than PAGE, NUMPAGES and HYPERLINK, kept as its shown result.
	| "field"
	// Office Math, kept as its text.
	| "math"
	// A content control, kept as its content.
	| "contentControl"
	// A tracked change other than inserted or deleted text, or one in a header or footer (shown as accepted).
	| "trackedChange"
	// A comment's replies or done state (ODF 1.3 annotations record neither).
	| "commentThread"
	// An imported chunk (altChunk) or subdocument, which the package holds as another file.
	| "altChunk"
	| "subDocument"
	// Section page setup ODF page styles cannot express (unequal columns, line numbers, page borders...), or a
	// section break where ODF has none (in a table cell, header or note).
	| "sectionLayout"
	// A style property without an ODF form: the style keeps the rest.
	| "styleDefinition"
	// A list level ODF list styles cannot express, written as the nearest form.
	| "listFormat"
	// A note referenced from a note, comment, header or footer (ODF notes hold no notes; LibreOffice opens none in a
	// header or footer): its text written where it is referenced, the notes it references in turn (past 16 deep, left out
	// and counted once more).
	| "nestedNote"
	// Footnote or endnote numbering or placement ODF's notes configuration has no form for (a number format other than
	// decimal, letters or roman numerals, numbering restarting per section, endnotes ending each section, settings
	// differing by section): the nearest written.
	| "noteSettings"
	// An editing restriction, or part of one, an OpenDocument text cannot hold: comments-only editing, the lock on
	// change tracking, a restriction's password, a range left editable.
	| "documentProtection"
	// A link whose target cannot be written as a URI (RFC 3986): kept as its text.
	| "linkTarget"

// What a standalone .docx of an ODT translation leaves out beyond the translation's own approximations.
export type OdtToDocxLossyKind =
	// style:flow-with-text false (ODF 1.3 Part 3 §20.267: anchored frames kept on their anchor's page), which
	// WordprocessingML has no setting for.
	| "floatsOnAnchorPage"
	// A comment's range marks inside another comment's text: the comment is anchored where its range starts.
	| "commentRange"
	// A line holding only as-char objects, which LibreOffice makes as tall as they are (Word adds the paragraph
	// mark's font): the translation's object-line-height compatibility setting, which w:compatSetting cannot hold.
	| "objectLineHeight"
	// Paragraph borders and padding as LibreOffice lays them out (the translation's paragraph-padding setting): a start
	// or end border inside the indents (Word draws it outside them), padding beside no line (Word leaves it out),
	// neighbours' borders joined, and a bordered paragraph kept apart from the next by style:join-border false.
	| "paragraphPadding"

export type TextLossyKind = OdtLossyReason | DocNotConvertedKind | RtfLossyKind | OdtToDocxLossyKind | DocxToOdtLossyKind

export interface TextLossyFeature {
	readonly kind: TextLossyKind
	readonly count: number
}

export interface TextLossyReport {
	readonly from: TextSourceFormat
	readonly to: TextTargetFormat
	// One entry per kind, counted, sorted by kind; empty when nothing was approximated or left out.
	readonly features: readonly TextLossyFeature[]
}

export type TextRefusalReason =
	// `from` → `to` is not a conversion.
	| "unsupported"
	// Not a file of the source format.
	| "notSource"
	// A password-protected source (an encrypted .doc without its password; an encrypted package or ODF file).
	| "encrypted"
	| "wrongPassword"
	// Word 6/95 or earlier.
	| "legacyVersion"
	| "corrupt"
	| "tooLarge"
	| "aborted"

export interface TextConversionRefusal {
	readonly reason: TextRefusalReason
	// A diagnostic in English, for logs.
	readonly message: string
	// "encrypted" of a Word 97–2003 document: its cipher ("unknown": one no password opens here).
	readonly encryption?: DocEncryptionScheme
}

export type TextConversion =
	| { readonly ok: true; readonly bytes: Uint8Array; readonly lossy: TextLossyReport }
	| { readonly ok: false; readonly refusal: TextConversionRefusal }

export class LossyCounts {
	private readonly counts = new Map<TextLossyKind, number>()

	add(kind: TextLossyKind, count = 1): void {
		if (count > 0) this.counts.set(kind, (this.counts.get(kind) ?? 0) + count)
	}

	report(from: TextSourceFormat, to: TextTargetFormat): TextLossyReport {
		return {
			from,
			to,
			features: [...this.counts].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)).map(([kind, count]) => ({ kind, count }))
		}
	}
}

export class TextConversionError extends Error {
	readonly reason: TextRefusalReason
	readonly encryption: DocEncryptionScheme | undefined

	constructor(reason: TextRefusalReason, message: string, encryption?: DocEncryptionScheme) {
		super(message)
		this.name = "TextConversionError"
		this.reason = reason
		this.encryption = encryption
	}

	get refusal(): TextConversionRefusal {
		return { reason: this.reason, message: this.message, ...(this.encryption === undefined ? {} : { encryption: this.encryption }) }
	}
}
