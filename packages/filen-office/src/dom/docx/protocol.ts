import type { DocxGateReason, DocxOpenFailure, DocxVariant } from "../../core/docx/index"
import type { DecorationValue } from "../../core/docx/layout/index"
import type { DrawingClip, DrawingImage, DrawingText, Paint, Shadow, Stroke } from "../../core/drawingml/index"
import type { FontFaceDescriptor } from "../../core/fonts/index"
import type { MathElement } from "../../core/math/index"
import type { DocxAnnotations } from "./annotations"

// What the DOCX engine sends to a renderer. Everything here is plain data (numbers, enums, strings and
// typed arrays) so it crosses a worker boundary by structured clone, with the typed arrays transferred. A
// renderer still validates what it turns into DOM: colours, sizes and tree shapes are checked again on
// arrival.

// The engine reports an encrypted package by why it cannot open it, never as "encrypted" (that stays for a
// host's own checks). Unlocking (encrypted.ts) comes before the engine is given the decrypted package.
//   needsPassword          password encryption, opened once given the password
//   rightsManaged          IRM / a sensitivity label
//   unsupportedEncryption  an encryption scheme Filen does not implement
//   damagedEncryption      a malformed encrypted container
export type DocxErrorReason =
	DocxOpenFailure | "needsPassword" | "rightsManaged" | "unsupportedEncryption" | "damagedEncryption" | "malformed" | "failed"

export interface DocxOpenInfo {
	readonly variant: DocxVariant
	// Why the document would open view-only (gate 0); informational for the viewer.
	readonly gate: readonly DocxGateReason[]
	// Pages expected before layout knows: Word's recorded page breaks + 1.
	readonly estimatedPages: number
	// The first section's page size (px), for placeholders of pages not laid out yet.
	readonly pageWidth: number
	readonly pageHeight: number
	// Uncompressed size of the main document part, from the zip directory (a host can tell how heavy any
	// other renderer would find the file).
	readonly mainPartBytes: number
	// How many comments and tracked changes the document holds (an app offers its markup controls before
	// the annotations themselves are ready).
	readonly comments?: number
	readonly revisions?: number
	// Of the tracked changes, those to formatting (run, paragraph, numbering, table properties).
	readonly formatting?: number
}

// One text style of the layout (the fragment STYLE ints index these).
export interface DocxTextStyle {
	// CSS font shorthand naming an internal alias (never a document's family name).
	readonly font: string
	// The font for the lower-case letters of a small-caps fragment.
	readonly smallCapsFont: string
	// Six hex digits, or "" for the document's automatic colour.
	readonly color: string
	readonly kerning: boolean
	// px added after every character.
	readonly letterSpacing: number
	// Horizontal glyph scale.
	readonly scaleX: number
	// Baseline raised (positive) or lowered.
	readonly shift: number
	// The drawn font's ascent and descent as the browser lays it out (px), to put the baseline exactly.
	readonly ascent: number
	readonly descent: number
	// TrackedLook of the run (0 none).
	readonly revision: number
	// The colour slot (1..8) of the tracked change's author, 0 without one: inserted and deleted text is drawn
	// in --fo-revision-<slot>.
	readonly author?: number
}

export type DocxLinkTarget =
	| { readonly kind: "external"; readonly href: string; readonly tooltip: string | undefined }
	| { readonly kind: "internal"; readonly anchor: string; readonly tooltip: string | undefined }

// linkInts per range: source, text start, text end, link (index into the session's link table).
export const LINK_INTS = 4

export interface DocxPageData {
	readonly index: number
	readonly number: number
	readonly numberText: string
	readonly section: number
	readonly width: number
	readonly height: number
	readonly lineCount: number
	readonly fragmentCount: number
	readonly decorationCount: number
	readonly lineFloats: Float32Array
	readonly lineInts: Int32Array
	readonly fragmentFloats: Float32Array
	readonly fragmentInts: Int32Array
	readonly decorationFloats: Float32Array
	readonly decorationInts: Int32Array
	readonly graphicCount: number
	readonly graphicFloats: Float32Array
	readonly graphicInts: Int32Array
	readonly objectCount: number
	readonly objectFloats: Float32Array
	readonly objectInts: Int32Array
	// ObjectKind of each object entry.
	readonly objectKinds: Uint8Array
	readonly sourceCount: number
	readonly sourceInts: Int32Array
	readonly sourceTexts: readonly string[]
	readonly linkCount: number
	readonly linkInts: Int32Array
}

// A drawing's display list as data (DrawingList.compact() without its interning maps).
export interface DrawingData {
	// The frame the drawing was built for (px); the renderer scales it onto the object's box.
	readonly width: number
	readonly height: number
	readonly commands: Uint8Array
	readonly coords: Float32Array
	readonly itemCount: number
	readonly itemFloats: Float32Array
	readonly itemInts: Int32Array
	readonly paints: readonly Paint[]
	readonly strokes: readonly Stroke[]
	readonly images: readonly DrawingImage[]
	readonly texts: readonly DrawingText[]
	readonly shadows: readonly Shadow[]
	readonly clips?: readonly DrawingClip[]
	// Per item: its clip, or -1; absent when no item is clipped.
	readonly itemClips?: Int32Array
	// Per text: the font family list the engine resolved for it (an internal alias and a generic family),
	// or "" for none.
	readonly textFonts?: readonly string[]
	// The surface the drawing starts on (0xRRGGBBAA), absent for a transparent one.
	readonly background?: number
}

export type DocxObjectData =
	// alt: the picture's accessible name (its alternative text, else its title), absent for none.
	| { readonly kind: "drawing"; readonly id: number; readonly drawing: DrawingData; readonly alt?: string }
	// fontSize: px of the run holding the equation.
	| { readonly kind: "math"; readonly id: number; readonly tree: MathElement; readonly fontSize: number }
	| { readonly kind: "separator"; readonly id: number }
	| { readonly kind: "checkbox"; readonly id: number; readonly checked: boolean }
	| { readonly kind: "placeholder"; readonly id: number; readonly reason: string }

export interface DocxMedia {
	readonly bytes: Uint8Array
	// The image MIME type a renderer may show it as, or "" when it cannot be shown natively.
	readonly type: string
}

export interface DocxLocation {
	readonly page: number
	// px from the page's top.
	readonly y: number
}

// A face to register in the renderer's realm under the same alias the layout measured with.
export type DocxFontFace =
	| { readonly kind: "bytes"; readonly face: FontFaceDescriptor; readonly bytes: Uint8Array }
	| { readonly kind: "local"; readonly face: FontFaceDescriptor; readonly name: string }

export type DocxEvent =
	| { readonly type: "opened"; readonly info: DocxOpenInfo }
	| { readonly type: "fonts"; readonly faces: readonly DocxFontFace[] }
	| {
			readonly type: "pages"
			readonly pages: readonly DocxPageData[]
			// Appended to the session's tables, in order (the pages' indices refer to the whole tables).
			readonly styles: readonly DocxTextStyle[]
			readonly decorations: readonly DecorationValue[]
			readonly links: readonly DocxLinkTarget[]
			// Every page is laid out; pageCount is the final count.
			readonly final: boolean
			readonly pageCount: number
	  }
	// mainPartBytes: as in DocxOpenInfo, when the package got that far.
	| { readonly type: "error"; readonly reason: DocxErrorReason; readonly message: string; readonly mainPartBytes?: number }

// The engine's side of an open document, as a renderer calls it (worker or in-thread).
export interface DocxSessionApi {
	objects(ids: readonly number[]): Promise<readonly DocxObjectData[]>
	media(partName: string): Promise<DocxMedia | null>
	locate(anchor: string): Promise<DocxLocation | null>
	// Every match of `query` in the whole document (laid out so far), as FIND_INTS per match.
	find(query: string, options?: DocxFindOptions): Promise<Int32Array>
	// Comments and tracked changes with their boxes on the pages, once every page is laid out (null when
	// the document closed first).
	annotations?(): Promise<DocxAnnotations | null>
	close(): void
}

export interface DocxFindOptions {
	readonly caseSensitive?: boolean
	// Stops after this many matches and one more, which tells that there are more (default DOCX_FIND_LIMIT).
	readonly limit?: number
}

// Matches a find shows at most (a counter shows "10,000+" beyond).
export const DOCX_FIND_LIMIT = 10_000

// Per match: page, source (the page's paragraph), start and end offsets in the source's text.
export const FIND_INTS = 4

export interface DocxOpenOptions {
	// Lay out tracked changes accepted (Final, what Simple Markup shows), rejected (the original text) or
	// inline with insertions and deletions marked (All Markup).
	readonly view?: "final" | "original" | "markup"
	// Where the host serves each bundled font (path under assets/fonts -> same-origin URL).
	readonly fontUrls?: Readonly<Record<string, string>>
	// Keep the model after layout (editing needs it). A viewer lets the engine drop it once every page and
	// object is prepared, keeping only what find, links, bookmarks and images need.
	readonly retain?: boolean
}
