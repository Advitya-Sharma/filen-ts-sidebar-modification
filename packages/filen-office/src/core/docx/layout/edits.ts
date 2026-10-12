// Edits a layout shows before they are saved: blocks laid out from overlay parts (an edited block written
// as it will be saved and parsed again) in place of their source block, and the body's block order once
// blocks were added, removed or moved. A ref packs a part index and an item of that part.

export const REF_SPAN = 0x4000000

export function ref(part: number, item: number): number {
	return part * REF_SPAN + item
}

export function refPart(value: number): number {
	return Math.floor(value / REF_SPAN)
}

export function refItem(value: number): number {
	return value - Math.floor(value / REF_SPAN) * REF_SPAN
}

// The body's blocks (paragraphs and tables, wrappers entered) in order, the sections one after another.
export interface BodyBlocks {
	readonly refs: Float64Array
	readonly count: number
	// Per section, where its blocks end in refs (exclusive).
	readonly sectionEnds: Int32Array
}

export interface LayoutEdits {
	// Source block ref -> the ref laid out instead (paragraphs, and tables whose structure changed).
	readonly substitutes: ReadonlyMap<number, number>
	// Undefined while the body's block structure is the model's.
	readonly body: BodyBlocks | undefined
	// Other stories (headers, footers, notes) whose block structure changed: their blocks, by story index.
	readonly stories?: ReadonlyMap<number, Float64Array>
	// Whether a part's paragraphs are laid out by the body flow (the body part and overlays of its blocks):
	// what the flow counted for them is forgotten when layout restarts.
	bodyPart(part: number): boolean
	// The key what layout counts for a paragraph (list markers, note numbers) is kept under: an edited
	// paragraph keeps its source's, whichever overlay it is laid out from now.
	memoKey?(part: number, item: number): number
	// The source block (a ref) an edited block holding item `item` of overlay part `part` stands for, or -1.
	sourceBlock?(part: number, item: number): number
	// The document-wide switch for even pages' headers and footers as edited (undefined: the settings').
	readonly evenAndOddHeaders?: boolean | undefined
}
