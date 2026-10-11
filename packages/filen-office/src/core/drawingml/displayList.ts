// The drawing layer's output: a small display list a renderer turns into SVG/DOM elements one to one. It
// holds numbers and enums only (plus part names, typefaces and run text, which a renderer passes to
// createElementNS/textContent/the font manager, never into markup or CSS text), so document data cannot
// become markup.
//
// Struct-of-arrays in growable typed arrays, like the text layout's display list; compact() returns tight
// copies to transfer to another thread.
//
// Units: px at 96 per inch (EMU / 9525); angles in degrees, clockwise (y points down). Each item is drawn in
// its own local space and placed by an affine matrix [a b c d e f] (x' = a x + c y + e, y' = b x + d y + f)
// into the drawing's space, whose origin is the top-left of the drawing's frame.

export const PathCommand = { MOVE: 0, LINE: 1, CUBIC: 2, QUAD: 3, CLOSE: 4 } as const

export type PathCommand = (typeof PathCommand)[keyof typeof PathCommand]

// Coordinates (x, y pairs count as two) each command consumes.
export const PATH_COMMAND_COORDS: readonly number[] = [2, 2, 6, 4, 0]

export const ItemKind = {
	// Path commands filled with FILL and stroked with STROKE.
	PATH: 0,
	// Image REF drawn into the local box; clipped by the item's path when it has one.
	IMAGE: 1,
	// Text REF laid out inside the local box (a placeholder: no glyphs are produced here).
	TEXT: 2
} as const

export type ItemKind = (typeof ItemKind)[keyof typeof ItemKind]

export const ITEM_FLOATS = 10
export const ITEM_INTS = 10

// itemFloats: matrix a, b, c, d, e, f; local box x, y, width, height (the shape's frame for paths, so
// gradients and picture fills map onto it; the destination for images; the text rectangle for text).
export const ItemFloat = { A: 0, B: 1, C: 2, D: 3, E: 4, F: 5, X: 6, Y: 7, WIDTH: 8, HEIGHT: 9 } as const

// itemInts: kind, first path command, command count, first coordinate, fill paint (-1 none), stroke
// (-1 none), ref (image or text index, -1), shadow (-1 none), flags, source (the caller's id).
export const ItemInt = {
	KIND: 0,
	FIRST_COMMAND: 1,
	COMMAND_COUNT: 2,
	FIRST_COORD: 3,
	FILL: 4,
	STROKE: 5,
	REF: 6,
	SHADOW: 7,
	FLAGS: 8,
	SOURCE: 9
} as const

export const ItemFlag = {
	// Bits 0-2: the path fill modifier applied to the fill colour (ST_PathFillMode).
	LIGHTEN: 1,
	LIGHTEN_LESS: 2,
	DARKEN: 3,
	DARKEN_LESS: 4,
	// Fill with the even-odd rule (DrawingML and VML fill alternately).
	EVEN_ODD: 8,
	// The item is hidden (kept for hit testing and alt text).
	HIDDEN: 16,
	// An arrowhead or other decoration generated for the preceding item.
	DECORATION: 32
} as const

export const ITEM_FILL_MODIFIER_MASK = 7

// Colours are 0xRRGGBBAA, unsigned.
export type Rgba = number

export function rgba(hex: string, alpha: number = 1): Rgba {
	const value = Number.parseInt(hex, 16)
	const a = Math.round(Math.min(1, Math.max(0, alpha)) * 255)

	return ((Number.isFinite(value) ? value & 0xffffff : 0) * 256 + a) >>> 0
}

export function rgbaHex(color: Rgba): string {
	return (color >>> 8).toString(16).padStart(6, "0").toUpperCase()
}

export function rgbaAlpha(color: Rgba): number {
	return (color & 255) / 255
}

export interface GradientStops {
	// Positions 0..1, ascending.
	readonly offsets: Float32Array
	readonly colors: Uint32Array
}

// ST_PresetPatternVal (§20.1.10.51).
export type PatternPreset =
	| "pct5"
	| "pct10"
	| "pct20"
	| "pct25"
	| "pct30"
	| "pct40"
	| "pct50"
	| "pct60"
	| "pct70"
	| "pct75"
	| "pct80"
	| "pct90"
	| "horz"
	| "vert"
	| "ltHorz"
	| "ltVert"
	| "dkHorz"
	| "dkVert"
	| "narHorz"
	| "narVert"
	| "dashHorz"
	| "dashVert"
	| "cross"
	| "dnDiag"
	| "upDiag"
	| "ltDnDiag"
	| "ltUpDiag"
	| "dkDnDiag"
	| "dkUpDiag"
	| "wdDnDiag"
	| "wdUpDiag"
	| "dashDnDiag"
	| "dashUpDiag"
	| "diagCross"
	| "smCheck"
	| "lgCheck"
	| "smGrid"
	| "lgGrid"
	| "dotGrid"
	| "smConfetti"
	| "lgConfetti"
	| "horzBrick"
	| "diagBrick"
	| "solidDmnd"
	| "openDmnd"
	| "dotDmnd"
	| "plaid"
	| "sphere"
	| "weave"
	| "divot"
	| "shingle"
	| "wave"
	| "trellis"
	| "zigZag"

export interface Inset {
	readonly left: number
	readonly top: number
	readonly right: number
	readonly bottom: number
}

export type Paint =
	| { readonly kind: "solid"; readonly color: Rgba }
	// Angle in degrees from the +x axis, clockwise, in the item's local box; `scaled` stretches the
	// angle with the box's aspect ratio (a:lin@scaled).
	| { readonly kind: "linear"; readonly stops: GradientStops; readonly angle: number; readonly scaled: boolean }
	// a:path gradients: `focus` is the fillToRect (fractions of the box inset from each side) the
	// first stop fills; the last stop reaches the box's corners ("circle"/"shape") or sides ("rect").
	| { readonly kind: "path"; readonly shape: "circle" | "rect" | "shape"; readonly stops: GradientStops; readonly focus: Inset }
	| { readonly kind: "pattern"; readonly pattern: PatternPreset; readonly foreground: Rgba; readonly background: Rgba }
	// A picture fill: `stretch` insets the image box from the item box (fractions, a:fillRect); `tile`
	// places it at its own size scaled by sx/sy, offset by tx/ty px, mirrored per flip, from `align`.
	| {
			readonly kind: "image"
			readonly image: number
			readonly stretch: Inset | undefined
			readonly tile: ImageTile | undefined
	  }

export interface ImageTile {
	readonly tx: number
	readonly ty: number
	readonly sx: number
	readonly sy: number
	readonly flip: "none" | "x" | "y" | "xy"
	readonly align: "tl" | "t" | "tr" | "l" | "ctr" | "r" | "bl" | "b" | "br"
}

export type DashPreset =
	"solid" | "dot" | "dash" | "lgDash" | "dashDot" | "lgDashDot" | "lgDashDotDot" | "sysDash" | "sysDot" | "sysDashDot" | "sysDashDotDot"

export interface Stroke {
	readonly width: number
	readonly paint: number
	readonly dash: DashPreset
	// a:custDash as dash, gap pairs in multiples of the width; null for the preset.
	readonly customDash: Float32Array | null
	readonly cap: "flat" | "round" | "square"
	readonly join: "round" | "bevel" | "miter"
	readonly miterLimit: number
	readonly compound: "sng" | "dbl" | "thickThin" | "thinThick" | "tri"
	// Insets the stroke inside closed paths (a:ln@algn="in").
	readonly inset: boolean
}

export type ImageFormat = "png" | "jpeg" | "gif" | "bmp" | "webp" | "svg" | "tiff" | "emf" | "wmf" | "unknown"

// How a renderer may show an image: as is, after sanitising (SVG), after converting (TIFF), through the
// metafile converter (EMF/WMF), or not at all (external, missing, unknown).
export type ImageHandling = "native" | "sanitize" | "convert" | "metafile" | "placeholder"

export interface DrawingImage {
	// Part name of the image in the package; undefined for external or missing targets (never fetched).
	readonly partName: string | undefined
	readonly format: ImageFormat
	readonly handling: ImageHandling
	// An SVG rendition (asvg:svgBlip); `partName` is then its raster fallback.
	readonly svgPartName: string | undefined
	// a:srcRect: fractions cropped from each side of the source (negative pads).
	readonly crop: Inset
	readonly alpha: number
	readonly grayscale: boolean
	// a:biLevel threshold 0..1, or -1.
	readonly biLevel: number
	// a:duotone colours (dark, light), or null.
	readonly duotone: readonly [Rgba, Rgba] | null
	// a:lum brightness and contrast, -1..1.
	readonly brightness: number
	readonly contrast: number
	// a:clrChange: pixels of colour `from` become `to` (alpha included), or null.
	readonly colorChange: readonly [Rgba, Rgba] | null
}

export type TextAnchor = "t" | "ctr" | "b" | "just" | "dist"

export type TextVertical = "horz" | "vert" | "vert270" | "wordArtVert" | "eaVert" | "mongolianVert" | "wordArtVertRtl"

// a:bodyPr, as input for text layout.
export interface TextBodyProps {
	readonly insets: Inset
	readonly anchor: TextAnchor
	readonly anchorCenter: boolean
	readonly vertical: TextVertical
	readonly wrap: boolean
	// Extra rotation of the text in degrees.
	readonly rotation: number
	readonly upright: boolean
	readonly columns: number
	readonly columnSpacing: number
	readonly autofit: "none" | "shape" | "normal"
	// a:normAutofit fontScale and lnSpcReduction (fractions).
	readonly fontScale: number
	readonly lineSpacingReduction: number
}

export type TextAlign = "l" | "ctr" | "r" | "just" | "dist" | "justLow" | "thaiDist"

// ST_TextUnderlineType (§20.1.10.82).
export const UNDERLINES = [
	"none",
	"words",
	"sng",
	"dbl",
	"heavy",
	"dotted",
	"dottedHeavy",
	"dash",
	"dashHeavy",
	"dashLong",
	"dashLongHeavy",
	"dotDash",
	"dotDashHeavy",
	"dotDotDash",
	"dotDotDashHeavy",
	"wavy",
	"wavyHeavy",
	"wavyDbl"
] as const

export type TextUnderline = (typeof UNDERLINES)[number]

export interface TextRun {
	readonly text: string
	// Points.
	readonly size: number
	readonly bold: boolean
	readonly italic: boolean
	readonly underline: TextUnderline
	readonly strike: "noStrike" | "sngStrike" | "dblStrike"
	readonly color: Rgba
	// Typeface as written in the document, theme references resolved; empty when unspecified.
	readonly font: string
	// Superscript/subscript as a fraction of the size (positive raises).
	readonly baseline: number
	readonly caps: "none" | "small" | "all"
	// Letter spacing in points.
	readonly spacing: number
	// A line break run (a:br).
	readonly lineBreak: boolean
}

export interface TextParagraph {
	readonly align: TextAlign
	readonly level: number
	// px.
	readonly marginLeft: number
	readonly indent: number
	// Line spacing as a fraction of single (> 0) or exact points (< 0, negated); 0 for the default.
	readonly lineSpacing: number
	// Points.
	readonly spaceBefore: number
	readonly spaceAfter: number
	readonly rtl: boolean
	// The bullet character with its font, or empty.
	readonly bullet: string
	readonly bulletFont: string
	readonly runs: readonly TextRun[]
	// The size of the end-of-paragraph mark (for empty paragraphs), points.
	readonly endSize: number
}

export type DrawingText =
	// Rich text the drawing carries itself (DrawingML a:txBody, chart text).
	| { readonly kind: "body"; readonly body: TextBodyProps; readonly paragraphs: readonly TextParagraph[] }
	// A story the host lays out (a WordprocessingML text box): `story` is the caller's reference.
	| { readonly kind: "story"; readonly body: TextBodyProps; readonly story: number }

// a:outerShdw (only outer shadows are drawn).
export interface Shadow {
	readonly color: Rgba
	readonly blur: number
	readonly dx: number
	readonly dy: number
	readonly sx: number
	readonly sy: number
}

export const IDENTITY: readonly number[] = [1, 0, 0, 1, 0, 0]

// Largest coordinate stored (px); beyond it, and for non-finite values, a coordinate is clamped (NaN reads
// as 0), so extreme or overflowing input never reaches the Float32 arrays as Infinity or NaN.
export const MAX_PX = 2 ** 31

function safe(value: number): number {
	return value > MAX_PX ? MAX_PX : value < -MAX_PX ? -MAX_PX : value === value ? value : 0
}

function grownFloats(array: Float32Array, needed: number): Float32Array {
	if (needed <= array.length) return array

	const grown = new Float32Array(Math.max(needed, array.length * 2))

	grown.set(array)

	return grown
}

function grownInts(array: Int32Array, needed: number): Int32Array {
	if (needed <= array.length) return array

	const grown = new Int32Array(Math.max(needed, array.length * 2))

	grown.set(array)

	return grown
}

function grownBytes(array: Uint8Array, needed: number): Uint8Array {
	if (needed <= array.length) return array

	const grown = new Uint8Array(Math.max(needed, array.length * 2))

	grown.set(array)

	return grown
}

function insetKey(inset: Inset | undefined): string {
	return inset === undefined ? "" : `${inset.left},${inset.top},${inset.right},${inset.bottom}`
}

function stopsKey(stops: GradientStops): string {
	let key = ""

	for (let k = 0; k < stops.offsets.length; k++) key += `${stops.offsets[k]}:${stops.colors[k]};`

	return key
}

function paintKey(paint: Paint): string {
	switch (paint.kind) {
		case "solid":
			return `s${paint.color}`
		case "linear":
			return `l${paint.angle},${paint.scaled ? 1 : 0},${stopsKey(paint.stops)}`
		case "path":
			return `p${paint.shape},${insetKey(paint.focus)},${stopsKey(paint.stops)}`
		case "pattern":
			return `t${paint.pattern},${paint.foreground},${paint.background}`
		case "image": {
			const tile = paint.tile

			return `i${paint.image},${insetKey(paint.stretch)},${tile === undefined ? "" : `${tile.tx},${tile.ty},${tile.sx},${tile.sy},${tile.flip},${tile.align}`}`
		}
	}
}

function strokeKey(stroke: Stroke): string {
	return `${stroke.width},${stroke.paint},${stroke.dash},${stroke.customDash?.join(",") ?? ""},${stroke.cap},${stroke.join},${stroke.miterLimit},${stroke.compound},${stroke.inset ? 1 : 0}`
}

// A drawing's display list. Paints and strokes are interned, so a chart with a thousand bars of one series
// shares one paint.
export class DrawingList {
	commandCount = 0
	coordCount = 0
	itemCount = 0
	commands: Uint8Array = new Uint8Array(64)
	coords: Float32Array = new Float32Array(128)
	itemFloats: Float32Array = new Float32Array(ITEM_FLOATS * 8)
	itemInts: Int32Array = new Int32Array(ITEM_INTS * 8)
	// Per item: min x, min y, max x, max y in drawing space (strokes included).
	itemBounds: Float32Array = new Float32Array(4 * 8)
	readonly paints: Paint[] = []
	readonly strokes: Stroke[] = []
	readonly images: DrawingImage[] = []
	readonly texts: DrawingText[] = []
	readonly shadows: Shadow[] = []
	// The frame the drawing was built for (px); content may overflow it (see bounds()).
	width = 0
	height = 0
	private readonly paintIndex = new Map<string, number>()
	private readonly strokeIndex = new Map<string, number>()

	paint(paint: Paint): number {
		const key = paintKey(paint)
		let index = this.paintIndex.get(key)

		if (index === undefined) {
			index = this.paints.length
			this.paints.push(paint)
			this.paintIndex.set(key, index)
		}

		return index
	}

	stroke(stroke: Stroke): number {
		const key = strokeKey(stroke)
		let index = this.strokeIndex.get(key)

		if (index === undefined) {
			index = this.strokes.length
			this.strokes.push(stroke)
			this.strokeIndex.set(key, index)
		}

		return index
	}

	image(image: DrawingImage): number {
		this.images.push(image)

		return this.images.length - 1
	}

	text(text: DrawingText): number {
		this.texts.push(text)

		return this.texts.length - 1
	}

	shadow(shadow: Shadow): number {
		this.shadows.push(shadow)

		return this.shadows.length - 1
	}

	private command(code: PathCommand, coords: number): void {
		this.commands = grownBytes(this.commands, this.commandCount + 1)
		this.commands[this.commandCount++] = code
		this.coords = grownFloats(this.coords, this.coordCount + coords)
	}

	moveTo(x: number, y: number): void {
		this.command(PathCommand.MOVE, 2)
		this.coords[this.coordCount++] = safe(x)
		this.coords[this.coordCount++] = safe(y)
	}

	lineTo(x: number, y: number): void {
		this.command(PathCommand.LINE, 2)
		this.coords[this.coordCount++] = safe(x)
		this.coords[this.coordCount++] = safe(y)
	}

	cubicTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): void {
		this.command(PathCommand.CUBIC, 6)

		const coords = this.coords
		let at = this.coordCount

		coords[at++] = safe(x1)
		coords[at++] = safe(y1)
		coords[at++] = safe(x2)
		coords[at++] = safe(y2)
		coords[at++] = safe(x)
		coords[at++] = safe(y)
		this.coordCount = at
	}

	quadTo(x1: number, y1: number, x: number, y: number): void {
		this.command(PathCommand.QUAD, 4)

		const coords = this.coords
		let at = this.coordCount

		coords[at++] = safe(x1)
		coords[at++] = safe(y1)
		coords[at++] = safe(x)
		coords[at++] = safe(y)
		this.coordCount = at
	}

	close(): void {
		this.command(PathCommand.CLOSE, 0)
	}

	// Discards path commands appended since `commandMark`/`coordMark` (an item that turned out empty).
	rewindPath(commandMark: number, coordMark: number): void {
		this.commandCount = commandMark
		this.coordCount = coordMark
	}

	// Adds an item whose path is the commands appended since `firstCommand`/`firstCoord`, or the
	// `commandCount` commands from there when given (a path shared by several items).
	addItem(
		kind: ItemKind,
		matrix: ArrayLike<number>,
		x: number,
		y: number,
		width: number,
		height: number,
		firstCommand: number,
		firstCoord: number,
		fill: number,
		stroke: number,
		ref: number,
		shadow: number,
		flags: number,
		source: number,
		commandCount: number = this.commandCount - firstCommand
	): number {
		const item = this.itemCount

		this.itemFloats = grownFloats(this.itemFloats, (item + 1) * ITEM_FLOATS)
		this.itemInts = grownInts(this.itemInts, (item + 1) * ITEM_INTS)
		this.itemBounds = grownFloats(this.itemBounds, (item + 1) * 4)

		const floats = this.itemFloats
		const ints = this.itemInts
		const f = item * ITEM_FLOATS
		const i = item * ITEM_INTS

		for (let k = 0; k < 6; k++) floats[f + k] = safe(matrix[k] ?? 0)

		floats[f + ItemFloat.X] = safe(x)
		floats[f + ItemFloat.Y] = safe(y)
		floats[f + ItemFloat.WIDTH] = safe(width)
		floats[f + ItemFloat.HEIGHT] = safe(height)
		ints[i + ItemInt.KIND] = kind
		ints[i + ItemInt.FIRST_COMMAND] = firstCommand
		ints[i + ItemInt.COMMAND_COUNT] = commandCount
		ints[i + ItemInt.FIRST_COORD] = firstCoord
		ints[i + ItemInt.FILL] = fill
		ints[i + ItemInt.STROKE] = stroke
		ints[i + ItemInt.REF] = ref
		ints[i + ItemInt.SHADOW] = shadow
		ints[i + ItemInt.FLAGS] = flags
		ints[i + ItemInt.SOURCE] = source
		this.itemCount = item + 1
		this.measureItem(item)

		return item
	}

	kind(item: number): ItemKind {
		return (this.itemInts[item * ITEM_INTS + ItemInt.KIND] ?? 0) as ItemKind
	}

	itemInt(item: number, field: number): number {
		return this.itemInts[item * ITEM_INTS + field] ?? 0
	}

	itemFloat(item: number, field: number): number {
		return this.itemFloats[item * ITEM_FLOATS + field] ?? 0
	}

	// Bounds of every item, [minX, minY, maxX, maxY]; the frame when there are none.
	bounds(): Float32Array {
		const out = new Float32Array([Infinity, Infinity, -Infinity, -Infinity])

		for (let item = 0; item < this.itemCount; item++) {
			const b = item * 4

			out[0] = Math.min(out[0] ?? 0, this.itemBounds[b] ?? 0)
			out[1] = Math.min(out[1] ?? 0, this.itemBounds[b + 1] ?? 0)
			out[2] = Math.max(out[2] ?? 0, this.itemBounds[b + 2] ?? 0)
			out[3] = Math.max(out[3] ?? 0, this.itemBounds[b + 3] ?? 0)
		}

		if (this.itemCount === 0) out.set([0, 0, this.width, this.height])

		return out
	}

	// Tight copies of the arrays, for transfer.
	compact(): DrawingList {
		const out = new DrawingList()

		out.commandCount = this.commandCount
		out.coordCount = this.coordCount
		out.itemCount = this.itemCount
		out.commands = this.commands.slice(0, this.commandCount)
		out.coords = this.coords.slice(0, this.coordCount)
		out.itemFloats = this.itemFloats.slice(0, this.itemCount * ITEM_FLOATS)
		out.itemInts = this.itemInts.slice(0, this.itemCount * ITEM_INTS)
		out.itemBounds = this.itemBounds.slice(0, this.itemCount * 4)
		// Appended one by one: spreading a large table into push() overflows the argument stack.
		for (const paint of this.paints) out.paints.push(paint)
		for (const stroke of this.strokes) out.strokes.push(stroke)
		for (const image of this.images) out.images.push(image)
		for (const text of this.texts) out.texts.push(text)
		for (const shadow of this.shadows) out.shadows.push(shadow)
		out.width = this.width
		out.height = this.height

		return out
	}

	// Drawing-space bounds of one item: the path's extrema (curves solved exactly), widened by half the
	// stroke width; the transformed local box for images and text.
	private measureItem(item: number): void {
		const f = item * ITEM_FLOATS
		const i = item * ITEM_INTS
		const floats = this.itemFloats
		const ints = this.itemInts
		const a = floats[f] ?? 1
		const b = floats[f + 1] ?? 0
		const c = floats[f + 2] ?? 0
		const d = floats[f + 3] ?? 1
		const e = floats[f + 4] ?? 0
		const ff = floats[f + 5] ?? 0
		const box = { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity }
		const add = (x: number, y: number): void => {
			const tx = a * x + c * y + e
			const ty = b * x + d * y + ff

			if (tx < box.minX) box.minX = tx
			if (tx > box.maxX) box.maxX = tx
			if (ty < box.minY) box.minY = ty
			if (ty > box.maxY) box.maxY = ty
		}
		const count = ints[i + ItemInt.COMMAND_COUNT] ?? 0
		const kind = ints[i + ItemInt.KIND] ?? 0

		if (kind === ItemKind.PATH && count > 0) {
			measurePath(this.commands, this.coords, ints[i + ItemInt.FIRST_COMMAND] ?? 0, count, ints[i + ItemInt.FIRST_COORD] ?? 0, add)
		} else {
			const x = floats[f + ItemFloat.X] ?? 0
			const y = floats[f + ItemFloat.Y] ?? 0
			const w = floats[f + ItemFloat.WIDTH] ?? 0
			const h = floats[f + ItemFloat.HEIGHT] ?? 0

			add(x, y)
			add(x + w, y)
			add(x, y + h)
			add(x + w, y + h)
		}

		if (box.minX === Infinity) {
			box.minX = box.maxX = e
			box.minY = box.maxY = ff
		}

		const strokeIndex = ints[i + ItemInt.STROKE] ?? -1
		const stroke = strokeIndex < 0 ? undefined : this.strokes[strokeIndex]
		// Strokes widen by half their width in drawing space (scaled by the matrix's larger axis).
		const half = stroke === undefined ? 0 : (stroke.width / 2) * Math.max(Math.hypot(a, b), Math.hypot(c, d))
		const out = this.itemBounds
		const at = item * 4

		out[at] = box.minX - half
		out[at + 1] = box.minY - half
		out[at + 2] = box.maxX + half
		out[at + 3] = box.maxY + half
	}
}

// Roots of the derivative of one cubic coordinate in (0, 1).
function cubicExtrema(p0: number, p1: number, p2: number, p3: number, out: number[]): void {
	const a = -p0 + 3 * p1 - 3 * p2 + p3
	const b = 2 * (p0 - 2 * p1 + p2)
	const c = p1 - p0

	if (Math.abs(a) < 1e-12) {
		if (Math.abs(b) > 1e-12) {
			const t = -c / b

			if (t > 0 && t < 1) out.push(t)
		}

		return
	}

	const discriminant = b * b - 4 * a * c

	if (discriminant < 0) return

	const root = Math.sqrt(discriminant)
	const t1 = (-b + root) / (2 * a)
	const t2 = (-b - root) / (2 * a)

	if (t1 > 0 && t1 < 1) out.push(t1)
	if (t2 > 0 && t2 < 1) out.push(t2)
}

function cubicAt(p0: number, p1: number, p2: number, p3: number, t: number): number {
	const u = 1 - t

	return u * u * u * p0 + 3 * u * u * t * p1 + 3 * u * t * t * p2 + t * t * t * p3
}

// Feeds every extreme point of a path to `add`, in local space. Extrema are found per axis in local space,
// which bounds the curve exactly under the axis-aligned part of the matrix and tightly under rotation.
export function measurePath(
	commands: Uint8Array,
	coords: Float32Array,
	firstCommand: number,
	count: number,
	firstCoord: number,
	add: (x: number, y: number) => void
): void {
	let at = firstCoord
	let x = 0
	let y = 0
	const ts: number[] = []

	for (let k = firstCommand; k < firstCommand + count; k++) {
		switch (commands[k]) {
			case PathCommand.MOVE:
			case PathCommand.LINE:
				x = coords[at] ?? 0
				y = coords[at + 1] ?? 0
				add(x, y)
				at += 2
				break
			case PathCommand.QUAD: {
				// Elevated to a cubic for the extrema.
				const qx = coords[at] ?? 0
				const qy = coords[at + 1] ?? 0
				const ex = coords[at + 2] ?? 0
				const ey = coords[at + 3] ?? 0

				measureCubic(
					x,
					y,
					x + ((qx - x) * 2) / 3,
					y + ((qy - y) * 2) / 3,
					ex + ((qx - ex) * 2) / 3,
					ey + ((qy - ey) * 2) / 3,
					ex,
					ey,
					ts,
					add
				)
				x = ex
				y = ey
				at += 4
				break
			}
			case PathCommand.CUBIC: {
				const ex = coords[at + 4] ?? 0
				const ey = coords[at + 5] ?? 0

				measureCubic(x, y, coords[at] ?? 0, coords[at + 1] ?? 0, coords[at + 2] ?? 0, coords[at + 3] ?? 0, ex, ey, ts, add)
				x = ex
				y = ey
				at += 6
				break
			}
			default:
				break
		}
	}
}

function measureCubic(
	x0: number,
	y0: number,
	x1: number,
	y1: number,
	x2: number,
	y2: number,
	x3: number,
	y3: number,
	ts: number[],
	add: (x: number, y: number) => void
): void {
	ts.length = 0
	cubicExtrema(x0, x1, x2, x3, ts)
	cubicExtrema(y0, y1, y2, y3, ts)

	for (const t of ts) add(cubicAt(x0, x1, x2, x3, t), cubicAt(y0, y1, y2, y3, t))

	add(x3, y3)
}

// m1 * m2 (apply m2 first).
export function multiply(m1: ArrayLike<number>, m2: ArrayLike<number>): number[] {
	const a1 = m1[0] ?? 1
	const b1 = m1[1] ?? 0
	const c1 = m1[2] ?? 0
	const d1 = m1[3] ?? 1
	const e1 = m1[4] ?? 0
	const f1 = m1[5] ?? 0
	const a2 = m2[0] ?? 1
	const b2 = m2[1] ?? 0
	const c2 = m2[2] ?? 0
	const d2 = m2[3] ?? 1
	const e2 = m2[4] ?? 0
	const f2 = m2[5] ?? 0

	return [a1 * a2 + c1 * b2, b1 * a2 + d1 * b2, a1 * c2 + c1 * d2, b1 * c2 + d1 * d2, a1 * e2 + c1 * f2 + e1, b1 * e2 + d1 * f2 + f1]
}

// The placement of a box at (x, y) of w x h, rotated by `rotation` degrees about its centre and mirrored
// per flip about its centre (flip first, then rotate, then translate: ECMA-376 §20.1.7.6).
export function boxMatrix(x: number, y: number, w: number, h: number, rotation: number, flipH: boolean, flipV: boolean): number[] {
	const cx = w / 2
	const cy = h / 2
	const radians = (rotation * Math.PI) / 180
	const cos = Math.cos(radians)
	const sin = Math.sin(radians)
	const sx = flipH ? -1 : 1
	const sy = flipV ? -1 : 1
	// translate(x + cx, y + cy) * rotate * scale(sx, sy) * translate(-cx, -cy)
	const a = cos * sx
	const b = sin * sx
	const c = -sin * sy
	const d = cos * sy

	return [a, b, c, d, x + cx - (a * cx + c * cy), y + cy - (b * cx + d * cy)]
}
