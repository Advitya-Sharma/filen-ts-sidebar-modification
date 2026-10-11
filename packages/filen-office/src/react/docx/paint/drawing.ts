import {
	ITEM_FILL_MODIFIER_MASK,
	ITEM_FLOATS,
	ITEM_INTS,
	ItemFlag,
	ItemFloat,
	ItemInt,
	ItemKind,
	PATH_COMMAND_COORDS,
	PathCommand,
	type DrawingImage,
	type DrawingText,
	type GradientStops,
	type Inset,
	type Paint,
	type Rgba,
	type Shadow,
	type TextParagraph,
	type TextRun
} from "../../../core/drawingml/index"
import type { DrawingData } from "../../../dom/docx/protocol"

// Paints a drawing's display list as one SVG element. The data crossed a thread boundary, so every number is
// checked before it is formatted into an attribute and every enum is mapped through a switch onto this
// module's own strings; the only document text that reaches the DOM is run text, through textContent.

export interface DrawingPaintContext {
	readonly document: Document
	// A safe id token, unique per painted drawing on the page.
	readonly idPrefix: string
	// A blob: URL of the image's loaded media, or undefined when it cannot be shown.
	imageUrl(image: DrawingImage): string | undefined
}

const SVG_NS = "http://www.w3.org/2000/svg"
const XML_NS = "http://www.w3.org/XML/1998/namespace"
const MAX_ITEMS = 200_000
const MAX_COORD = 2 ** 31
const MAX_STOPS = 64
const MAX_DASHES = 32
const MAX_RUNS = 50_000
const PX_PER_PT = 96 / 72

type PaintSlot = "fill" | "stroke"

// Finite, clamped, three decimals: short attribute strings that can never read NaN or Infinity.
function fmt(value: number): string {
	const v = Number.isFinite(value) ? (value > MAX_COORD ? MAX_COORD : value < -MAX_COORD ? -MAX_COORD : value) : 0
	const rounded = Math.round(v * 1000) / 1000

	return rounded === 0 ? "0" : String(rounded)
}

function finite(value: number, fallback: number): number {
	return Number.isFinite(value) ? value : fallback
}

function clamp(value: number, min: number, max: number, fallback: number = min): number {
	return Number.isFinite(value) ? (value < min ? min : value > max ? max : value) : fallback
}

function index(value: number | undefined): number {
	return value !== undefined && Number.isInteger(value) && value >= 0 ? value : -1
}

function rgbaOf(value: Rgba): number {
	return Number.isFinite(value) ? value >>> 0 : 0
}

function hex(color: number): string {
	return `#${((color >>> 8) & 0xffffff).toString(16).padStart(6, "0")}`
}

function alphaOf(color: number): number {
	return (color & 255) / 255
}

function rgbaFrom(r: number, g: number, b: number, a: number): number {
	const channel = (v: number): number => clamp(Math.round(v), 0, 255)

	return ((channel(r) << 24) | (channel(g) << 16) | (channel(b) << 8) | channel(a)) >>> 0
}

function mix(c1: number, c2: number, t: number): number {
	const u = 1 - t

	return rgbaFrom(
		((c1 >>> 24) & 255) * u + ((c2 >>> 24) & 255) * t,
		((c1 >>> 16) & 255) * u + ((c2 >>> 16) & 255) * t,
		((c1 >>> 8) & 255) * u + ((c2 >>> 8) & 255) * t,
		(c1 & 255) * u + (c2 & 255) * t
	)
}

// ST_PathFillMode, approximated as a blend toward white or black.
function modify(color: number, modifier: number): number {
	switch (modifier) {
		case ItemFlag.LIGHTEN:
			return mix(color, 0xffffff00 | (color & 255), 0.4)
		case ItemFlag.LIGHTEN_LESS:
			return mix(color, 0xffffff00 | (color & 255), 0.2)
		case ItemFlag.DARKEN:
			return mix(color, color & 255, 0.4)
		case ItemFlag.DARKEN_LESS:
			return mix(color, color & 255, 0.2)
		default:
			return color
	}
}

function stopCount(stops: GradientStops): number {
	return Math.min(stops.offsets.length, stops.colors.length, MAX_STOPS)
}

function averageColor(stops: GradientStops): number {
	const count = stopCount(stops)
	let r = 0
	let g = 0
	let b = 0
	let a = 0

	for (let k = 0; k < count; k++) {
		const color = rgbaOf(stops.colors[k] ?? 0)

		r += (color >>> 24) & 255
		g += (color >>> 16) & 255
		b += (color >>> 8) & 255
		a += color & 255
	}

	return count === 0 ? 0 : rgbaFrom(r / count, g / count, b / count, a / count)
}

// Dash, gap lengths in multiples of the line width (ST_PresetLineDashVal).
function presetDashes(dash: string): readonly number[] | undefined {
	switch (dash) {
		case "dot":
			return [1, 3]
		case "dash":
			return [4, 3]
		case "lgDash":
			return [8, 3]
		case "dashDot":
			return [4, 3, 1, 3]
		case "lgDashDot":
			return [8, 3, 1, 3]
		case "lgDashDotDot":
			return [8, 3, 1, 3, 1, 3]
		case "sysDash":
			return [3, 1]
		case "sysDot":
			return [1, 1]
		case "sysDashDot":
			return [3, 1, 1, 1]
		case "sysDashDotDot":
			return [3, 1, 1, 1, 1, 1]
		default:
			return undefined
	}
}

function dashArray(dashes: ArrayLike<number>, width: number): string {
	const count = Math.min(dashes.length, MAX_DASHES)
	let out = ""
	let total = 0

	for (let k = 0; k < count; k++) {
		const length = clamp(dashes[k] ?? 0, 0, 1000, 0) * width

		total += length
		out += k === 0 ? fmt(length) : ` ${fmt(length)}`
	}

	return total > 0 ? out : ""
}

const BAYER = [
	0, 32, 8, 40, 2, 34, 10, 42, 48, 16, 56, 24, 50, 18, 58, 26, 12, 44, 4, 36, 14, 46, 6, 38, 60, 28, 52, 20, 62, 30, 54, 22, 3, 35, 11,
	43, 1, 33, 9, 41, 51, 19, 59, 27, 49, 17, 57, 25, 15, 47, 7, 39, 13, 45, 5, 37, 63, 31, 55, 23, 61, 29, 53, 21
]

function percent(value: number): (x: number, y: number) => boolean {
	const threshold = Math.round((value / 100) * 64)

	return (x, y) => (BAYER[y * 8 + x] ?? 64) < threshold
}

const mod = (value: number, by: number): number => ((value % by) + by) % by

const SMALL_CONFETTI = [0x80, 0x10, 0x02, 0x40, 0x04, 0x20, 0x01, 0x08]
const LARGE_CONFETTI = [0xc1, 0x0c, 0x30, 0xc3, 0x18, 0x60, 0x86, 0x03]

function confetti(rows: readonly number[], x: number, y: number): boolean {
	return (((rows[y] ?? 0) >> (7 - x)) & 1) === 1
}

// 8x8 foreground masks of the preset patterns; the rest blend foreground and background.
const PATTERNS = new Map<string, (x: number, y: number) => boolean>([
	["pct5", percent(5)],
	["pct10", percent(10)],
	["pct20", percent(20)],
	["pct25", percent(25)],
	["pct30", percent(30)],
	["pct40", percent(40)],
	["pct50", percent(50)],
	["pct60", percent(60)],
	["pct70", percent(70)],
	["pct75", percent(75)],
	["pct80", percent(80)],
	["pct90", percent(90)],
	["ltHorz", (_x, y) => y === 0],
	["horz", (_x, y) => y % 4 === 0],
	["narHorz", (_x, y) => y % 2 === 0],
	["dkHorz", (_x, y) => y % 4 < 2],
	["ltVert", x => x === 0],
	["vert", x => x % 4 === 0],
	["narVert", x => x % 2 === 0],
	["dkVert", x => x % 4 < 2],
	["dashHorz", (x, y) => (y === 0 && x < 4) || (y === 4 && x >= 4)],
	["dashVert", (x, y) => (x === 0 && y < 4) || (x === 4 && y >= 4)],
	["ltDnDiag", (x, y) => mod(x - y, 8) === 0],
	["dnDiag", (x, y) => mod(x - y, 4) === 0],
	["dkDnDiag", (x, y) => mod(x - y, 4) < 2],
	["wdDnDiag", (x, y) => mod(x - y, 8) < 3],
	["dashDnDiag", (x, y) => mod(x - y, 4) === 0 && y % 4 < 2],
	["ltUpDiag", (x, y) => mod(x + y, 8) === 7],
	["upDiag", (x, y) => mod(x + y, 4) === 3],
	["dkUpDiag", (x, y) => mod(x + y, 4) >= 2],
	["wdUpDiag", (x, y) => mod(x + y, 8) >= 5],
	["dashUpDiag", (x, y) => mod(x + y, 4) === 3 && y % 4 < 2],
	["cross", (x, y) => x === 0 || y === 0],
	["lgGrid", (x, y) => x === 0 || y === 0],
	["smGrid", (x, y) => x % 4 === 0 || y % 4 === 0],
	["dotGrid", (x, y) => (y === 0 && x % 2 === 0) || (x === 0 && y % 2 === 0)],
	["diagCross", (x, y) => mod(x - y, 4) === 0 || mod(x + y, 4) === 3],
	["smCheck", (x, y) => ((x >> 1) + (y >> 1)) % 2 === 0],
	["lgCheck", (x, y) => ((x >> 2) + (y >> 2)) % 2 === 0],
	["horzBrick", (x, y) => y % 4 === 0 || (y < 4 ? x === 0 : x === 4)],
	["diagBrick", (x, y) => mod(x + y, 8) === 7 || (y > 3 && mod(x - y, 8) === 4)],
	["solidDmnd", (x, y) => Math.abs(x - 3.5) + Math.abs(y - 3.5) <= 3],
	["openDmnd", (x, y) => Math.abs(x - 3.5) + Math.abs(y - 3.5) === 3],
	["dotDmnd", (x, y) => (x + y) % 2 === 0 && Math.abs(x - 3.5) + Math.abs(y - 3.5) === 3],
	["zigZag", (x, y) => y % 4 === Math.abs((x % 4) - 2) || y % 4 === Math.abs((x % 4) - 2) + 1],
	["trellis", (x, y) => (y % 2 === 0 ? x % 4 !== 3 : x % 4 === 1)],
	["smConfetti", (x, y) => confetti(SMALL_CONFETTI, x, y)],
	["lgConfetti", (x, y) => confetti(LARGE_CONFETTI, x, y)]
])

// The foreground of a preset as path data in an 8x8 tile, built once per preset.
const patternPaths = new Map<string, string>()

function patternPath(preset: string): string | undefined {
	const cached = patternPaths.get(preset)

	if (cached !== undefined) return cached

	const mask = PATTERNS.get(preset)

	if (mask === undefined) return undefined

	let d = ""

	for (let y = 0; y < 8; y++) {
		let x = 0

		while (x < 8) {
			if (!mask(x, y)) {
				x++
				continue
			}

			const start = x

			while (x < 8 && mask(x, y)) x++

			d += `M${start} ${y}h${x - start}v1h${start - x}z`
		}
	}

	patternPaths.set(preset, d)

	return d
}

function genericFamily(font: string): string {
	const name = typeof font === "string" ? font.toLowerCase() : ""

	if (name.includes("mono") || name.includes("courier") || name.includes("consol")) return "monospace"
	if (name.includes("sans")) return "sans-serif"
	if (
		name.includes("times") ||
		name.includes("cambria") ||
		name.includes("georgia") ||
		name.includes("garamond") ||
		name.includes("serif")
	) {
		return "serif"
	}

	return "sans-serif"
}

function textAnchor(align: string): "start" | "middle" | "end" {
	switch (align) {
		case "ctr":
			return "middle"
		case "r":
			return "end"
		default:
			return "start"
	}
}

function verticalRotation(vertical: string): number {
	switch (vertical) {
		case "vert":
		case "eaVert":
		case "wordArtVert":
		case "mongolianVert":
			return 90
		case "vert270":
		case "wordArtVertRtl":
			return 270
		default:
			return 0
	}
}

function insetOf(inset: Inset | undefined, min: number, max: number): [number, number, number, number] {
	if (inset === undefined || inset === null) return [0, 0, 0, 0]

	return [
		clamp(inset.left, min, max, 0),
		clamp(inset.top, min, max, 0),
		clamp(inset.right, min, max, 0),
		clamp(inset.bottom, min, max, 0)
	]
}

interface Line {
	readonly paragraph: TextParagraph
	readonly from: number
	readonly to: number
	readonly first: boolean
	readonly size: number
	readonly height: number
	before: number
	after: number
}

function runPx(run: TextRun, scale: number): number {
	const px = clamp(run.size, 1, 1000, 11) * PX_PER_PT * scale

	return finite(run.baseline, 0) !== 0 ? (px * 2) / 3 : px
}

class Painter {
	readonly root: SVGSVGElement
	private defsElement: SVGDefsElement | undefined = undefined
	private counter = 0
	private readonly paintIds = new Map<string, string>()
	private readonly shadowIds = new Map<string, string>()
	private readonly filterIds = new Map<number, string>()
	private readonly urls = new Map<number, string | undefined>()
	private runsLeft = MAX_RUNS
	// The opacity of the last paint value resolved.
	private opacity = 1
	private readonly data: DrawingData
	private readonly context: DrawingPaintContext

	constructor(data: DrawingData, context: DrawingPaintContext) {
		this.data = data
		this.context = context
		this.root = this.el("svg")
	}

	private el<K extends keyof SVGElementTagNameMap>(tag: K): SVGElementTagNameMap[K] {
		return this.context.document.createElementNS(SVG_NS, tag)
	}

	private id(kind: string): string {
		return `${this.context.idPrefix}-${kind}${this.counter++}`
	}

	private defs(): SVGDefsElement {
		if (this.defsElement === undefined) {
			this.defsElement = this.el("defs")
			this.root.insertBefore(this.defsElement, this.root.firstChild)
		}

		return this.defsElement
	}

	paint(): void {
		const data = this.data
		const root = this.root
		const width = clamp(data.width, 0, MAX_COORD, 0)
		const height = clamp(data.height, 0, MAX_COORD, 0)

		if (width > 0 && height > 0) root.setAttribute("viewBox", `0 0 ${fmt(width)} ${fmt(height)}`)

		root.setAttribute("preserveAspectRatio", "none")
		root.setAttribute("overflow", "visible")
		root.setAttribute("width", "100%")
		root.setAttribute("height", "100%")
		root.setAttribute("aria-hidden", "true")

		const floats = data.itemFloats
		const ints = data.itemInts
		const count = Math.min(
			index(data.itemCount) < 0 ? 0 : data.itemCount,
			MAX_ITEMS,
			Math.floor(floats.length / ITEM_FLOATS),
			Math.floor(ints.length / ITEM_INTS)
		)

		for (let item = 0; item < count; item++) {
			const i = item * ITEM_INTS
			const flags = ints[i + ItemInt.FLAGS] ?? 0

			if ((flags & ItemFlag.HIDDEN) !== 0) continue

			const f = item * ITEM_FLOATS
			let element: SVGElement | undefined

			switch (ints[i + ItemInt.KIND]) {
				case ItemKind.PATH:
					element = this.pathItem(i, f, flags)
					break
				case ItemKind.IMAGE:
					element = this.imageItem(i, f)
					break
				case ItemKind.TEXT:
					element = this.textItem(i, f)
					break
				default:
					break
			}

			if (element === undefined) continue

			const shadow = this.shadow(index(ints[i + ItemInt.SHADOW]))

			if (shadow === undefined) {
				root.appendChild(element)
			} else {
				// Wrapped so the offset applies in drawing space, not scaled or rotated by the item's matrix.
				const group = this.el("g")

				group.setAttribute("filter", `url(#${shadow})`)
				group.appendChild(element)
				root.appendChild(group)
			}
		}
	}

	private transform(f: number): string {
		const m = this.data.itemFloats
		const a = finite(m[f + ItemFloat.A] ?? 1, 1)
		const b = finite(m[f + ItemFloat.B] ?? 0, 0)
		const c = finite(m[f + ItemFloat.C] ?? 0, 0)
		const d = finite(m[f + ItemFloat.D] ?? 1, 1)
		const e = finite(m[f + ItemFloat.E] ?? 0, 0)
		const ff = finite(m[f + ItemFloat.F] ?? 0, 0)

		if (a === 1 && b === 0 && c === 0 && d === 1 && e === 0 && ff === 0) return ""

		return `matrix(${fmt(a)} ${fmt(b)} ${fmt(c)} ${fmt(d)} ${fmt(e)} ${fmt(ff)})`
	}

	private box(f: number): [number, number, number, number] {
		const m = this.data.itemFloats

		return [
			clamp(m[f + ItemFloat.X] ?? 0, -MAX_COORD, MAX_COORD, 0),
			clamp(m[f + ItemFloat.Y] ?? 0, -MAX_COORD, MAX_COORD, 0),
			clamp(m[f + ItemFloat.WIDTH] ?? 0, 0, MAX_COORD, 0),
			clamp(m[f + ItemFloat.HEIGHT] ?? 0, 0, MAX_COORD, 0)
		]
	}

	private pathData(i: number): string {
		const ints = this.data.itemInts
		const commands = this.data.commands
		const coords = this.data.coords
		const first = index(ints[i + ItemInt.FIRST_COMMAND])
		const count = index(ints[i + ItemInt.COMMAND_COUNT])
		let at = index(ints[i + ItemInt.FIRST_COORD])

		if (first < 0 || count <= 0 || at < 0) return ""

		const end = Math.min(commands.length, first + count)
		let d = ""

		for (let k = first; k < end; k++) {
			const command = commands[k] ?? -1
			const needed = PATH_COMMAND_COORDS[command] ?? -1

			if (needed < 0 || at + needed > coords.length) break

			switch (command) {
				case PathCommand.MOVE:
					d += `M${fmt(coords[at] ?? 0)} ${fmt(coords[at + 1] ?? 0)}`
					break
				case PathCommand.LINE:
					d += `L${fmt(coords[at] ?? 0)} ${fmt(coords[at + 1] ?? 0)}`
					break
				case PathCommand.CUBIC:
					d += `C${fmt(coords[at] ?? 0)} ${fmt(coords[at + 1] ?? 0)} ${fmt(coords[at + 2] ?? 0)} ${fmt(coords[at + 3] ?? 0)} ${fmt(coords[at + 4] ?? 0)} ${fmt(coords[at + 5] ?? 0)}`
					break
				case PathCommand.QUAD:
					d += `Q${fmt(coords[at] ?? 0)} ${fmt(coords[at + 1] ?? 0)} ${fmt(coords[at + 2] ?? 0)} ${fmt(coords[at + 3] ?? 0)}`
					break
				default:
					d += "Z"
					break
			}

			at += needed
		}

		return d
	}

	private pathItem(i: number, f: number, flags: number): SVGElement | undefined {
		const d = this.pathData(i)

		if (d === "") return undefined

		const ints = this.data.itemInts
		const path = this.el("path")
		const transform = this.transform(f)
		const [, , width, height] = this.box(f)

		path.setAttribute("d", d)

		if (transform !== "") path.setAttribute("transform", transform)

		this.applyPaint(path, "fill", index(ints[i + ItemInt.FILL]), width, height, flags & ITEM_FILL_MODIFIER_MASK)

		if ((flags & ItemFlag.EVEN_ODD) !== 0) path.setAttribute("fill-rule", "evenodd")

		this.applyStroke(path, index(ints[i + ItemInt.STROKE]), width, height)

		return path
	}

	private applyPaint(element: SVGElement, slot: PaintSlot, paintIndex: number, width: number, height: number, modifier: number): boolean {
		const value = paintIndex < 0 ? "none" : this.paintValue(paintIndex, width, height, modifier)

		if (value === "") {
			placeholderStyle(element, slot === "fill")

			return true
		}

		// Unpainted strokes are SVG's default.
		if (value === "none" && slot === "stroke") return false

		element.setAttribute(slot, value)

		if (value !== "none" && this.opacity < 1) element.setAttribute(`${slot}-opacity`, fmt(this.opacity))

		return value !== "none"
	}

	private applyStroke(element: SVGElement, strokeIndex: number, width: number, height: number): void {
		const stroke = strokeIndex < 0 ? undefined : this.data.strokes[strokeIndex]

		if (stroke === undefined || !this.applyPaint(element, "stroke", index(stroke.paint), width, height, 0)) return

		const strokeWidth = clamp(stroke.width, 0, 10_000, 1) || 1

		if (strokeWidth !== 1) element.setAttribute("stroke-width", fmt(strokeWidth))

		switch (stroke.cap) {
			case "round":
				element.setAttribute("stroke-linecap", "round")
				break
			case "square":
				element.setAttribute("stroke-linecap", "square")
				break
			default:
				break
		}

		switch (stroke.join) {
			case "round":
				element.setAttribute("stroke-linejoin", "round")
				break
			case "bevel":
				element.setAttribute("stroke-linejoin", "bevel")
				break
			default:
				element.setAttribute("stroke-miterlimit", fmt(clamp(stroke.miterLimit, 1, 100, 8)))
				break
		}

		const custom = stroke.customDash
		const dashes = custom !== null && custom !== undefined && custom.length > 0 ? custom : presetDashes(stroke.dash)

		if (dashes !== undefined) {
			const array = dashArray(dashes, strokeWidth)

			if (array !== "") element.setAttribute("stroke-dasharray", array)
		}
	}

	// A fill or stroke value: "none", a colour (opacity in this.opacity), url(#id) of a def, or "" for an
	// image that cannot be shown.
	private paintValue(paintIndex: number, width: number, height: number, modifier: number): string {
		const paint: Paint | undefined = this.data.paints[paintIndex]

		this.opacity = 1

		if (paint === undefined || paint === null) return "none"

		// Bounding-box defs vanish on a flat box (a straight line): show their average colour instead.
		const flat = !(width > 0 && height > 0)

		switch (paint.kind) {
			case "solid":
				return this.color(modify(rgbaOf(paint.color), modifier))
			case "linear":
				if (flat || stopCount(paint.stops) === 0) return this.color(averageColor(paint.stops))

				return this.linear(paintIndex, paint, width, height)
			case "path":
				if (flat || stopCount(paint.stops) === 0) return this.color(averageColor(paint.stops))

				return this.radial(paintIndex, paint)
			case "pattern":
				if (flat) return this.color(mix(rgbaOf(paint.foreground), rgbaOf(paint.background), 0.5))

				return this.pattern(paintIndex, paint)
			case "image":
				return this.imagePaint(paintIndex, paint)
			default:
				return "none"
		}
	}

	private color(color: number): string {
		const alpha = alphaOf(color)

		if (alpha === 0) return "none"

		this.opacity = alpha

		return hex(color)
	}

	private stops(gradient: SVGElement, stops: GradientStops, inner: number): void {
		const count = stopCount(stops)
		let previous = 0

		for (let k = 0; k < count; k++) {
			const offset = Math.max(previous, clamp(stops.offsets[k] ?? 0, 0, 1, 0))
			const color = rgbaOf(stops.colors[k] ?? 0)
			const stop = this.el("stop")
			const alpha = alphaOf(color)

			previous = offset
			stop.setAttribute("offset", fmt(inner + offset * (1 - inner)))
			stop.setAttribute("stop-color", hex(color))

			if (alpha < 1) stop.setAttribute("stop-opacity", fmt(alpha))

			gradient.appendChild(stop)
		}
	}

	private linear(paintIndex: number, paint: Extract<Paint, { kind: "linear" }>, width: number, height: number): string {
		const angle = mod(finite(paint.angle, 0), 360)
		const radians = (angle * Math.PI) / 180
		let dx = Math.cos(radians)
		let dy = Math.sin(radians)
		// The gradient is drawn in the unit box: an unscaled angle is kept true by stretching its direction
		// with the box (the gradient's normal transforms with the inverse of the box scale).
		const aspect = paint.scaled === true || angle % 90 === 0 ? "" : fmt(width / height)

		if (aspect !== "") {
			dx *= width
			dy *= height

			const length = Math.hypot(dx, dy)

			dx = length > 0 ? dx / length : 1
			dy = length > 0 ? dy / length : 0
		}

		const key = `l${paintIndex}:${aspect}`
		const cached = this.paintIds.get(key)

		if (cached !== undefined) return `url(#${cached})`

		const id = this.id("l")
		const gradient = this.el("linearGradient")
		// Spans the box: the first and last stops sit on the corners' projections.
		const half = (Math.abs(dx) + Math.abs(dy)) / 2

		gradient.setAttribute("id", id)
		gradient.setAttribute("x1", fmt(0.5 - dx * half))
		gradient.setAttribute("y1", fmt(0.5 - dy * half))
		gradient.setAttribute("x2", fmt(0.5 + dx * half))
		gradient.setAttribute("y2", fmt(0.5 + dy * half))
		this.stops(gradient, paint.stops, 0)
		this.defs().appendChild(gradient)
		this.paintIds.set(key, id)

		return `url(#${id})`
	}

	// a:path gradients as a radial gradient centred on the focus rectangle, which the first stop fills.
	private radial(paintIndex: number, paint: Extract<Paint, { kind: "path" }>): string {
		const key = `p${paintIndex}`
		const cached = this.paintIds.get(key)

		if (cached !== undefined) return `url(#${cached})`

		const [left, top, right, bottom] = insetOf(paint.focus, 0, 1)
		const halfWidth = Math.max(0, 1 - left - right) / 2
		const halfHeight = Math.max(0, 1 - top - bottom) / 2
		const cx = clamp(left + halfWidth, 0, 1, 0.5)
		const cy = clamp(top + halfHeight, 0, 1, 0.5)
		const reachX = Math.max(cx, 1 - cx)
		const reachY = Math.max(cy, 1 - cy)
		const rect = paint.shape === "rect"
		const radius = Math.max(rect ? Math.max(reachX, reachY) : Math.hypot(reachX, reachY), 0.001)
		const inner = clamp((rect ? Math.max(halfWidth, halfHeight) : Math.hypot(halfWidth, halfHeight)) / radius, 0, 0.99, 0)
		const id = this.id("r")
		const gradient = this.el("radialGradient")

		gradient.setAttribute("id", id)
		gradient.setAttribute("cx", fmt(cx))
		gradient.setAttribute("cy", fmt(cy))
		gradient.setAttribute("r", fmt(radius))
		this.stops(gradient, paint.stops, inner)
		this.defs().appendChild(gradient)
		this.paintIds.set(key, id)

		return `url(#${id})`
	}

	private pattern(paintIndex: number, paint: Extract<Paint, { kind: "pattern" }>): string {
		const foreground = rgbaOf(paint.foreground)
		const background = rgbaOf(paint.background)
		const d = typeof paint.pattern === "string" ? patternPath(paint.pattern) : undefined

		if (d === undefined) return this.color(mix(foreground, background, 0.5))

		const key = `t${paintIndex}`
		const cached = this.paintIds.get(key)

		if (cached !== undefined) return `url(#${cached})`

		const id = this.id("t")
		const pattern = this.el("pattern")

		pattern.setAttribute("id", id)
		pattern.setAttribute("patternUnits", "userSpaceOnUse")
		pattern.setAttribute("width", "8")
		pattern.setAttribute("height", "8")

		if (alphaOf(background) > 0) {
			const rect = this.el("rect")

			rect.setAttribute("width", "8")
			rect.setAttribute("height", "8")
			setColor(rect, "fill", background)
			pattern.appendChild(rect)
		}

		if (d !== "" && alphaOf(foreground) > 0) {
			const path = this.el("path")

			path.setAttribute("d", d)
			setColor(path, "fill", foreground)
			pattern.appendChild(path)
		}

		this.defs().appendChild(pattern)
		this.paintIds.set(key, id)

		return `url(#${id})`
	}

	// Picture fills stretch the image over the box (minus a:fillRect); tiles are approximated the same way.
	private imagePaint(paintIndex: number, paint: Extract<Paint, { kind: "image" }>): string {
		const key = `i${paintIndex}`
		const cached = this.paintIds.get(key)

		if (cached !== undefined) return `url(#${cached})`

		const [left, top, right, bottom] = insetOf(paint.stretch, -10, 1)
		const content = this.imageContent(index(paint.image), left, top, 1 - left - right, 1 - top - bottom)

		if (content === undefined) return ""

		const id = this.id("i")
		const pattern = this.el("pattern")

		pattern.setAttribute("id", id)
		pattern.setAttribute("width", "1")
		pattern.setAttribute("height", "1")
		pattern.setAttribute("patternContentUnits", "objectBoundingBox")
		pattern.appendChild(content)
		this.defs().appendChild(pattern)
		this.paintIds.set(key, id)

		return `url(#${id})`
	}

	private url(imageIndex: number, image: DrawingImage): string | undefined {
		if (this.urls.has(imageIndex)) return this.urls.get(imageIndex)

		const url = this.context.imageUrl(image)
		// Only object URLs of media the caller loaded; anything else stays a placeholder.
		const safe = typeof url === "string" && url.startsWith("blob:") ? url : undefined

		this.urls.set(imageIndex, safe)

		return safe
	}

	// The image drawn into (x, y, width, height) with its crop and effects, or undefined when it cannot be.
	private imageContent(imageIndex: number, x: number, y: number, width: number, height: number): SVGElement | undefined {
		const image = imageIndex < 0 ? undefined : this.data.images[imageIndex]

		if (image === undefined || image === null || !(width > 0 && height > 0)) return undefined

		const url = this.url(imageIndex, image)

		if (url === undefined) return undefined

		const [left, top, right, bottom] = insetOf(image.crop, -10, 1)
		const cropWidth = 1 - left - right
		const cropHeight = 1 - top - bottom

		if (!(cropWidth > 0.0001 && cropHeight > 0.0001)) return undefined

		const element = this.el("image")
		const alpha = clamp(image.alpha, 0, 1, 1)
		const filter = this.imageFilter(imageIndex, image)

		element.setAttribute("href", url)
		element.setAttribute("preserveAspectRatio", "none")

		if (alpha < 1) element.setAttribute("opacity", fmt(alpha))
		if (filter !== undefined) element.setAttribute("filter", `url(#${filter})`)

		if (left === 0 && top === 0 && right === 0 && bottom === 0) {
			element.setAttribute("x", fmt(x))
			element.setAttribute("y", fmt(y))
			element.setAttribute("width", fmt(width))
			element.setAttribute("height", fmt(height))

			return element
		}

		// a:srcRect as a viewBox over the unit image: it clips a crop and pads a negative one.
		const viewport = this.el("svg")

		viewport.setAttribute("x", fmt(x))
		viewport.setAttribute("y", fmt(y))
		viewport.setAttribute("width", fmt(width))
		viewport.setAttribute("height", fmt(height))
		viewport.setAttribute("viewBox", `${fmt(left)} ${fmt(top)} ${fmt(cropWidth)} ${fmt(cropHeight)}`)
		viewport.setAttribute("preserveAspectRatio", "none")
		element.setAttribute("width", "1")
		element.setAttribute("height", "1")
		viewport.appendChild(element)

		return viewport
	}

	// Grayscale, brightness/contrast, bi-level and duotone as one filter per image; colour change is not drawn.
	private imageFilter(imageIndex: number, image: DrawingImage): string | undefined {
		if (this.filterIds.has(imageIndex)) return this.filterIds.get(imageIndex)

		const biLevel = clamp(image.biLevel, -1, 1, -1)
		const brightness = clamp(image.brightness, -1, 1, 0)
		const contrast = clamp(image.contrast, -0.99, 0.99, 0)
		const duotone = Array.isArray(image.duotone) && image.duotone.length === 2 ? image.duotone : undefined
		const gray = image.grayscale === true || biLevel >= 0 || duotone !== undefined

		if (!gray && brightness === 0 && contrast === 0) return undefined

		const id = this.id("f")
		const filter = this.el("filter")

		filter.setAttribute("id", id)
		filter.setAttribute("color-interpolation-filters", "sRGB")

		if (gray) {
			const matrix = this.el("feColorMatrix")

			matrix.setAttribute("type", "saturate")
			matrix.setAttribute("values", "0")
			filter.appendChild(matrix)
		}

		if (brightness !== 0 || contrast !== 0) {
			const slope = (1 + contrast) / (1 - contrast)

			this.transfer(filter, "linear", [slope, 0.5 - 0.5 * slope + brightness], undefined)
		}

		if (biLevel >= 0) {
			const steps = 16
			const table: number[] = []

			for (let k = 0; k < steps; k++) table.push(k / steps < biLevel ? 0 : 1)

			this.transfer(filter, "discrete", table, undefined)
		}

		if (duotone !== undefined) this.transfer(filter, "table", [], [rgbaOf(duotone[0]), rgbaOf(duotone[1])])

		this.defs().appendChild(filter)
		this.filterIds.set(imageIndex, id)

		return id
	}

	// feComponentTransfer over R, G and B: linear [slope, intercept], a shared table, or a two-colour ramp.
	private transfer(
		filter: SVGFilterElement,
		type: "linear" | "discrete" | "table",
		values: readonly number[],
		ramp: readonly [number, number] | undefined
	): void {
		const transfer = this.el("feComponentTransfer")
		const funcs = [this.el("feFuncR"), this.el("feFuncG"), this.el("feFuncB")]

		for (let channel = 0; channel < 3; channel++) {
			const func = funcs[channel]

			if (func === undefined) continue

			func.setAttribute("type", type)

			if (type === "linear") {
				func.setAttribute("slope", fmt(values[0] ?? 1))
				func.setAttribute("intercept", fmt(values[1] ?? 0))
			} else if (ramp !== undefined) {
				const shift = 24 - channel * 8

				func.setAttribute("tableValues", `${fmt(((ramp[0] >>> shift) & 255) / 255)} ${fmt(((ramp[1] >>> shift) & 255) / 255)}`)
			} else {
				func.setAttribute("tableValues", values.map(fmt).join(" "))
			}

			transfer.appendChild(func)
		}

		filter.appendChild(transfer)
	}

	private imageItem(i: number, f: number): SVGElement | undefined {
		const [x, y, width, height] = this.box(f)

		if (!(width > 0 && height > 0)) return undefined

		const group = this.el("g")
		const transform = this.transform(f)
		const content = this.imageContent(index(this.data.itemInts[i + ItemInt.REF]), x, y, width, height)
		const clip = this.pathData(i)

		if (transform !== "") group.setAttribute("transform", transform)

		if (clip !== "") {
			const id = this.id("c")
			const clipPath = this.el("clipPath")
			const path = this.el("path")

			path.setAttribute("d", clip)
			clipPath.setAttribute("id", id)
			clipPath.appendChild(path)
			this.defs().appendChild(clipPath)
			group.setAttribute("clip-path", `url(#${id})`)
		}

		if (content !== undefined) {
			group.appendChild(content)
		} else {
			const rect = this.el("rect")

			rect.setAttribute("x", fmt(x))
			rect.setAttribute("y", fmt(y))
			rect.setAttribute("width", fmt(width))
			rect.setAttribute("height", fmt(height))
			placeholderStyle(rect, true)
			group.appendChild(rect)
		}

		return group
	}

	private shadow(shadowIndex: number): string | undefined {
		const shadow: Shadow | undefined = shadowIndex < 0 ? undefined : this.data.shadows[shadowIndex]

		if (shadow === undefined || shadow === null) return undefined

		const color = rgbaOf(shadow.color)
		const dx = shadow.dx
		const dy = shadow.dy
		const blur = shadow.blur

		if (!Number.isFinite(dx) || !Number.isFinite(dy) || !Number.isFinite(blur)) return undefined
		if (Math.abs(dx) > 4096 || Math.abs(dy) > 4096 || blur < 0 || blur > 1024 || alphaOf(color) === 0) return undefined

		const key = `${color},${fmt(dx)},${fmt(dy)},${fmt(blur)}`
		const cached = this.shadowIds.get(key)

		if (cached !== undefined) return cached

		const id = this.id("s")
		const filter = this.el("filter")
		const drop = this.el("feDropShadow")
		const width = clamp(this.data.width, 0, MAX_COORD, 0)
		const height = clamp(this.data.height, 0, MAX_COORD, 0)
		// A user-space region: a bounding-box region is empty for straight lines and would hide them.
		const margin = Math.abs(dx) + Math.abs(dy) + blur * 3 + 16

		filter.setAttribute("id", id)
		filter.setAttribute("filterUnits", "userSpaceOnUse")
		filter.setAttribute("x", fmt(-width - margin))
		filter.setAttribute("y", fmt(-height - margin))
		filter.setAttribute("width", fmt(width * 3 + margin * 2))
		filter.setAttribute("height", fmt(height * 3 + margin * 2))
		drop.setAttribute("dx", fmt(dx))
		drop.setAttribute("dy", fmt(dy))
		drop.setAttribute("stdDeviation", fmt(blur / 2))
		drop.setAttribute("flood-color", hex(color))
		drop.setAttribute("flood-opacity", fmt(alphaOf(color)))
		filter.appendChild(drop)
		this.defs().appendChild(filter)
		this.shadowIds.set(key, id)

		return id
	}

	// Rich text laid out simply: a line per paragraph or line break, no wrapping, anchored in the box.
	private textItem(i: number, f: number): SVGElement | undefined {
		const text: DrawingText | undefined = this.data.texts[index(this.data.itemInts[i + ItemInt.REF])]

		// A story is a text box the page layout draws itself.
		if (text === undefined || text === null || text.kind !== "body" || !Array.isArray(text.paragraphs)) return undefined

		const body = text.body

		if (typeof body !== "object" || body === null) return undefined

		const [x, y, width, height] = this.box(f)
		const turn = verticalRotation(body.vertical)
		const rotation = mod(finite(body.rotation, 0) + turn, 360)
		const swap = turn !== 0
		const cx = x + width / 2
		const cy = y + height / 2
		const boxWidth = swap ? height : width
		const boxHeight = swap ? width : height
		const [insetLeft, insetTop, insetRight, insetBottom] = insetOf(body.insets, 0, 10_000)
		const left = cx - boxWidth / 2 + insetLeft
		const right = cx + boxWidth / 2 - insetRight
		const top = cy - boxHeight / 2 + insetTop
		const bottom = cy + boxHeight / 2 - insetBottom
		const normal = body.autofit === "normal"
		const scale = normal ? clamp(body.fontScale, 0.01, 1, 1) : 1
		const reduction = normal ? clamp(body.lineSpacingReduction, 0, 0.9, 0) : 0
		const lines = this.lines(text.paragraphs, scale, reduction)

		if (lines.length === 0) return undefined

		let total = 0

		for (const line of lines) total += line.before + line.height + line.after

		let cursor = body.anchor === "ctr" ? top + (bottom - top - total) / 2 : body.anchor === "b" ? bottom - total : top
		const group = this.el("g")
		let transform = this.transform(f)

		if (rotation !== 0) transform += `${transform === "" ? "" : " "}rotate(${fmt(rotation)} ${fmt(cx)} ${fmt(cy)})`
		if (transform !== "") group.setAttribute("transform", transform)

		for (const line of lines) {
			cursor += line.before

			const baseline = cursor + (line.height - line.size * 1.2) / 2 + line.size * 0.95
			const element = this.line(line, left, right, baseline, scale)

			if (element !== undefined) group.appendChild(element)

			cursor += line.height + line.after
		}

		return group.firstChild === null ? undefined : group
	}

	private lines(paragraphs: readonly TextParagraph[], scale: number, reduction: number): Line[] {
		const lines: Line[] = []

		for (let p = 0; p < paragraphs.length && this.runsLeft > 0; p++) {
			const paragraph = paragraphs[p]

			if (paragraph === undefined || paragraph === null || !Array.isArray(paragraph.runs)) continue

			const runs = paragraph.runs
			const count = Math.min(runs.length, this.runsLeft)
			const spacing = finite(paragraph.lineSpacing, 0)
			const before = p === 0 ? 0 : clamp(paragraph.spaceBefore, 0, 1000, 0) * PX_PER_PT
			const after = clamp(paragraph.spaceAfter, 0, 1000, 0) * PX_PER_PT
			const firstLine = lines.length
			let from = 0

			this.runsLeft -= count

			for (let k = 0; k <= count; k++) {
				const run = runs[k]

				if (k < count && run?.lineBreak !== true) continue

				let size = 0

				for (let r = from; r < k; r++) {
					const sized = runs[r]

					if (sized !== undefined) size = Math.max(size, runPx(sized, scale))
				}

				if (size === 0) size = clamp(paragraph.endSize, 1, 1000, 11) * PX_PER_PT * scale

				const single = size * 1.2
				const height =
					(spacing > 0 ? single * Math.min(spacing, 10) : spacing < 0 ? Math.min(-spacing, 1000) * PX_PER_PT : single) *
					(1 - reduction)

				lines.push({ paragraph, from, to: k, first: lines.length === firstLine, size, height, before: 0, after: 0 })
				from = k + 1
			}

			const first = lines[firstLine]
			const last = lines[lines.length - 1]

			if (first !== undefined) first.before = before
			if (last !== undefined) last.after = after
		}

		return lines
	}

	private line(line: Line, left: number, right: number, baseline: number, scale: number): SVGElement | undefined {
		const paragraph = line.paragraph
		const runs = paragraph.runs
		const anchor = textAnchor(paragraph.align)
		const marginLeft = clamp(paragraph.marginLeft, -10_000, 10_000, 0) + (line.first ? clamp(paragraph.indent, -10_000, 10_000, 0) : 0)
		const x = anchor === "middle" ? (left + right) / 2 : anchor === "end" ? right : left + marginLeft
		const element = this.el("text")
		const bullet = line.first && typeof paragraph.bullet === "string" && paragraph.bullet !== "" ? paragraph.bullet : ""
		let parts = bullet === "" ? 0 : 1

		for (let r = line.from; r < line.to; r++) if (runs[r] !== undefined) parts++

		if (parts === 0) return undefined

		element.setAttribute("x", fmt(x))
		element.setAttribute("y", fmt(baseline))
		element.setAttributeNS(XML_NS, "xml:space", "preserve")

		if (anchor !== "start") element.setAttribute("text-anchor", anchor)

		const firstRun = runs[line.from]

		if (bullet !== "") {
			const span = parts === 1 ? element : this.el("tspan")

			span.textContent = `${symbolBullet(paragraph.bulletFont) ? "•" : bullet} `

			if (firstRun !== undefined) this.runStyle(span, firstRun, scale)
			if (span !== element) element.appendChild(span)
		}

		// A raised or lowered run shifts the following text too, so the next run shifts back.
		let shift = 0

		for (let r = line.from; r < line.to; r++) {
			const run = runs[r]

			if (run === undefined) continue

			const span = parts === 1 ? element : this.el("tspan")
			const raw = typeof run.text === "string" ? run.text : ""
			const baselineShift = clamp(run.baseline, -1, 1, 0) * clamp(run.size, 1, 1000, 11) * PX_PER_PT * scale

			span.textContent = run.caps === "all" ? raw.toUpperCase() : raw
			this.runStyle(span, run, scale)

			if (baselineShift !== 0 || shift !== 0) span.setAttribute("dy", fmt(shift - baselineShift))

			shift = baselineShift

			if (span !== element) element.appendChild(span)
		}

		return element
	}

	private runStyle(element: SVGElement, run: TextRun, scale: number): void {
		const color = rgbaOf(run.color)
		const alpha = alphaOf(color)
		const spacing = clamp(run.spacing, -100, 100, 0) * PX_PER_PT
		const underline = typeof run.underline === "string" && run.underline !== "none"
		const strike = run.strike === "sngStrike" || run.strike === "dblStrike"

		element.setAttribute("font-size", fmt(runPx(run, scale)))
		element.setAttribute("font-family", genericFamily(run.font))
		element.setAttribute("fill", hex(color))

		if (alpha < 1) element.setAttribute("fill-opacity", fmt(alpha))
		if (run.bold === true) element.setAttribute("font-weight", "bold")
		if (run.italic === true) element.setAttribute("font-style", "italic")
		if (run.caps === "small") element.setAttribute("font-variant", "small-caps")
		if (spacing !== 0) element.setAttribute("letter-spacing", fmt(spacing))
		if (underline || strike)
			element.setAttribute(
				"text-decoration",
				underline && strike ? "underline line-through" : underline ? "underline" : "line-through"
			)
	}
}

// Symbol-font bullets (Wingdings, Symbol) map their glyphs onto letters; a plain bullet reads better.
function symbolBullet(font: string): boolean {
	const name = typeof font === "string" ? font.toLowerCase() : ""

	return name.includes("wingdings") || name.includes("symbol") || name.includes("webdings")
}

function setColor(element: SVGElement, slot: PaintSlot, color: number): void {
	const alpha = alphaOf(color)

	element.setAttribute(slot, hex(color))

	if (alpha < 1) element.setAttribute(`${slot}-opacity`, fmt(alpha))
}

// Themeable through CSS custom properties on the page.
function placeholderStyle(element: SVGElement, filled: boolean): void {
	if (filled) element.style.setProperty("fill", "var(--fo-placeholder-fill, #f1f1f1)")

	element.style.setProperty("stroke", "var(--fo-placeholder-stroke, #c8c8c8)")
	element.setAttribute("stroke-width", "1")
	element.setAttribute("vector-effect", "non-scaling-stroke")
}

export function paintDrawing(data: DrawingData, context: DrawingPaintContext): SVGSVGElement {
	const painter = new Painter(data, context)

	painter.paint()

	return painter.root
}

// Every image the drawing shows (picture fills and image items, hidden items skipped), once per part, so the
// caller can load media before painting.
export function drawingImages(data: DrawingData): DrawingImage[] {
	const ints = data.itemInts
	const count = Math.min(index(data.itemCount) < 0 ? 0 : data.itemCount, MAX_ITEMS, Math.floor(ints.length / ITEM_INTS))
	const seen = new Set<string>()
	const out: DrawingImage[] = []
	const add = (imageIndex: number): void => {
		const image = imageIndex < 0 ? undefined : data.images[imageIndex]

		if (image === undefined || image === null) return

		const key = image.svgPartName ?? image.partName

		if (typeof key !== "string" || seen.has(key)) return

		seen.add(key)
		out.push(image)
	}
	const addPaint = (paintIndex: number): void => {
		const paint = paintIndex < 0 ? undefined : data.paints[paintIndex]

		if (paint !== undefined && paint !== null && paint.kind === "image") add(index(paint.image))
	}

	for (let item = 0; item < count; item++) {
		const i = item * ITEM_INTS

		if (((ints[i + ItemInt.FLAGS] ?? 0) & ItemFlag.HIDDEN) !== 0) continue

		switch (ints[i + ItemInt.KIND]) {
			case ItemKind.PATH: {
				addPaint(index(ints[i + ItemInt.FILL]))

				const stroke = data.strokes[index(ints[i + ItemInt.STROKE])]

				if (stroke !== undefined && stroke !== null) addPaint(index(stroke.paint))

				break
			}
			case ItemKind.IMAGE:
				add(index(ints[i + ItemInt.REF]))
				break
			default:
				break
		}
	}

	return out
}

// The MIME type an image is shown as through <image>/<img> (an SVG there runs no script), or "" when it
// cannot be shown natively. An SVG rendition, when there is one, is what the caller loads.
export function imageMime(image: DrawingImage): string {
	if (typeof image.svgPartName === "string") return "image/svg+xml"

	switch (image.format) {
		case "png":
			return "image/png"
		case "jpeg":
			return "image/jpeg"
		case "gif":
			return "image/gif"
		case "bmp":
			return "image/bmp"
		case "webp":
			return "image/webp"
		case "svg":
			return "image/svg+xml"
		default:
			return ""
	}
}
