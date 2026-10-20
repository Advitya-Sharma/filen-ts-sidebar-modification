import type { SheetPatternType, SheetRenderStyle } from "../../../core/sheet/protocol"
import { rgb, type GridCanvas } from "./canvas"

// Pattern fills (ECMA-376 §18.18.55) as Excel's 8×8 swatches: a set bit draws the pattern colour, a clear bit the
// background. Without a canvas to build the swatch on, the fill is the two colours mixed by the pattern's share.

const SWATCHES: Record<SheetPatternType, readonly number[]> = {
	gray0625: [0x80, 0, 0x08, 0, 0x80, 0, 0x08, 0],
	gray125: [0x88, 0, 0x22, 0, 0x88, 0, 0x22, 0],
	lightGray: [0x88, 0x22, 0x88, 0x22, 0x88, 0x22, 0x88, 0x22],
	mediumGray: [0xaa, 0x55, 0xaa, 0x55, 0xaa, 0x55, 0xaa, 0x55],
	darkGray: [0x77, 0xdd, 0x77, 0xdd, 0x77, 0xdd, 0x77, 0xdd],
	darkHorizontal: [0xff, 0xff, 0, 0, 0xff, 0xff, 0, 0],
	lightHorizontal: [0xff, 0, 0, 0, 0xff, 0, 0, 0],
	darkVertical: [0xcc, 0xcc, 0xcc, 0xcc, 0xcc, 0xcc, 0xcc, 0xcc],
	lightVertical: [0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88],
	darkDown: [0xcc, 0x66, 0x33, 0x99, 0xcc, 0x66, 0x33, 0x99],
	lightDown: [0x88, 0x44, 0x22, 0x11, 0x88, 0x44, 0x22, 0x11],
	darkUp: [0x33, 0x66, 0xcc, 0x99, 0x33, 0x66, 0xcc, 0x99],
	lightUp: [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88],
	darkGrid: [0xff, 0xff, 0xcc, 0xcc, 0xff, 0xff, 0xcc, 0xcc],
	lightGrid: [0xff, 0x88, 0x88, 0x88, 0xff, 0x88, 0x88, 0x88],
	darkTrellis: [0x77, 0xff, 0xdd, 0xff, 0x77, 0xff, 0xdd, 0xff],
	lightTrellis: [0x88, 0x55, 0x22, 0x55, 0x88, 0x55, 0x22, 0x55]
}

function share(type: SheetPatternType): number {
	let bits = 0

	for (const row of SWATCHES[type]) for (let bit = row; bit !== 0; bit &= bit - 1) bits++

	return bits / 64
}

function mix(a: number, b: number, t: number): number {
	let out = 0

	for (const shift of [16, 8, 0]) {
		const value = Math.round(((a >> shift) & 0xff) * t + ((b >> shift) & 0xff) * (1 - t))

		out |= value << shift
	}

	return out
}

// An 8×8 swatch to build a pattern from: the image createPattern takes, and its context.
export type SwatchFactory = (size: number) => { readonly image: unknown; readonly context: GridCanvas } | null

interface CanvasDocument {
	createElement(tag: "canvas"): { width: number; height: number; getContext(type: "2d"): GridCanvas | null }
}

export function canvasSwatches(): SwatchFactory {
	return size => {
		const doc: unknown = Reflect.get(globalThis, "document")

		if (typeof doc !== "object" || doc === null) return null

		const element = (doc as CanvasDocument).createElement("canvas")

		element.width = size
		element.height = size

		const context = element.getContext("2d")

		return context === null ? null : { image: element, context }
	}
}

export class PatternCache {
	private readonly patterns = new Map<string, object | string>()

	private readonly swatches: SwatchFactory

	constructor(swatches: SwatchFactory) {
		this.swatches = swatches
	}

	// The fill style a patterned cell is painted with.
	fill(canvas: GridCanvas, pattern: NonNullable<SheetRenderStyle["pattern"]>, pixelRatio: number): object | string {
		const key = `${pattern.type}|${String(pattern.color)}|${String(pattern.background)}|${String(pixelRatio)}`
		let fill = this.patterns.get(key)

		if (fill !== undefined) return fill

		const swatch = this.swatches(8 * pixelRatio)

		if (swatch === null) {
			fill = rgb(mix(pattern.color, pattern.background, share(pattern.type)))
		} else {
			const { context } = swatch

			context.fillStyle = rgb(pattern.background)
			context.fillRect(0, 0, 8 * pixelRatio, 8 * pixelRatio)
			context.fillStyle = rgb(pattern.color)
			SWATCHES[pattern.type].forEach((row, y) => {
				for (let x = 0; x < 8; x++)
					if ((row & (0x80 >> x)) !== 0) context.fillRect(x * pixelRatio, y * pixelRatio, pixelRatio, pixelRatio)
			})

			const created = canvas.createPattern(swatch.image as never, "repeat")

			fill = created ?? rgb(mix(pattern.color, pattern.background, share(pattern.type)))

			// The swatch is drawn in device pixels; the grid's transform scales by the ratio.
			const matrix: unknown = Reflect.get(globalThis, "DOMMatrix")
			const transform: unknown = created === null ? undefined : Reflect.get(created, "setTransform")

			if (pixelRatio !== 1 && typeof matrix === "function" && typeof transform === "function") {
				const scaled = new (matrix as new (init: number[]) => object)([1 / pixelRatio, 0, 0, 1 / pixelRatio, 0, 0])

				;(transform as (this: object, matrix: object) => void).call(created as object, scaled)
			}
		}

		if (this.patterns.size > 256) this.patterns.clear()

		this.patterns.set(key, fill)

		return fill
	}
}
