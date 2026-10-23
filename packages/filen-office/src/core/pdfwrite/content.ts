import { PATH_COMMAND_COORDS, PathCommand } from "../drawingml/displayList"
import { ByteWriter } from "./bytes"

// Content stream operators (ISO 32000-1 §8.4-8.5, §9.4, Annex A) written into a reusable buffer. Coordinates
// are whatever space the current matrix sets up; the page sets up px from the top-left (see page.ts).

export type Matrix = readonly [number, number, number, number, number, number]

export const LineCap = { butt: 0, round: 1, square: 2 } as const
export const LineJoin = { miter: 0, round: 1, bevel: 2 } as const

// Coordinates to 1/100 (px: 1/9600 inch); matrix scale and shear terms to 1/100000.
const COORD = 2
const SCALE = 5

export class ContentBuilder {
	readonly out: ByteWriter
	// The current point and the start of the current subpath, for quadratic segments and closes.
	private cx = 0
	private cy = 0
	private sx = 0
	private sy = 0

	constructor(out: ByteWriter = new ByteWriter(16384)) {
		this.out = out
	}

	op(operator: string): this {
		this.out.ascii(operator).newline()

		return this
	}

	protected n(value: number): this {
		this.out.num(value, COORD).space()

		return this
	}

	protected s(value: number): this {
		this.out.num(value, SCALE).space()

		return this
	}

	saveState(): this {
		return this.op("q")
	}

	restoreState(): this {
		return this.op("Q")
	}

	concat(m: Matrix): this {
		this.s(m[0]).s(m[1]).s(m[2]).s(m[3]).n(m[4]).n(m[5])

		return this.op("cm")
	}

	moveTo(x: number, y: number): this {
		this.cx = this.sx = x
		this.cy = this.sy = y

		return this.n(x).n(y).op("m")
	}

	lineTo(x: number, y: number): this {
		this.cx = x
		this.cy = y

		return this.n(x).n(y).op("l")
	}

	curveTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): this {
		this.cx = x
		this.cy = y

		return this.n(x1).n(y1).n(x2).n(y2).n(x).n(y).op("c")
	}

	// A quadratic Bézier as the cubic with the same curve (control points 2/3 of the way to the quadratic's).
	quadTo(qx: number, qy: number, x: number, y: number): this {
		const x0 = this.cx
		const y0 = this.cy

		return this.curveTo(x0 + ((qx - x0) * 2) / 3, y0 + ((qy - y0) * 2) / 3, x + ((qx - x) * 2) / 3, y + ((qy - y) * 2) / 3, x, y)
	}

	closePath(): this {
		this.cx = this.sx
		this.cy = this.sy

		return this.op("h")
	}

	rect(x: number, y: number, width: number, height: number): this {
		this.cx = this.sx = x
		this.cy = this.sy = y

		return this.n(x).n(y).n(width).n(height).op("re")
	}

	// Display-list path commands (drawingml/displayList.ts), each point through `m` when given (a stroke then
	// keeps its width; set the matrix with concat to scale strokes along).
	pathCommands(commands: Uint8Array, coords: Float32Array, firstCommand: number, count: number, firstCoord: number, m?: Matrix): this {
		let at = firstCoord
		const a = m?.[0] ?? 1
		const b = m?.[1] ?? 0
		const c = m?.[2] ?? 0
		const d = m?.[3] ?? 1
		const e = m?.[4] ?? 0
		const f = m?.[5] ?? 0
		const px = (index: number): number => {
			const x = coords[index] ?? 0
			const y = coords[index + 1] ?? 0

			return a * x + c * y + e
		}
		const py = (index: number): number => {
			const x = coords[index] ?? 0
			const y = coords[index + 1] ?? 0

			return b * x + d * y + f
		}

		for (let index = firstCommand; index < firstCommand + count; index++) {
			const command = commands[index] ?? PathCommand.CLOSE

			switch (command) {
				case PathCommand.MOVE:
					this.moveTo(px(at), py(at))
					break
				case PathCommand.LINE:
					this.lineTo(px(at), py(at))
					break
				case PathCommand.CUBIC:
					this.curveTo(px(at), py(at), px(at + 2), py(at + 2), px(at + 4), py(at + 4))
					break
				case PathCommand.QUAD:
					this.quadTo(px(at), py(at), px(at + 2), py(at + 2))
					break
				default:
					this.closePath()
			}

			at += PATH_COMMAND_COORDS[command] ?? 0
		}

		return this
	}

	fill(evenOdd: boolean = false): this {
		return this.op(evenOdd ? "f*" : "f")
	}

	stroke(): this {
		return this.op("S")
	}

	fillStroke(evenOdd: boolean = false): this {
		return this.op(evenOdd ? "B*" : "B")
	}

	endPath(): this {
		return this.op("n")
	}

	// Intersects the clip with the current path and ends it (W n).
	clip(evenOdd: boolean = false): this {
		return this.op(evenOdd ? "W* n" : "W n")
	}

	lineWidth(width: number): this {
		return this.n(width).op("w")
	}

	lineCap(cap: keyof typeof LineCap): this {
		return this.op(`${LineCap[cap]} J`)
	}

	lineJoin(join: keyof typeof LineJoin): this {
		return this.op(`${LineJoin[join]} j`)
	}

	miterLimit(limit: number): this {
		return this.n(Math.max(1, limit)).op("M")
	}

	// Dash lengths in the current space; empty for solid.
	dash(lengths: readonly number[], phase: number = 0): this {
		this.out.byte(0x5b)

		for (const length of lengths) this.n(Math.max(0, length))

		this.out.ascii("] ")

		return this.n(phase).op("d")
	}

	// 0xRRGGBB (alpha is separate: graphics state ca/CA). Protected: the page tracks colour and alpha state.
	protected fillRgb(rgb: number): this {
		return this.rgb(rgb).op("rg")
	}

	protected strokeRgb(rgb: number): this {
		return this.rgb(rgb).op("RG")
	}

	private rgb(rgb: number): this {
		this.out.num(((rgb >>> 16) & 255) / 255, 3).space()
		this.out.num(((rgb >>> 8) & 255) / 255, 3).space()
		this.out.num((rgb & 255) / 255, 3).space()

		return this
	}

	// Name operators (gs, Do, sh, Tf).
	named(name: string, operator: string): this {
		this.out.byte(0x2f).ascii(name).space()

		return this.op(operator)
	}
}
