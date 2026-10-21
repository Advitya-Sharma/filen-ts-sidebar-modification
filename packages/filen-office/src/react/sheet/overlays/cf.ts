import { CellFlag, type SheetTile } from "../../../core/sheet/protocol"
import { CfBarFlag, ICON_SETS } from "../../../core/sheet/cf/layers"
import type { GridDecoration, GridPaintScene, SheetCellDecorator, SheetGridExtension } from "../layers"
import { rgb, type GridCanvas } from "../paint/canvas"

// Conditional formatting's data bars and icons (the tile's cfBars and cfIcons layers), drawn as Excel draws them: a
// bar under the cell's text, inset from its edges, a gradient fading out or solid, its border and the axis between
// negative and positive values; an icon at the cell's left, as the cell aligns its text vertically.

// What the icons' paths need of a 2D context beyond GridCanvas (a browser's has it; without it icons are drawn from
// rectangles).
interface PathCanvas extends GridCanvas {
	fill(): void
	closePath(): void
	arc(x: number, y: number, radius: number, start: number, end: number): void
}

// A linear gradient (a browser's 2D context has it; without it a gradient bar is drawn solid).
interface GradientCanvas extends GridCanvas {
	createLinearGradient(x0: number, y0: number, x1: number, y1: number): { addColorStop(offset: number, color: string): void }
}

function gradients(canvas: GridCanvas): GradientCanvas | null {
	return typeof (canvas as Partial<GradientCanvas>).createLinearGradient === "function" ? (canvas as GradientCanvas) : null
}

function paths(canvas: GridCanvas): PathCanvas | null {
	const candidate = canvas as Partial<PathCanvas>

	return typeof candidate.fill === "function" && typeof candidate.closePath === "function" && typeof candidate.arc === "function"
		? (canvas as PathCanvas)
		: null
}

// Each layer's entries by entry, built when a cell of its tile is first drawn.
const indexes = new WeakMap<object, Map<number, number>>()

function indexOf(layer: { readonly entries: Uint16Array }, entry: number): number {
	let index = indexes.get(layer)

	if (index === undefined) {
		const built = new Map<number, number>()

		layer.entries.forEach((value, k) => built.set(value, k))
		index = built
		indexes.set(layer, built)
	}

	return index.get(entry) ?? -1
}

// Bars and icons keep off the cell's edges and gridline by this much at 100 %.
const INSET = 2
const ICON = 16
// Excel's gradient fades to (nearly) white: to a tenth of the colour over the cell, so over a dark theme's cells it
// fades to them and their light text stays readable.
const FADE_ALPHA = 0.1

function rgba(color: number, alpha: number): string {
	return `rgba(${String((color >> 16) & 0xff)}, ${String((color >> 8) & 0xff)}, ${String(color & 0xff)}, ${String(alpha)})`
}

function paintBar(canvas: GridCanvas, cell: GridDecoration, scene: GridPaintScene): void {
	const layer = (cell.tile as SheetTile).layers?.cfBars

	if (layer === undefined) return

	const k = indexOf(layer, cell.entry)

	if (k < 0) return

	const { box } = cell
	const inset = INSET * scene.scale
	const left = box.x + inset
	const width = Math.max(0, box.width - 1 - 2 * inset)
	const top = box.y + inset
	const height = Math.max(0, box.height - 1 - 2 * inset)
	const start = left + (layer.starts[k] ?? 0) * width
	const end = left + (layer.ends[k] ?? 0) * width
	const fill = layer.fills[k] ?? 0
	const flags = layer.flags[k] ?? 0
	const negative = !Number.isNaN(layer.axes[k] ?? Number.NaN) && end <= left + (layer.axes[k] ?? 0) * width + 0.001 && start < end
	const span = end - start

	if (span > 0 && height > 0) {
		const gradient = (flags & CfBarFlag.GRADIENT) !== 0 ? gradients(canvas) : null

		if (gradient !== null) {
			// From the colour at the bar's base (its axis side for a negative bar), one gradient per bar.
			const fade = negative ? gradient.createLinearGradient(end, 0, start, 0) : gradient.createLinearGradient(start, 0, end, 0)

			fade.addColorStop(0, rgb(fill))
			fade.addColorStop(1, rgba(fill, FADE_ALPHA))
			canvas.fillStyle = fade
			canvas.fillRect(start, top, span, height)
		} else {
			canvas.fillStyle = rgb(fill)
			canvas.fillRect(start, top, span, height)
		}

		const border = layer.borders[k] ?? -1

		if (border >= 0) {
			canvas.strokeStyle = rgb(border)
			canvas.lineWidth = 1
			canvas.beginPath()
			canvas.rect(start + 0.5, top + 0.5, Math.max(0, span - 1), Math.max(0, height - 1))
			canvas.stroke()
		}
	}

	const axis = layer.axes[k] ?? Number.NaN
	const axisColor = layer.axisColors[k] ?? -1

	if (!Number.isNaN(axis) && axisColor >= 0) {
		// A dashed line, as Excel draws the axis.
		const x = Math.round(left + axis * width) + 0.5

		canvas.strokeStyle = rgb(axisColor)
		canvas.lineWidth = 1
		canvas.setLineDash([2, 1])
		canvas.beginPath()
		canvas.moveTo(x, box.y + 1)
		canvas.lineTo(x, box.y + box.height - 2)
		canvas.stroke()
		canvas.setLineDash([])
	}
}

// --- Icons ---

const RED = 0xd9383a
const YELLOW = 0xf2b632
const GREEN = 0x3e9b4f
const GRAY = 0x8c8c8c
const BLACK = 0x3a3a3a
const PINK = 0xf2a3a5
const BLUE = 0x4472c4
const SILVER = 0xd0d0d0

type Shape =
	| { readonly kind: "arrow"; readonly angle: number; readonly color: number }
	| { readonly kind: "circle"; readonly color: number; readonly rim?: boolean; readonly fraction?: number }
	| { readonly kind: "flag"; readonly color: number }
	| { readonly kind: "diamond"; readonly color: number }
	| { readonly kind: "triangle"; readonly color: number; readonly down: boolean }
	| { readonly kind: "dash"; readonly color: number }
	| { readonly kind: "symbol"; readonly symbol: "cross" | "bang" | "check"; readonly color: number; readonly circled: boolean }
	| { readonly kind: "bars"; readonly filled: number; readonly of: number }
	| { readonly kind: "star"; readonly fill: number }
	| { readonly kind: "boxes"; readonly filled: number }

// Each set's icons, the lowest value's first ([MS-XLSX] §2.6.81 / §18.18.42 icon order).
const ICONS: Readonly<Record<(typeof ICON_SETS)[number], readonly Shape[]>> = {
	"3Arrows": [
		{ kind: "arrow", angle: 90, color: RED },
		{ kind: "arrow", angle: 0, color: YELLOW },
		{ kind: "arrow", angle: -90, color: GREEN }
	],
	"3ArrowsGray": [
		{ kind: "arrow", angle: 90, color: GRAY },
		{ kind: "arrow", angle: 0, color: GRAY },
		{ kind: "arrow", angle: -90, color: GRAY }
	],
	"3Flags": [
		{ kind: "flag", color: RED },
		{ kind: "flag", color: YELLOW },
		{ kind: "flag", color: GREEN }
	],
	"3TrafficLights1": [
		{ kind: "circle", color: RED },
		{ kind: "circle", color: YELLOW },
		{ kind: "circle", color: GREEN }
	],
	"3TrafficLights2": [
		{ kind: "circle", color: RED, rim: true },
		{ kind: "circle", color: YELLOW, rim: true },
		{ kind: "circle", color: GREEN, rim: true }
	],
	"3Signs": [
		{ kind: "diamond", color: RED },
		{ kind: "triangle", color: YELLOW, down: false },
		{ kind: "circle", color: GREEN }
	],
	"3Symbols": [
		{ kind: "symbol", symbol: "cross", color: RED, circled: true },
		{ kind: "symbol", symbol: "bang", color: YELLOW, circled: true },
		{ kind: "symbol", symbol: "check", color: GREEN, circled: true }
	],
	"3Symbols2": [
		{ kind: "symbol", symbol: "cross", color: RED, circled: false },
		{ kind: "symbol", symbol: "bang", color: YELLOW, circled: false },
		{ kind: "symbol", symbol: "check", color: GREEN, circled: false }
	],
	"4Arrows": [
		{ kind: "arrow", angle: 90, color: RED },
		{ kind: "arrow", angle: 45, color: YELLOW },
		{ kind: "arrow", angle: -45, color: YELLOW },
		{ kind: "arrow", angle: -90, color: GREEN }
	],
	"4ArrowsGray": [
		{ kind: "arrow", angle: 90, color: GRAY },
		{ kind: "arrow", angle: 45, color: GRAY },
		{ kind: "arrow", angle: -45, color: GRAY },
		{ kind: "arrow", angle: -90, color: GRAY }
	],
	"4RedToBlack": [
		{ kind: "circle", color: BLACK },
		{ kind: "circle", color: GRAY },
		{ kind: "circle", color: PINK },
		{ kind: "circle", color: RED }
	],
	"4Rating": [
		{ kind: "bars", filled: 1, of: 4 },
		{ kind: "bars", filled: 2, of: 4 },
		{ kind: "bars", filled: 3, of: 4 },
		{ kind: "bars", filled: 4, of: 4 }
	],
	"4TrafficLights": [
		{ kind: "circle", color: BLACK },
		{ kind: "circle", color: RED },
		{ kind: "circle", color: YELLOW },
		{ kind: "circle", color: GREEN }
	],
	"5Arrows": [
		{ kind: "arrow", angle: 90, color: RED },
		{ kind: "arrow", angle: 45, color: YELLOW },
		{ kind: "arrow", angle: 0, color: YELLOW },
		{ kind: "arrow", angle: -45, color: YELLOW },
		{ kind: "arrow", angle: -90, color: GREEN }
	],
	"5ArrowsGray": [
		{ kind: "arrow", angle: 90, color: GRAY },
		{ kind: "arrow", angle: 45, color: GRAY },
		{ kind: "arrow", angle: 0, color: GRAY },
		{ kind: "arrow", angle: -45, color: GRAY },
		{ kind: "arrow", angle: -90, color: GRAY }
	],
	"5Rating": [
		{ kind: "bars", filled: 0, of: 4 },
		{ kind: "bars", filled: 1, of: 4 },
		{ kind: "bars", filled: 2, of: 4 },
		{ kind: "bars", filled: 3, of: 4 },
		{ kind: "bars", filled: 4, of: 4 }
	],
	"5Quarters": [
		{ kind: "circle", color: BLACK, fraction: 0 },
		{ kind: "circle", color: BLACK, fraction: 0.25 },
		{ kind: "circle", color: BLACK, fraction: 0.5 },
		{ kind: "circle", color: BLACK, fraction: 0.75 },
		{ kind: "circle", color: BLACK, fraction: 1 }
	],
	"3Stars": [
		{ kind: "star", fill: 0 },
		{ kind: "star", fill: 0.5 },
		{ kind: "star", fill: 1 }
	],
	"3Triangles": [
		{ kind: "triangle", color: RED, down: true },
		{ kind: "dash", color: YELLOW },
		{ kind: "triangle", color: GREEN, down: false }
	],
	"5Boxes": [
		{ kind: "boxes", filled: 0 },
		{ kind: "boxes", filled: 1 },
		{ kind: "boxes", filled: 2 },
		{ kind: "boxes", filled: 3 },
		{ kind: "boxes", filled: 4 }
	]
}

export function iconShape(set: number, icon: number): Shape | null {
	const name = ICON_SETS[set]

	return name === undefined ? null : (ICONS[name][icon] ?? null)
}

function polygon(canvas: PathCanvas, points: readonly (readonly [number, number])[], color: number): void {
	canvas.fillStyle = rgb(color)
	canvas.beginPath()

	points.forEach(([x, y], i) => (i === 0 ? canvas.moveTo(x, y) : canvas.lineTo(x, y)))
	canvas.closePath()
	canvas.fill()
}

function circle(canvas: PathCanvas, cx: number, cy: number, r: number, color: number): void {
	canvas.fillStyle = rgb(color)
	canvas.beginPath()
	canvas.arc(cx, cy, r, 0, Math.PI * 2)
	canvas.fill()
}

function ring(canvas: PathCanvas, cx: number, cy: number, r: number, color: number, width: number): void {
	canvas.strokeStyle = rgb(color)
	canvas.lineWidth = width
	canvas.beginPath()
	canvas.arc(cx, cy, r, 0, Math.PI * 2)
	canvas.stroke()
}

function line(canvas: GridCanvas, points: readonly (readonly [number, number])[], color: number, width: number): void {
	canvas.strokeStyle = rgb(color)
	canvas.lineWidth = width
	canvas.beginPath()
	points.forEach(([x, y], i) => (i === 0 ? canvas.moveTo(x, y) : canvas.lineTo(x, y)))
	canvas.stroke()
}

// One icon in the square at (x, y) of side s.
export function drawIcon(canvas: GridCanvas, shape: Shape, x: number, y: number, s: number): void {
	const path = paths(canvas)
	const cx = x + s / 2
	const cy = y + s / 2
	const r = s * 0.42

	if (path === null) {
		// No paths: a square of the icon's colour.
		const color = "color" in shape ? shape.color : shape.kind === "bars" ? BLUE : shape.kind === "star" ? YELLOW : BLACK

		canvas.fillStyle = rgb(color)
		canvas.fillRect(x + s * 0.2, y + s * 0.2, s * 0.6, s * 0.6)

		return
	}

	switch (shape.kind) {
		case "arrow": {
			// An arrow pointing right, turned (90: down, -90: up).
			const angle = (shape.angle * Math.PI) / 180
			const cos = Math.cos(angle)
			const sin = Math.sin(angle)
			const at = (px: number, py: number): [number, number] => [cx + (px * cos - py * sin) * s, cy + (px * sin + py * cos) * s]

			polygon(
				path,
				[at(-0.4, -0.1), at(0.05, -0.1), at(0.05, -0.32), at(0.42, 0), at(0.05, 0.32), at(0.05, 0.1), at(-0.4, 0.1)],
				shape.color
			)
			break
		}
		case "circle": {
			if (shape.fraction === undefined) {
				if (shape.rim === true) circle(path, cx, cy, r, BLACK)

				circle(path, cx, cy, shape.rim === true ? r * 0.78 : r, shape.color)
				break
			}

			circle(path, cx, cy, r, 0xffffff)

			if (shape.fraction >= 1) {
				circle(path, cx, cy, r, shape.color)
			} else if (shape.fraction > 0) {
				path.fillStyle = rgb(shape.color)
				path.beginPath()
				path.moveTo(cx, cy)
				path.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + shape.fraction * Math.PI * 2)
				path.closePath()
				path.fill()
			}

			ring(path, cx, cy, r, shape.color, Math.max(1, s / 16))
			break
		}
		case "flag":
			line(
				path,
				[
					[x + s * 0.25, y + s * 0.1],
					[x + s * 0.25, y + s * 0.92]
				],
				BLACK,
				Math.max(1, s / 12)
			)
			polygon(
				path,
				[
					[x + s * 0.27, y + s * 0.1],
					[x + s * 0.85, y + s * 0.32],
					[x + s * 0.27, y + s * 0.55]
				],
				shape.color
			)
			break
		case "diamond":
			polygon(
				path,
				[
					[cx, y + s * 0.08],
					[x + s * 0.92, cy],
					[cx, y + s * 0.92],
					[x + s * 0.08, cy]
				],
				shape.color
			)
			break
		case "triangle":
			polygon(
				path,
				shape.down
					? [
							[x + s * 0.1, y + s * 0.25],
							[x + s * 0.9, y + s * 0.25],
							[cx, y + s * 0.85]
						]
					: [
							[cx, y + s * 0.12],
							[x + s * 0.92, y + s * 0.85],
							[x + s * 0.08, y + s * 0.85]
						],
				shape.color
			)
			break
		case "dash":
			path.fillStyle = rgb(shape.color)
			path.fillRect(x + s * 0.2, cy - s * 0.1, s * 0.6, s * 0.2)
			break
		case "symbol": {
			const ink = shape.circled ? 0xffffff : shape.color
			const width = Math.max(1.5, s / 7)

			if (shape.circled) circle(path, cx, cy, r, shape.color)

			if (shape.symbol === "cross") {
				line(
					path,
					[
						[cx - r * 0.45, cy - r * 0.45],
						[cx + r * 0.45, cy + r * 0.45]
					],
					ink,
					width
				)
				line(
					path,
					[
						[cx + r * 0.45, cy - r * 0.45],
						[cx - r * 0.45, cy + r * 0.45]
					],
					ink,
					width
				)
			} else if (shape.symbol === "bang") {
				line(
					path,
					[
						[cx, cy - r * 0.55],
						[cx, cy + r * 0.15]
					],
					ink,
					width
				)
				path.fillStyle = rgb(ink)
				path.fillRect(cx - width / 2, cy + r * 0.35, width, width)
			} else {
				line(
					path,
					[
						[cx - r * 0.5, cy],
						[cx - r * 0.1, cy + r * 0.4],
						[cx + r * 0.5, cy - r * 0.4]
					],
					ink,
					width
				)
			}
			break
		}
		case "bars": {
			const gap = s * 0.06
			const bar = (s * 0.8 - gap * (shape.of - 1)) / shape.of

			for (let i = 0; i < shape.of; i++) {
				const height = s * 0.8 * ((i + 1) / shape.of)

				path.fillStyle = rgb(i < shape.filled ? BLUE : SILVER)
				path.fillRect(x + s * 0.1 + i * (bar + gap), y + s * 0.9 - height, bar, height)
			}
			break
		}
		case "star": {
			const points: [number, number][] = []

			for (let i = 0; i < 10; i++) {
				const radius = i % 2 === 0 ? r : r * 0.45
				const angle = -Math.PI / 2 + (i * Math.PI) / 5

				points.push([cx + Math.cos(angle) * radius, cy + Math.sin(angle) * radius])
			}

			polygon(path, points, SILVER)

			if (shape.fill > 0) {
				path.save()
				path.beginPath()
				path.rect(x, y, s * (shape.fill >= 1 ? 1 : shape.fill), s)
				path.clip()
				polygon(path, points, YELLOW)
				path.restore()
			}
			break
		}
		case "boxes": {
			const half = s * 0.38
			const gap = s * 0.06
			const order: [number, number][] = [
				[0, 1],
				[1, 1],
				[0, 0],
				[1, 0]
			]

			order.forEach(([col, row], i) => {
				path.fillStyle = rgb(i < shape.filled ? BLUE : SILVER)
				path.fillRect(x + s * 0.1 + col * (half + gap), y + s * 0.1 + row * (half + gap), half, half)
			})
			break
		}
	}
}

function paintIcon(canvas: GridCanvas, cell: GridDecoration, scene: GridPaintScene): void {
	const layer = (cell.tile as SheetTile).layers?.cfIcons

	if (layer === undefined) return

	const k = indexOf(layer, cell.entry)

	if (k < 0) return

	const shape = iconShape(layer.sets[k] ?? 0, layer.icons[k] ?? 0)

	if (shape === null) return

	const { box } = cell
	const size = Math.min(ICON * scene.scale, box.height - 2, box.width - 2)

	if (size <= 2) return

	// The icon follows the cell's vertical alignment (bottom unless it says otherwise).
	const valign = scene.model.style(cell.tile.styles[cell.entry] ?? 0).valign
	const y =
		valign === "top"
			? box.y + 1
			: valign === "center" || valign === "justify" || valign === "distributed"
				? box.y + (box.height - 1 - size) / 2
				: box.y + box.height - 2 - size

	drawIcon(canvas, shape, box.x + INSET * scene.scale, y, size)
}

export const CF_BAR_DECORATOR: SheetCellDecorator = { id: "cfBars", mask: CellFlag.CF_BAR, phase: "under", paint: paintBar }
// The cell's text lays out after the icon, as Excel draws it.
export const CF_ICON_DECORATOR: SheetCellDecorator = {
	id: "cfIcons",
	mask: CellFlag.CF_ICON,
	phase: "over",
	inset: ICON + INSET,
	paint: paintIcon
}

// The grid extension that draws conditional formatting's bars and icons.
export const CONDITIONAL_FORMATTING_EXTENSION: SheetGridExtension = {
	id: "conditionalFormatting",
	decorators: [CF_BAR_DECORATOR, CF_ICON_DECORATOR]
}
