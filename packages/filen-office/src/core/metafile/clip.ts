import { IDENTITY, PathCommand, type DrawingClipShape, type DrawingList } from "../drawingml/displayList"
import { MAX_CLIP_DEPTH, MAX_CLIP_WEIGHT, MetafileLimit } from "./limits"
import { combineRects, normalizeRects, translateRects, type Charge, type RectOp } from "./region"

// Clip regions of GDI and GDI+ playback, exact under every combination: unions of shapes (each with its own
// fill rule and placement), chains of intersections, rectangle regions with exact band algebra, and boolean
// regions (union minus union) for the rest. Every coordinate a combination touches is charged to the
// conversion's work budget, and the coordinates a live region retains are bounded.

// Path data in output space, shared by every clip that uses it and written to a list once.
export class ClipGeometry {
	readonly commands: readonly number[]
	readonly coords: readonly number[]
	list: DrawingList | null = null
	firstCommand = 0
	firstCoord = 0

	constructor(commands: readonly number[], coords: readonly number[]) {
		this.commands = commands
		this.coords = coords
	}
}

export interface ClipShape {
	readonly geometry: ClipGeometry
	readonly evenOdd: boolean
	readonly matrix: readonly number[]
}

export interface ClipEnv {
	readonly charge: Charge
	// A rectangle this far around the origin stands for the whole plane.
	readonly extent: number
}

export const ClipCombine = { AND: 1, OR: 2, XOR: 3, DIFF: 4, COPY: 5, COMPLEMENT: 6 } as const

// A region: the union of `shapes` (or the rectangles `rects`, or `add` minus `subtract`), intersected with
// `parent`. A node with nothing to add hides everything.
export class ClipNode {
	readonly parent: ClipNode | null
	readonly shapes: readonly ClipShape[]
	readonly rects: Float64Array | null
	readonly add: readonly ClipNode[] | null
	readonly subtract: readonly ClipNode[]
	readonly hidesAll: boolean
	readonly depth: number
	// Coordinates the region holds on to, parents and members included.
	readonly weight: number
	// The geometry of `rects`, built when first written.
	rectGeometry: ClipGeometry | null = null
	list: DrawingList | null = null
	index = -1

	constructor(
		parent: ClipNode | null,
		shapes: readonly ClipShape[],
		rects: Float64Array | null,
		add: readonly ClipNode[] | null,
		subtract: readonly ClipNode[]
	) {
		this.parent = parent
		this.shapes = shapes
		this.rects = rects
		this.add = add
		this.subtract = subtract

		let depth = (parent?.depth ?? 0) + 1
		let weight = (parent?.weight ?? 0) + (rects?.length ?? 0)

		for (const shape of shapes) weight += shape.geometry.coords.length

		for (const members of [add ?? [], subtract]) {
			for (const member of members) {
				depth = Math.max(depth, member.depth + 1)
				weight += member.weight
			}
		}

		if (depth > MAX_CLIP_DEPTH) throw new MetafileLimit("clip depth")
		if (weight > MAX_CLIP_WEIGHT) throw new MetafileLimit("clip coordinates")

		this.depth = depth
		this.weight = weight

		const empty = add !== null ? add.every(member => member.hidesAll) : rects !== null ? rects.length === 0 : shapes.length === 0

		this.hidesAll = empty || (parent?.hidesAll ?? false)
	}

	// The node's own region, rebuilt on another parent.
	withParent(parent: ClipNode | null): ClipNode {
		return new ClipNode(parent, this.shapes, this.rects, this.add, this.subtract)
	}
}

export function emptyClip(): ClipNode {
	return new ClipNode(null, [], new Float64Array(0), null, [])
}

// A path region (output-space commands and coordinates, copied).
export function pathClip(commands: readonly number[], coords: readonly number[], evenOdd: boolean, env: ClipEnv): ClipNode {
	env.charge(commands.length + coords.length)

	return new ClipNode(null, [{ geometry: new ClipGeometry(commands.slice(), coords.slice()), evenOdd, matrix: IDENTITY }], null, null, [])
}

// A region of shared geometry.
export function geometryClip(geometry: ClipGeometry, evenOdd: boolean, matrix: readonly number[] = IDENTITY): ClipNode {
	return new ClipNode(null, [{ geometry, evenOdd, matrix }], null, null, [])
}

// Rectangles (left, top, right, bottom) in output space as an exact region.
export function rectsClip(rects: ArrayLike<number>, env: ClipEnv): ClipNode {
	return new ClipNode(null, [], normalizeRects(rects, env.charge), null, [])
}

// A logical rectangle through `m`: an exact rectangle region under an axis-aligned mapping, else a path.
export function rectClip(left: number, top: number, right: number, bottom: number, m: readonly number[], env: ClipEnv): ClipNode {
	const [a = 1, b = 0, c = 0, d = 1, e = 0, f = 0] = m
	const x = (px: number, py: number): number => a * px + c * py + e
	const y = (px: number, py: number): number => b * px + d * py + f

	if (b === 0 && c === 0) return rectsClip([x(left, top), y(left, top), x(right, bottom), y(right, bottom)], env)

	return pathClip(
		[PathCommand.MOVE, PathCommand.LINE, PathCommand.LINE, PathCommand.LINE, PathCommand.CLOSE],
		[x(left, top), y(left, top), x(right, top), y(right, top), x(right, bottom), y(right, bottom), x(left, bottom), y(left, bottom)],
		false,
		env
	)
}

function whole(env: ClipEnv): ClipNode {
	const extent = env.extent

	return rectsClip([-extent, -extent, extent, extent], env)
}

// A region exactly representable as rectangles and nothing else.
function pureRects(node: ClipNode | null): node is ClipNode & { rects: Float64Array } {
	return node !== null && node.rects !== null && node.parent === null
}

// `node`'s chain rebuilt on top of `base` (intersection).
function rebase(node: ClipNode, base: ClipNode, env: ClipEnv): ClipNode {
	const chain: ClipNode[] = []

	for (let at: ClipNode | null = node; at !== null; at = at.parent) chain.push(at)

	env.charge(chain.length)

	let out = base

	for (let k = chain.length - 1; k >= 0; k--) out = chain[k]?.withParent(out) ?? out

	return out
}

export function intersectClips(a: ClipNode | null, b: ClipNode | null, env: ClipEnv): ClipNode | null {
	if (a === null) return b
	if (b === null) return a
	if (pureRects(a) && pureRects(b)) return new ClipNode(null, [], combineRects(a.rects, b.rects, "and", env.charge), null, [])

	return rebase(b, a, env)
}

function rectOp(mode: number): RectOp | undefined {
	switch (mode) {
		case ClipCombine.AND:
			return "and"
		case ClipCombine.OR:
			return "or"
		case ClipCombine.XOR:
			return "xor"
		case ClipCombine.DIFF:
			return "diff"
		default:
			return undefined
	}
}

// RGN_* / CombineMode: AND (intersect), OR (union), XOR, DIFF (current minus region), COPY (replace),
// COMPLEMENT (region minus current). `null` is the whole plane.
export function combineClips(current: ClipNode | null, region: ClipNode | null, mode: number, env: ClipEnv): ClipNode | null {
	if (mode === ClipCombine.COPY) return region
	if (mode === ClipCombine.AND) return intersectClips(current, region, env)

	const left = mode === ClipCombine.COMPLEMENT ? region : current
	const right = mode === ClipCombine.COMPLEMENT ? current : region
	const op = rectOp(mode === ClipCombine.COMPLEMENT ? ClipCombine.DIFF : mode)

	if (op === undefined) return current

	// The whole plane stands in as a rectangle far beyond the picture, so every combination stays exact.
	const a = left ?? whole(env)
	const b = right ?? whole(env)

	if (op === "or" && (left === null || right === null)) return null

	if (pureRects(a) && pureRects(b)) return new ClipNode(null, [], combineRects(a.rects, b.rects, op, env.charge), null, [])

	switch (op) {
		case "or":
			return new ClipNode(null, [], null, [a, b], [])
		case "diff":
			return new ClipNode(null, [], null, [a], [b])
		default:
			return new ClipNode(null, [], null, [a, b], [intersectClips(a, b, env) ?? whole(env)])
	}
}

// The region moved by (dx, dy) in output space: shapes by their placement, rectangles exactly.
export function offsetClip(node: ClipNode | null, dx: number, dy: number, env: ClipEnv): ClipNode | null {
	const moved = new Map<ClipNode, ClipNode>()
	const move = (at: ClipNode): ClipNode => {
		const done = moved.get(at)

		if (done !== undefined) return done

		env.charge(1 + at.shapes.length)

		const shapes = at.shapes.map(shape => {
			const [a = 1, b = 0, c = 0, d = 1, e = 0, f = 0] = shape.matrix

			return { ...shape, matrix: [a, b, c, d, e + dx, f + dy] }
		})
		const out = new ClipNode(
			at.parent === null ? null : move(at.parent),
			shapes,
			at.rects === null ? null : translateRects(at.rects, dx, dy, env.charge),
			at.add === null ? null : at.add.map(move),
			at.subtract.map(move)
		)

		moved.set(at, out)

		return out
	}

	return node === null ? null : move(node)
}

function writeGeometry(list: DrawingList, geometry: ClipGeometry, reserve: (coords: number) => void): void {
	if (geometry.list === list) return

	reserve(geometry.coords.length)
	geometry.list = list
	geometry.firstCommand = list.commandCount
	geometry.firstCoord = list.coordCount

	const coords = geometry.coords
	let at = 0

	for (const command of geometry.commands) {
		switch (command) {
			case PathCommand.MOVE:
				list.moveTo(coords[at] ?? 0, coords[at + 1] ?? 0)
				at += 2
				break
			case PathCommand.LINE:
				list.lineTo(coords[at] ?? 0, coords[at + 1] ?? 0)
				at += 2
				break
			case PathCommand.CUBIC:
				list.cubicTo(
					coords[at] ?? 0,
					coords[at + 1] ?? 0,
					coords[at + 2] ?? 0,
					coords[at + 3] ?? 0,
					coords[at + 4] ?? 0,
					coords[at + 5] ?? 0
				)
				at += 6
				break
			case PathCommand.QUAD:
				list.quadTo(coords[at] ?? 0, coords[at + 1] ?? 0, coords[at + 2] ?? 0, coords[at + 3] ?? 0)
				at += 4
				break
			default:
				list.close()
				break
		}
	}
}

function rectGeometry(node: ClipNode & { rects: Float64Array }): ClipGeometry {
	if (node.rectGeometry === null) {
		const commands: number[] = []
		const coords: number[] = []
		const rects = node.rects

		for (let k = 0; k + 3 < rects.length; k += 4) {
			const x0 = rects[k] ?? 0
			const y0 = rects[k + 1] ?? 0
			const x1 = rects[k + 2] ?? 0
			const y1 = rects[k + 3] ?? 0

			commands.push(PathCommand.MOVE, PathCommand.LINE, PathCommand.LINE, PathCommand.LINE, PathCommand.CLOSE)
			coords.push(x0, y0, x1, y0, x1, y1, x0, y1)
		}

		node.rectGeometry = new ClipGeometry(commands, coords)
	}

	return node.rectGeometry
}

// The list's clip for a node, written with its parents and members on first use.
export function writeClip(list: DrawingList, node: ClipNode, reserve: (coords: number) => void, opacity?: number): number {
	if (node.list === list && opacity === undefined) return node.index

	const parent = node.parent === null ? -1 : writeClip(list, node.parent, reserve)
	let index: number

	if (node.add !== null) {
		const add = node.add.map(member => writeClip(list, member, reserve))
		const subtract = node.subtract.map(member => writeClip(list, member, reserve))

		index = list.clip(IDENTITY, list.commandCount, list.coordCount, false, parent, 0, {
			add,
			subtract,
			...(opacity === undefined ? {} : { opacity })
		})
	} else {
		const shapes: ClipShape[] =
			node.rects !== null
				? [{ geometry: rectGeometry(node as ClipNode & { rects: Float64Array }), evenOdd: false, matrix: IDENTITY }]
				: [...node.shapes]

		for (const shape of shapes) writeGeometry(list, shape.geometry, reserve)

		const placed: DrawingClipShape[] = shapes.map(shape => ({
			matrix: shape.matrix,
			firstCommand: shape.geometry.firstCommand,
			commandCount: shape.geometry.commands.length,
			firstCoord: shape.geometry.firstCoord,
			evenOdd: shape.evenOdd
		}))
		const first = placed[0]

		index = list.clip(
			first?.matrix ?? IDENTITY,
			first?.firstCommand ?? list.commandCount,
			first?.firstCoord ?? list.coordCount,
			first?.evenOdd ?? false,
			parent,
			first?.commandCount ?? 0,
			{ shapes: placed.slice(1), ...(opacity === undefined ? {} : { opacity }) }
		)
	}

	if (opacity === undefined) {
		node.list = list
		node.index = index
	}

	return index
}
