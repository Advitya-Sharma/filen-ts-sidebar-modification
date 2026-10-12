import type { Node } from "prosemirror-model"
import { NodeSelection, TextSelection } from "prosemirror-state"
import { changeDrawing, drawingState, type DrawingChanges } from "../../../core/docx/objects/edit"
import { docxSchema, type DocxEditorCommand, type DocxEditorSession } from "../harness"
import { textblocks } from "./base"
import type { EditOp, OpContext, Rand } from "./index"

// Pictures, shapes, charts and text boxes as objects: resize, move, wrap, rotate, describe and delete the
// document's drawings and insert new pictures, through the drawing commands under the paged editor (whose
// pages must stay those of the saved file); overrides set raw (extreme values included) and new pictures
// inserted raw for the writer and Gate 2; a resize and a new picture in the scripted suite.

const nodes = docxSchema.nodes
const WRAPS = ["inline", "square", "tight", "through", "topAndBottom", "behind", "inFront"] as const
const RELATIVE_H = ["margin", "page", "column", "character"] as const
const RELATIVE_V = ["margin", "page", "paragraph", "line"] as const
const TEXTS = ["A chart of results", "", "Line one\nline two", 'Quotes "and" <brackets> & ampersands', "Kraków — naïve"]
const EMU_PER_INCH = 914400

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

function u32(value: number): number[] {
	return [(value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255]
}

function chunk(type: string, data: readonly number[]): number[] {
	return [...u32(data.length), ...[...type].map(c => c.charCodeAt(0)), ...data, 0, 0, 0, 0]
}

// A PNG header of a width x height picture (what placing it needs; the pixels are not read).
export function pngHeader(width: number, height: number): Uint8Array {
	return new Uint8Array([
		0x89,
		0x50,
		0x4e,
		0x47,
		0x0d,
		0x0a,
		0x1a,
		0x0a,
		...chunk("IHDR", [...u32(width), ...u32(height), 8, 6, 0, 0, 0]),
		...chunk("IEND", [])
	])
}

// The drawings of a document with their positions.
function drawings(doc: Node): { node: Node; pos: number }[] {
	const out: { node: Node; pos: number }[] = []

	doc.descendants((node, pos) => {
		if (node.type === nodes.drawing) out.push({ node, pos })

		return !node.isLeaf && node.type !== nodes.locked_block
	})

	return out
}

function selectDrawing(session: DocxEditorSession, pos: number): boolean {
	try {
		return session.dispatch(session.state.tr.setSelection(NodeSelection.create(session.state.doc, pos)))
	} catch {
		return false
	}
}

// Selects a random drawing of the story being edited and runs the command `make` gives for it.
function onDrawing(context: OpContext, make: (node: Node, rand: Rand) => DocxEditorCommand | undefined): string {
	const session = context.session
	const found = drawings(session.state.doc)

	if (found.length === 0) return "noop"

	const target = pick(context.rand, found)

	if (!selectDrawing(session, target.pos)) return "noop"

	const command = make(target.node, context.rand)

	if (command === undefined) return "noop"

	return `${command.type}@${String(target.pos)} -> ${context.execute(command.type, command)}`
}

function resize(node: Node, rand: Rand, session: DocxEditorSession): DocxEditorCommand | undefined {
	const state = drawingState(session.context, node)

	if (!state.resizable) return undefined

	const scale = 0.4 + rand() * 1.2

	return {
		type: "drawing.resize",
		cx: Math.max(1, Math.round(state.cx * scale)),
		cy: Math.max(1, Math.round(state.cy * (rand() < 0.5 ? scale : 0.4 + rand() * 1.2))),
		keepRatio: rand() < 0.5
	}
}

function paged(name: string, weight: number, run: (context: OpContext) => string): EditOp {
	return {
		name,
		harness: "paged",
		weight,
		run(context) {
			return context.editor === undefined ? "noop" : run(context)
		}
	}
}

// A new picture where the caret goes: on a random paragraph, the caret anywhere in it.
function insertPicture(context: OpContext): string {
	const editor = context.editor
	const session = context.session
	const blocks = textblocks(session.state.doc)

	if (editor === undefined || blocks.length === 0) return "noop"

	const block = pick(context.rand, blocks)
	const at = block.pos + 1 + Math.floor(context.rand() * (block.node.content.size + 1))

	try {
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at)))
	} catch {
		return "noop"
	}

	const media = editor.addMedia(pngHeader(16 + Math.floor(context.rand() * 1200), 16 + Math.floor(context.rand() * 900)))

	if (media === undefined) return "noop"

	const command: DocxEditorCommand = { type: "image.insert", media, altText: pick(context.rand, TEXTS) }

	return `image.insert@${String(at)} -> ${context.execute(command.type, command)}`
}

// Overrides set raw on a random drawing: any values the schema allows, sizes from 1 EMU up.
function rawOverrides(context: OpContext): string {
	const session = context.session
	const found = drawings(session.state.doc)

	if (found.length === 0) return "noop"

	const target = pick(context.rand, found)
	const rand = context.rand
	const changes: DrawingChanges = {
		...(rand() < 0.5 ? { cx: rand() < 0.1 ? 1 : Math.round(rand() * 8 * EMU_PER_INCH) + 1 } : {}),
		...(rand() < 0.5 ? { cy: rand() < 0.1 ? 2147483647 : Math.round(rand() * 8 * EMU_PER_INCH) + 1 } : {}),
		...(rand() < 0.3 ? { wrap: pick(rand, WRAPS) } : {}),
		...(rand() < 0.3 ? { posX: Math.round((rand() - 0.5) * 4 * EMU_PER_INCH), relH: pick(rand, RELATIVE_H) } : {}),
		...(rand() < 0.3 ? { posY: Math.round((rand() - 0.5) * 4 * EMU_PER_INCH), relV: pick(rand, RELATIVE_V) } : {}),
		...(rand() < 0.3 ? { descr: pick(rand, TEXTS) } : {}),
		...(rand() < 0.3 ? { rot: pick(rand, [0, 90, 180, 270]) } : {})
	}
	const tr = changeDrawing(session.state, session.context, target.pos, target.node, changes)

	return tr === null ? "noop" : `drawingRaw@${String(target.pos)} ${JSON.stringify(changes)} -> ${String(session.dispatch(tr))}`
}

// A new picture node, at the end of the last paragraph (`last`) or anywhere in a random one.
function newPicture(session: DocxEditorSession, rand: Rand | undefined): string {
	const blocks = textblocks(session.state.doc)

	if (blocks.length === 0) return "noop"

	const block = rand === undefined ? (blocks[blocks.length - 1] as { node: Node; pos: number }) : pick(rand, blocks)
	const at = block.pos + 1 + (rand === undefined ? block.node.content.size : Math.floor(rand() * (block.node.content.size + 1)))
	const media = session.context.media.add(pngHeader(96, 48))

	if (media === undefined) return "noop"

	const node = nodes.drawing.create({
		src: -1,
		kind: "drawing",
		index: -1,
		overrides: {
			media,
			cx: 914400,
			cy: 457200,
			docPr: session.context.ids.next("docPr"),
			descr: "Inserted by the edit suite",
			...(rand !== undefined && rand() < 0.3 ? { wrap: "square", posX: 0, posY: 0, relH: "column", relV: "paragraph" } : {})
		}
	})

	try {
		return `pictureRaw@${String(at)} -> ${String(session.dispatch(session.state.tr.insert(at, node)))}`
	} catch {
		return "noop"
	}
}

export const drawingOps: readonly EditOp[] = [
	paged("drawingResize", 3, context => onDrawing(context, (node, rand) => resize(node, rand, context.session))),
	paged("drawingWrap", 2, context => onDrawing(context, (_node, rand) => ({ type: "drawing.wrap", mode: pick(rand, WRAPS) }))),
	paged("drawingPosition", 2, context =>
		onDrawing(context, (node, rand) =>
			drawingState(context.session.context, node).wrap === "inline"
				? undefined
				: {
						type: "drawing.position",
						x: Math.round((rand() - 0.3) * 3 * EMU_PER_INCH),
						y: Math.round((rand() - 0.3) * 3 * EMU_PER_INCH),
						...(rand() < 0.3 ? { relH: pick(rand, RELATIVE_H) } : {}),
						...(rand() < 0.3 ? { relV: pick(rand, RELATIVE_V) } : {})
					}
		)
	),
	paged("drawingRotate", 1, context =>
		onDrawing(context, (_node, rand) => ({ type: "drawing.rotate", by: pick(rand, [90, 180, 270] as const) }))
	),
	paged("drawingAltText", 1, context => onDrawing(context, (_node, rand) => ({ type: "drawing.altText", text: pick(rand, TEXTS) }))),
	paged("drawingDelete", 1, context => onDrawing(context, () => ({ type: "drawing.delete" }))),
	paged("imageInsert", 2, insertPicture),
	{ name: "drawingRaw", harness: "damage", weight: 3, run: rawOverrides },
	{ name: "pictureRaw", harness: "damage", weight: 1, run: context => newPicture(context.session, context.rand) },
	{
		name: "drawingResizeScripted",
		harness: "damage",
		weight: 0,
		run: rawOverrides,
		script: {
			order: 700,
			run(target) {
				const session = target.session
				const found = drawings(session.state.doc).find(entry => drawingState(session.context, entry.node).resizable)

				if (found === undefined) return false

				const state = drawingState(session.context, found.node)
				const tr = changeDrawing(session.state, session.context, found.pos, found.node, {
					cx: Math.max(1, Math.round(state.cx * 0.9)),
					cy: Math.max(1, Math.round(state.cy * 0.9)),
					descr: "Resized by the edit suite"
				})

				return tr !== null && session.dispatch(tr)
			}
		}
	},
	{
		name: "pictureInsertScripted",
		harness: "damage",
		weight: 0,
		run: context => newPicture(context.session, context.rand),
		writes: ["image/png"],
		script: {
			order: 710,
			run(target) {
				return newPicture(target.session, undefined).endsWith("true")
			}
		}
	}
]
