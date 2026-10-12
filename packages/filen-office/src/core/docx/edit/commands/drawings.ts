import { boolArg, choiceArg, field, intArg, optional, stringArg, type DocxCommandInfo, type DocxCommandParser } from "./base"

// Pictures, shapes, charts and text boxes as objects on the page. Image bytes never travel in a command:
// they are registered with the editor first (addMedia) and named by their handle.

export type DocxWrapMode = "inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront"
// ST_RelFromH / ST_RelFromV.
export type DocxRelativeH = "margin" | "page" | "column" | "character" | "leftMargin" | "rightMargin" | "insideMargin" | "outsideMargin"
export type DocxRelativeV = "margin" | "page" | "paragraph" | "line" | "topMargin" | "bottomMargin" | "insideMargin" | "outsideMargin"

export type DrawingCommand =
	| { readonly type: "image.insert"; readonly media: number }
	// EMUs.
	| { readonly type: "drawing.resize"; readonly cx: number; readonly cy: number; readonly keepRatio?: boolean }
	| {
			readonly type: "drawing.position"
			readonly x: number
			readonly y: number
			readonly relH?: DocxRelativeH
			readonly relV?: DocxRelativeV
	  }
	| { readonly type: "drawing.wrap"; readonly mode: DocxWrapMode }
	| { readonly type: "drawing.altText"; readonly text: string }

const object = { write: true, scope: "document", group: "drawing", context: "drawing" } as const

export const DRAWING_COMMANDS: readonly DocxCommandInfo<DrawingCommand["type"]>[] = [
	{
		type: "image.insert",
		labelKey: "docx.image.insert",
		icon: "image",
		write: true,
		scope: "document",
		group: "insert",
		input: "image",
		async: true
	},
	{ type: "drawing.resize", labelKey: "docx.drawing.resize", icon: "resize", ...object },
	{ type: "drawing.position", labelKey: "docx.drawing.position", icon: "move", ...object },
	{ type: "drawing.wrap", labelKey: "docx.drawing.wrap", icon: "wrap", input: "choice", ...object },
	{ type: "drawing.altText", labelKey: "docx.drawing.altText", icon: "alt-text", input: "text", ...object }
]

const WRAPS = new Set<string>(["inline", "square", "tight", "through", "topAndBottom", "behind", "inFront"])
const RELATIVE_H = new Set<string>(["margin", "page", "column", "character", "leftMargin", "rightMargin", "insideMargin", "outsideMargin"])
const RELATIVE_V = new Set<string>(["margin", "page", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin", "outsideMargin"])
// Extents and offsets are ST_PositiveCoordinate / ST_Coordinate32 (EMUs).
const MAX_EMU = 2147483647
const MAX_ALT_TEXT = 10_000

export const parseDrawingCommand: DocxCommandParser<DrawingCommand> = (type, value) => {
	switch (type) {
		case "image.insert": {
			const media = intArg(value, "media", 1, Number.MAX_SAFE_INTEGER)

			return media === undefined ? undefined : { type, media }
		}
		case "drawing.resize": {
			const cx = intArg(value, "cx", 1, MAX_EMU)
			const cy = intArg(value, "cy", 1, MAX_EMU)
			const keepRatio = optional(value, "keepRatio", boolArg)

			if (cx === undefined || cy === undefined || !keepRatio.ok) return undefined

			return { type, cx, cy, ...(keepRatio.value === undefined ? {} : { keepRatio: keepRatio.value }) }
		}
		case "drawing.position": {
			const x = intArg(value, "x", -MAX_EMU, MAX_EMU)
			const y = intArg(value, "y", -MAX_EMU, MAX_EMU)
			const relH = optional(value, "relH", (holder, key) => choiceArg<DocxRelativeH>(holder, key, RELATIVE_H))
			const relV = optional(value, "relV", (holder, key) => choiceArg<DocxRelativeV>(holder, key, RELATIVE_V))

			if (x === undefined || y === undefined || !relH.ok || !relV.ok) return undefined

			return {
				type,
				x,
				y,
				...(relH.value === undefined ? {} : { relH: relH.value }),
				...(relV.value === undefined ? {} : { relV: relV.value })
			}
		}
		case "drawing.wrap": {
			const mode = choiceArg<DocxWrapMode>(value, "mode", WRAPS)

			return mode === undefined ? undefined : { type, mode }
		}
		case "drawing.altText": {
			const text = field(value, "text") === undefined ? undefined : stringArg(value, "text", MAX_ALT_TEXT)

			return text === undefined ? undefined : { type, text }
		}
	}
}
