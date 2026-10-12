import type { DocxWrapMode } from "../commands/drawings"
import type { SnapshotInput } from "./index"

// The object the selection is (a picture, shape, chart or text box).

export interface DocxDrawingContext {
	readonly kind: "picture" | "shape" | "chart" | "group" | "textbox" | "object"
	readonly anchored: boolean
	readonly wrap: DocxWrapMode
	// EMUs.
	readonly cx: number
	readonly cy: number
	readonly resizable: boolean
	readonly altText: string
}

export function drawingOf(_input: SnapshotInput): DocxDrawingContext | null {
	return null
}
