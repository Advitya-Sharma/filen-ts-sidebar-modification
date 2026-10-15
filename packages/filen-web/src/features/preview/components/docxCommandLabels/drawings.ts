import type { DocxCommandLabels, DocxCommandVariant } from "@/features/preview/components/docxEditor.logic"

// Pictures: inserted from the Insert menu, the clipboard or a drop (docxImageInsert.tsx); resized, moved and
// cropped with their handles on the pages (docxObjectHandles.tsx); the rest from the Picture group
// (docxDrawingMenu.tsx) and the context menu.
export const DRAWINGS_LABELS: Required<DocxCommandLabels<"drawings">> = {
	"image.insert": "previewDocxInsertPicture",
	"drawing.resize": "previewDocxPictureResize",
	"drawing.position": "previewDocxPictureMove",
	"drawing.wrap": "previewDocxPictureWrap",
	"drawing.altText": "previewDocxPictureAltText",
	"drawing.rotate": "previewDocxPictureRotate",
	"drawing.crop": "previewDocxPictureCrop",
	"drawing.delete": "previewDocxPictureDelete"
}

// Wrap text ▸ each way, in Word's order.
export const WRAP_VARIANTS: readonly DocxCommandVariant[] = [
	{ label: "previewDocxWrapInline", command: { type: "drawing.wrap", mode: "inline" } },
	{ label: "previewDocxWrapSquare", command: { type: "drawing.wrap", mode: "square" } },
	{ label: "previewDocxWrapTight", command: { type: "drawing.wrap", mode: "tight" } },
	{ label: "previewDocxWrapThrough", command: { type: "drawing.wrap", mode: "through" } },
	{ label: "previewDocxWrapTopAndBottom", command: { type: "drawing.wrap", mode: "topAndBottom" } },
	{ label: "previewDocxWrapBehind", command: { type: "drawing.wrap", mode: "behind" } },
	{ label: "previewDocxWrapInFront", command: { type: "drawing.wrap", mode: "inFront" } }
]

// Rotate ▸ each turn.
export const ROTATE_VARIANTS: readonly DocxCommandVariant[] = [
	{ label: "previewDocxRotateRight", command: { type: "drawing.rotate", by: 90 } },
	{ label: "previewDocxRotateLeft", command: { type: "drawing.rotate", by: 270 } },
	{ label: "previewDocxRotateHalf", command: { type: "drawing.rotate", by: 180 } }
]
