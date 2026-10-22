import type { DocxPictureRefusal } from "@/features/preview/lib/docxImages.logic"
import type { SheetLayout } from "@filen/office/core/sheet"
import type { SheetDrawingObject } from "@filen/office/react/sheet/drawings"
import type { PreviewKey } from "@/lib/i18n"

// Drawing objects in the sheet view: whether the sheet lets its objects be edited, what a picture refusal says, and
// what an object is called for assistive technology.

// Sheet protection that does not allow "Edit objects" (§18.3.1.85 objects) locks every object of the sheet.
export function objectsEditable(layout: SheetLayout | null | undefined): boolean {
	return layout?.protection === null || layout?.protection === undefined || layout.protection.allows.has("objects")
}

const PICTURE_REFUSALS: Readonly<Record<DocxPictureRefusal, PreviewKey>> = {
	tooLarge: "previewSheetObjectTooLarge",
	pixels: "previewSheetObjectPixels",
	format: "previewSheetObjectFormat",
	heic: "previewSheetObjectHeic",
	avif: "previewSheetObjectAvif",
	corrupt: "previewSheetObjectCorrupt",
	// A workbook takes every picture: the document limit never applies.
	full: "previewSheetObjectFailed",
	place: "previewSheetObjectFailed",
	failed: "previewSheetObjectFailed"
}

export function pictureRefusalKey(refusal: DocxPictureRefusal): PreviewKey {
	return PICTURE_REFUSALS[refusal]
}

// The accessible name: the object's alternative text, else its name in the file ("Picture 1"), else what it is.
export function objectLabel(object: SheetDrawingObject, t: (key: PreviewKey) => string): string {
	for (const name of [object.alt, object.name]) {
		if (name.trim() !== "") {
			return name
		}
	}

	switch (object.kind) {
		case "picture":
			return t("previewSheetObjectPicture")
		case "shape":
		case "connector":
		case "group":
			return t("previewSheetObjectShape")
		case "chart":
		case "chartex":
			return t("previewSheetObjectChart")
		default:
			return t("previewSheetObjectOther")
	}
}
