import type { DocxEditorCommand } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxObjectPress, DocxPageOverlay } from "@filen/office/react/docx/editor"

export interface DocxObjectHandlesProps {
	// The page drawn over, in its own coordinates (DocxEditor `overlays`).
	page: DocxPageOverlay
	editor: PagedEditor
	press: DocxObjectPress | null
	readOnly: boolean
	execute: (command: DocxEditorCommand) => void
}

// The selection frame and resize handles of the drawing pressed on this page. Nothing until drawings can be
// edited.
export function DocxObjectHandles(_props: DocxObjectHandlesProps): null {
	return null
}
