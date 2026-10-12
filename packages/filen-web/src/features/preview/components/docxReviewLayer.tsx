import type { DocxEditorCommand } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxMarkupMode } from "@filen/office/dom/docx"
import type { DocxDocument } from "@filen/office/react/docx"
import type { DocxObjectPress, DocxPageOverlay } from "@filen/office/react/docx/editor"
import type { DocxAuthorGate } from "@/features/preview/components/docxEditor.logic"

export interface DocxReviewLayerProps {
	editor: PagedEditor
	document: DocxDocument
	markup: DocxMarkupMode
	zoom: number
	// Space between pages (px at 100 %).
	gap: number
	// The scrolling element the pages are in.
	container: HTMLDivElement | null
	readOnly: boolean
	execute: (command: DocxEditorCommand) => void
	withAuthor: DocxAuthorGate
	// The last press on a comment marker, until dismissed.
	press: DocxObjectPress | null
	onPressDone: () => void
}

// Comment balloons in the margin beside their anchors, composed and edited in place (an inline field in the
// balloon, a reply field under the thread), and change bars. Nothing yet: the editor shows no annotations.
export function DocxReviewLayer(_props: DocxReviewLayerProps): null {
	return null
}

export interface DocxReviewPageOverlayProps {
	page: DocxPageOverlay
	editor: PagedEditor
	markup: DocxMarkupMode
}

// What the review draws on a page (the anchors of comments being worked on).
export function DocxReviewPageOverlay(_props: DocxReviewPageOverlayProps): null {
	return null
}
