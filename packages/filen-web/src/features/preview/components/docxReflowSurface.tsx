import type { RefObject } from "react"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxEditorHandle, DocxEditorMark } from "@filen/office/react/docx/editor"
import type { DocxToolbarSlotProps } from "@/features/preview/components/docxEditor.logic"

export interface DocxReflowSurfaceProps {
	editor: PagedEditor
	// The same handle the paged editor fills (clipboard, hit tests, touch selection).
	handleRef: RefObject<DocxEditorHandle | null>
	label: string
	readOnly: boolean
	apple: boolean
	marks: readonly DocxEditorMark[]
	onFind: () => void
}

// The document reflowed to the window's width for editing on phones and narrow windows, in place of the
// pages. Never shown yet: the toggle below is not offered.
export function DocxReflowSurface(_props: DocxReflowSurfaceProps): null {
	return null
}

export interface DocxReflowToggleProps extends DocxToolbarSlotProps {
	reflow: boolean
	onReflowChange: (reflow: boolean) => void
}

// The toolbar's switch between pages and the reflowed view.
export function DocxReflowToggle(_props: DocxReflowToggleProps): null {
	return null
}
