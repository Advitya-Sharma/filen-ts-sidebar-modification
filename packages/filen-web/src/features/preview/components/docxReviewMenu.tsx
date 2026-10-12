import type { DocxMarkupMode } from "@filen/office/dom/docx"
import type { DocxToolbarSlotProps } from "@/features/preview/components/docxEditor.logic"

export interface DocxReviewMenuProps extends DocxToolbarSlotProps {
	// Null for a document without comments or tracked changes.
	markup: DocxMarkupMode | null
	onMarkupChange: (mode: DocxMarkupMode) => void
}

// The toolbar's one Review menu: the track changes switch, the markup modes, previous and next, accept all and
// reject all (it replaces the markup menu once offered). Nothing until its commands are labelled.
export function DocxReviewMenu(_props: DocxReviewMenuProps): null {
	return null
}
