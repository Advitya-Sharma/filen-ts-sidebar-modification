import type { DocxEditorCommand } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxObjectPress } from "@filen/office/react/docx/editor"

export interface DocxObjectControlProps {
	editor: PagedEditor
	// The scrolling element the pages are in (popovers anchor in it).
	container: HTMLDivElement | null
	zoom: number
	readOnly: boolean
	execute: (command: DocxEditorCommand) => void
	// The last press on an object of a feature offered (DocxEditor `interactions`), until dismissed.
	press: DocxObjectPress | null
	onPressDone: () => void
}

// Content controls in the document: checkbox clicks, the dropdown and date pickers. Nothing until form
// controls can be edited.
export function DocxFormControls(_props: DocxObjectControlProps): null {
	return null
}
