// What builds without DOCX editing (VITE_DOCX_EDITING unset) resolve docxEditor.tsx to (vite.config.ts), so the
// editor and everything it pulls in stay out of them. Never rendered there: DOCX_EDITING is false.
export function DocxEditView(): null {
	return null
}
