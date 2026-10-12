// Paged editing (a separate entry: the viewer does not load ProseMirror's view).
export {
	DocxEditor,
	type ClientPoint,
	type DocxEditorHandle,
	type DocxEditorInteractions,
	type DocxEditorProps,
	type DocxObjectPress,
	type DocxPageOverlay
} from "./DocxEditor"
export { EditSink, plainText, type EditSinkOptions } from "./edit/sink"
export { EditSurface, type DocxEditorMark } from "./edit/surface"
export { useDocxEditor, type DocxEditorControls, type DocxEditorState } from "./useDocxEditor"
