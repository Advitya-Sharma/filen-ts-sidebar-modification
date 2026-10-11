// Paged editing (a separate entry: the viewer does not load ProseMirror's view).
export { DocxEditor, type ClientPoint, type DocxEditorHandle, type DocxEditorProps } from "./DocxEditor"
export { EditSink, plainText, type EditSinkOptions } from "./edit/sink"
export { EditSurface, type DocxEditorMark } from "./edit/surface"
export { useDocxEditor, type DocxEditorControls, type DocxEditorState } from "./useDocxEditor"
