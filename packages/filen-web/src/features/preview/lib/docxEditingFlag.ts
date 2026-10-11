// DOCX editing is internal until the editor is complete (formatting, lists, tables, images, comments, tracked
// changes, headers/footers/notes, fields, paged editing): off in every build unless VITE_DOCX_EDITING=1. Read
// by dot access so the build replaces it with a constant and drops the editor from builds without it.
export const DOCX_EDITING: boolean = import.meta.env.VITE_DOCX_EDITING === "1"
