// The DOCX editor's commands as serialisable data, with their metadata: an app builds its own toolbar,
// menus and shortcuts from the table and the state snapshot, and sends commands back (directly on web,
// over a worker or the WebView bridge on mobile). Untrusted input goes through parseDocxEditorCommand.

export const DOCX_PARAGRAPH_STYLES = [
	"normal",
	"title",
	"subtitle",
	"heading1",
	"heading2",
	"heading3",
	"heading4",
	"heading5",
	"heading6",
	"heading7",
	"heading8",
	"heading9",
	"quote"
] as const

export type DocxParagraphStyle = (typeof DOCX_PARAGRAPH_STYLES)[number]

// The built-in name each style has in every document, whatever the UI language (style ids are localised).
export const PARAGRAPH_STYLE_NAMES: Readonly<Record<DocxParagraphStyle, string>> = {
	normal: "normal",
	title: "title",
	subtitle: "subtitle",
	heading1: "heading 1",
	heading2: "heading 2",
	heading3: "heading 3",
	heading4: "heading 4",
	heading5: "heading 5",
	heading6: "heading 6",
	heading7: "heading 7",
	heading8: "heading 8",
	heading9: "heading 9",
	quote: "quote"
}

export type DocxAlignment = "start" | "center" | "end" | "justify"

export type DocxEditorCommand =
	| { readonly type: "format.bold" }
	| { readonly type: "format.italic" }
	| { readonly type: "format.underline" }
	| { readonly type: "format.strike" }
	| { readonly type: "format.clear" }
	| { readonly type: "style.set"; readonly style: DocxParagraphStyle }
	| { readonly type: "align.start" }
	| { readonly type: "align.center" }
	| { readonly type: "align.end" }
	| { readonly type: "align.justify" }
	| { readonly type: "list.bullet" }
	| { readonly type: "list.numbered" }
	| { readonly type: "list.indent" }
	| { readonly type: "list.outdent" }
	| { readonly type: "history.undo" }
	| { readonly type: "history.redo" }

export type DocxEditorCommandType = DocxEditorCommand["type"]

export type DocxEditorCommandGroup = "format" | "style" | "align" | "list" | "history"

export interface DocxEditorCommandInfo {
	readonly type: DocxEditorCommandType
	// A stable id an app maps to its own translation.
	readonly labelKey: string
	// A neutral icon id an app maps to its own icon set; absent for commands without one.
	readonly icon?: string
	// Platform-neutral shortcut ("mod" = Cmd on Apple platforms, Ctrl elsewhere).
	readonly shortcut?: string
	readonly group: DocxEditorCommandGroup
	// Changes the document (refused in read-only mode).
	readonly write: boolean
}

export const DOCX_EDITOR_COMMANDS: readonly DocxEditorCommandInfo[] = [
	{ type: "format.bold", labelKey: "docx.format.bold", icon: "bold", shortcut: "mod+b", group: "format", write: true },
	{ type: "format.italic", labelKey: "docx.format.italic", icon: "italic", shortcut: "mod+i", group: "format", write: true },
	{ type: "format.underline", labelKey: "docx.format.underline", icon: "underline", shortcut: "mod+u", group: "format", write: true },
	{ type: "format.strike", labelKey: "docx.format.strike", icon: "strikethrough", group: "format", write: true },
	{ type: "format.clear", labelKey: "docx.format.clear", icon: "clear-formatting", shortcut: "mod+space", group: "format", write: true },
	{ type: "style.set", labelKey: "docx.style.set", icon: "paragraph-style", group: "style", write: true },
	{ type: "align.start", labelKey: "docx.align.start", icon: "align-start", shortcut: "mod+l", group: "align", write: true },
	{ type: "align.center", labelKey: "docx.align.center", icon: "align-center", shortcut: "mod+e", group: "align", write: true },
	{ type: "align.end", labelKey: "docx.align.end", icon: "align-end", shortcut: "mod+r", group: "align", write: true },
	{ type: "align.justify", labelKey: "docx.align.justify", icon: "align-justify", shortcut: "mod+j", group: "align", write: true },
	{ type: "list.bullet", labelKey: "docx.list.bullet", icon: "list-bullet", shortcut: "mod+shift+8", group: "list", write: true },
	{ type: "list.numbered", labelKey: "docx.list.numbered", icon: "list-numbered", shortcut: "mod+shift+7", group: "list", write: true },
	{ type: "list.indent", labelKey: "docx.list.indent", icon: "indent", group: "list", write: true },
	{ type: "list.outdent", labelKey: "docx.list.outdent", icon: "outdent", group: "list", write: true },
	{ type: "history.undo", labelKey: "docx.history.undo", icon: "undo", shortcut: "mod+z", group: "history", write: true },
	{ type: "history.redo", labelKey: "docx.history.redo", icon: "redo", shortcut: "mod+shift+z", group: "history", write: true }
]

export interface DocxParagraphStyleInfo {
	readonly style: DocxParagraphStyle
	readonly labelKey: string
	readonly shortcut?: string
}

// The paragraph styles `style.set` offers, in menu order.
export const DOCX_PARAGRAPH_STYLE_MENU: readonly DocxParagraphStyleInfo[] = DOCX_PARAGRAPH_STYLES.map(style => {
	const level = /^heading([1-3])$/.exec(style)?.[1]

	return {
		style,
		labelKey: `docx.style.${style}`,
		...(level !== undefined ? { shortcut: `mod+alt+${level}` } : style === "normal" ? { shortcut: "mod+shift+n" } : {})
	}
})

const COMMAND_TYPES = new Set<string>(DOCX_EDITOR_COMMANDS.map(command => command.type))
const STYLES = new Set<string>(DOCX_PARAGRAPH_STYLES)

// Untrusted input (a bridge message) to a command, or undefined.
export function parseDocxEditorCommand(value: unknown): DocxEditorCommand | undefined {
	if (typeof value !== "object" || value === null) return undefined

	const type: unknown = Reflect.get(value, "type")

	if (typeof type !== "string" || !COMMAND_TYPES.has(type)) return undefined

	if (type === "style.set") {
		const style: unknown = Reflect.get(value, "style")

		return typeof style === "string" && STYLES.has(style) ? { type, style: style as DocxParagraphStyle } : undefined
	}

	return { type } as DocxEditorCommand
}

export function isWriteCommand(type: DocxEditorCommandType): boolean {
	return DOCX_EDITOR_COMMANDS.find(command => command.type === type)?.write ?? true
}
