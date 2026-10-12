import { field, hasControl, stringArg, type DocxCommandInfo, type DocxCommandParser } from "./base"

// Links and breaks.

// A link goes to a URL (`target`, written as an external relationship, never fetched) or to a bookmark in
// the document (`anchor`); `text` is inserted as the link's text when the selection is empty.
export type InsertCommand =
	| { readonly type: "link.set"; readonly target: string; readonly text?: string }
	| { readonly type: "link.set"; readonly anchor: string; readonly text?: string }
	| { readonly type: "link.remove" }
	| { readonly type: "insert.pageBreak" }

export const INSERT_COMMANDS: readonly DocxCommandInfo<InsertCommand["type"]>[] = [
	{ type: "link.set", labelKey: "docx.link.set", icon: "link", group: "link", write: true, scope: "story", input: "link" },
	{ type: "link.remove", labelKey: "docx.link.remove", icon: "unlink", group: "link", write: true, scope: "story" },
	{ type: "insert.pageBreak", labelKey: "docx.insert.pageBreak", icon: "page-break", group: "insert", write: true, scope: "story" }
]

// URLs a link may hold (as typed; what opening one allows is the app's link policy).
const MAX_TARGET = 2048
// Word's bookmark names.
const BOOKMARK_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,39}$/
const MAX_LINK_TEXT = 1024

export const parseInsertCommand: DocxCommandParser<InsertCommand> = (type, value) => {
	if (type !== "link.set") return { type } as InsertCommand

	const hasText = field(value, "text") !== undefined
	const text = stringArg(value, "text", MAX_LINK_TEXT, 1)

	if (hasText && text === undefined) return undefined

	const extra = text === undefined ? {} : { text }
	const target = stringArg(value, "target", MAX_TARGET, 1)
	const anchor = field(value, "anchor")

	if (target !== undefined && anchor === undefined && !hasControl(target)) return { type, target, ...extra }

	if (target === undefined && field(value, "target") === undefined && typeof anchor === "string" && BOOKMARK_NAME.test(anchor))
		return { type, anchor, ...extra }

	return undefined
}
