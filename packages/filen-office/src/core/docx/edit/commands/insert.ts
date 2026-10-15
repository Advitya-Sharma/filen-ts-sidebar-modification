import { field, hasControl, intArg, stringArg, type DocxCommandInfo, type DocxCommandParser } from "./base"

// Links, breaks and characters.

// A link goes to a URL (`target`, written as an external relationship, never fetched), to a bookmark in the
// document (`anchor`), or to a heading of the story (`heading`: its paragraph's position), through the hidden
// _Toc bookmark Word's Place in This Document links by (the heading's own, or one made around its text); `text`
// is inserted as the link's text when the selection is empty.
export type InsertCommand =
	| { readonly type: "link.set"; readonly target: string; readonly text?: string }
	| { readonly type: "link.set"; readonly anchor: string; readonly text?: string }
	| { readonly type: "link.set"; readonly heading: number; readonly text?: string }
	| { readonly type: "link.remove" }
	| { readonly type: "insert.pageBreak" }
	| { readonly type: "insert.lineBreak" }
	// Word's horizontal line: an empty paragraph with a bottom border (AutoFormat's "---").
	| { readonly type: "insert.horizontalLine" }
	// Characters from a symbol picker (Unicode text, one or a few characters). U+2011 and U+00AD become Word's
	// nonbreaking and optional hyphens.
	| { readonly type: "insert.symbol"; readonly text: string }

export const INSERT_COMMANDS: readonly DocxCommandInfo<InsertCommand["type"]>[] = [
	{ type: "link.set", labelKey: "docx.link.set", icon: "link", group: "link", write: true, scope: "story", input: "link" },
	{ type: "link.remove", labelKey: "docx.link.remove", icon: "unlink", group: "link", write: true, scope: "story" },
	{ type: "insert.pageBreak", labelKey: "docx.insert.pageBreak", icon: "page-break", group: "insert", write: true, scope: "story" },
	{ type: "insert.lineBreak", labelKey: "docx.insert.lineBreak", icon: "line-break", group: "insert", write: true, scope: "story" },
	{
		type: "insert.horizontalLine",
		labelKey: "docx.insert.horizontalLine",
		icon: "horizontal-line",
		group: "insert",
		write: true,
		scope: "story"
	},
	{
		type: "insert.symbol",
		labelKey: "docx.insert.symbol",
		icon: "symbol",
		group: "insert",
		write: true,
		scope: "story",
		input: "text"
	}
]

// URLs a link may hold (as typed; what opening one allows is the app's link policy).
const MAX_TARGET = 2048
// Word's bookmark names.
const BOOKMARK_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,39}$/
export const MAX_LINK_TEXT = 1024
// A grapheme or a few (a flag, a family emoji): what a symbol picker inserts at once.
const MAX_SYMBOL = 16

// Text XML can hold: no controls but tab, no noncharacters or lone surrogates.
export function insertable(text: string): boolean {
	for (let k = 0; k < text.length; k++) {
		const code = text.charCodeAt(k)

		if (code >= 0xd800 && code <= 0xdbff && (text.charCodeAt(k + 1) & 0xfc00) === 0xdc00) k++
		else if (
			(code < 0x20 && code !== 0x09) ||
			code === 0x7f ||
			(code >= 0xd800 && code <= 0xdfff) ||
			code === 0xfffe ||
			code === 0xffff
		)
			return false
	}

	return true
}

export const parseInsertCommand: DocxCommandParser<InsertCommand> = (type, value) => {
	if (type === "insert.symbol") {
		const text = stringArg(value, "text", MAX_SYMBOL, 1)

		return text === undefined || !insertable(text) ? undefined : { type, text }
	}

	if (type !== "link.set") return { type } as InsertCommand

	const hasText = field(value, "text") !== undefined
	const text = stringArg(value, "text", MAX_LINK_TEXT, 1)

	if (hasText && (text === undefined || !insertable(text))) return undefined

	const extra = text === undefined ? {} : { text }
	const target = stringArg(value, "target", MAX_TARGET, 1)
	const anchor = field(value, "anchor")
	const heading = field(value, "heading")
	const given = [field(value, "target"), anchor, heading].filter(arg => arg !== undefined).length

	if (given !== 1) return undefined
	if (target !== undefined && !hasControl(target)) return { type, target, ...extra }
	if (typeof anchor === "string" && BOOKMARK_NAME.test(anchor)) return { type, anchor, ...extra }

	const pos = intArg(value, "heading", 0, Number.MAX_SAFE_INTEGER)

	return pos === undefined ? undefined : { type, heading: pos, ...extra }
}
