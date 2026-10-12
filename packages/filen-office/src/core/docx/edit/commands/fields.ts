import { boolArg, intArg, stringArg, type DocxCommandInfo, type DocxCommandParser } from "./base"

// Fields and tables of contents. Updates wait for layout (page numbers), so they run asynchronously.

export type FieldCommand =
	| { readonly type: "field.update" }
	| { readonly type: "field.updateAll" }
	| { readonly type: "toc.insert"; readonly levels: number; readonly hyperlinks: boolean }
	| { readonly type: "toc.update"; readonly pagesOnly: boolean }
	| { readonly type: "insert.pageNumber" }
	| { readonly type: "insert.pageCount" }
	// A Word date picture ("d MMMM yyyy").
	| { readonly type: "insert.date"; readonly format: string }
	| { readonly type: "bookmark.add"; readonly name: string }

const document = { write: true, scope: "document", group: "field" } as const

export const FIELD_COMMANDS: readonly DocxCommandInfo<FieldCommand["type"]>[] = [
	{ type: "field.update", labelKey: "docx.field.update", icon: "field-update", async: true, ...document },
	{ type: "field.updateAll", labelKey: "docx.field.updateAll", icon: "field-update-all", async: true, ...document },
	{ type: "toc.insert", labelKey: "docx.toc.insert", icon: "toc", async: true, ...document },
	{ type: "toc.update", labelKey: "docx.toc.update", icon: "toc-update", async: true, ...document },
	{ type: "insert.pageNumber", labelKey: "docx.insert.pageNumber", icon: "page-number", ...document },
	{ type: "insert.pageCount", labelKey: "docx.insert.pageCount", icon: "page-count", ...document },
	{ type: "insert.date", labelKey: "docx.insert.date", icon: "date", input: "date", ...document },
	{ type: "bookmark.add", labelKey: "docx.bookmark.add", icon: "bookmark", input: "text", ...document }
]

// Word's bookmark names: a letter, then letters, digits and underscores, 40 at most.
const BOOKMARK_NAME = /^[A-Za-z][A-Za-z0-9_]{0,39}$/
const DATE_PICTURE = /^[A-Za-z0-9 .,:/\\\-'"]{1,64}$/

export const parseFieldCommand: DocxCommandParser<FieldCommand> = (type, value) => {
	switch (type) {
		case "toc.insert": {
			const levels = intArg(value, "levels", 1, 9)
			const hyperlinks = boolArg(value, "hyperlinks")

			return levels === undefined || hyperlinks === undefined ? undefined : { type, levels, hyperlinks }
		}
		case "toc.update": {
			const pagesOnly = boolArg(value, "pagesOnly")

			return pagesOnly === undefined ? undefined : { type, pagesOnly }
		}
		case "insert.date": {
			const format = stringArg(value, "format", 64, 1)

			return format === undefined || !DATE_PICTURE.test(format) ? undefined : { type, format }
		}
		case "bookmark.add": {
			const name = stringArg(value, "name", 40, 1)

			return name === undefined || !BOOKMARK_NAME.test(name) ? undefined : { type, name }
		}
		default:
			return { type }
	}
}
