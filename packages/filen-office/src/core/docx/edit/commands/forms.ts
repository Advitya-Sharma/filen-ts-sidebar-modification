import { choiceArg, field, hasControl, stringArg, type DocxCommandInfo, type DocxCommandParser } from "./base"

// Content controls (ECMA-376 Part 1 §17.5.2; checkboxes and repeating sections MS-DOCX §2.6) and legacy form
// fields (§17.16.17 w:ffData) at the selection.

// The controls control.insert makes.
export type DocxControlKind = "plainText" | "richText" | "checkbox" | "dropDownList" | "comboBox" | "date" | "picture" | "repeatingSection"

export const DOCX_CONTROL_KINDS: readonly DocxControlKind[] = [
	"plainText",
	"richText",
	"checkbox",
	"dropDownList",
	"comboBox",
	"date",
	"picture",
	"repeatingSection"
]

export interface DocxControlItemInput {
	readonly text: string
	readonly value: string
}

export type FormCommand =
	// Toggles a checkbox (a content control or a legacy check box field).
	| { readonly type: "control.check" }
	// A list item's value (w:listItem w:value; a legacy drop-down's entry text).
	| { readonly type: "control.choose"; readonly value: string }
	// ISO 8601 date ("2026-10-12").
	| { readonly type: "control.date"; readonly iso: string }
	// Empties a control: its placeholder shows again (a list or date control loses its choice).
	| { readonly type: "control.clear" }
	// The text of a legacy text form field (w:textInput).
	| { readonly type: "control.text"; readonly text: string }
	// A new control at the selection (around the selected paragraphs for a repeating section); list entries
	// for a drop-down list or combo box.
	| { readonly type: "control.insert"; readonly kind: DocxControlKind; readonly items?: readonly DocxControlItemInput[] }
	// Adds an item after the repeating section item at the selection, or removes that item.
	| { readonly type: "control.item.add" }
	| { readonly type: "control.item.remove" }

const story = { write: true, scope: "story", group: "control", context: "control" } as const

export const FORM_COMMANDS: readonly DocxCommandInfo<FormCommand["type"]>[] = [
	{ type: "control.check", labelKey: "docx.control.check", icon: "checkbox", ...story },
	{ type: "control.choose", labelKey: "docx.control.choose", icon: "dropdown", input: "choice", ...story },
	{ type: "control.date", labelKey: "docx.control.date", icon: "date", input: "date", ...story },
	{ type: "control.clear", labelKey: "docx.control.clear", icon: "control-clear", ...story },
	{ type: "control.text", labelKey: "docx.control.text", icon: "control-text", input: "text", ...story },
	{ type: "control.insert", labelKey: "docx.control.insert", icon: "control-insert", input: "choice", ...story },
	{ type: "control.item.add", labelKey: "docx.control.item.add", icon: "control-item-add", ...story },
	{ type: "control.item.remove", labelKey: "docx.control.item.remove", icon: "control-item-remove", ...story }
]

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/
const KINDS: ReadonlySet<string> = new Set(DOCX_CONTROL_KINDS)
// Word caps list entries at 255 characters; a list holds at most this many entries here.
const MAX_ITEM = 255
const MAX_ITEMS = 1000
// A legacy text field's w:maxLength is an ST_DecimalNumber; Word stops at 255 by default.
const MAX_TEXT = 32767

function items(value: object): readonly DocxControlItemInput[] | undefined | null {
	const list = field(value, "items")

	if (list === undefined) return undefined
	if (!Array.isArray(list) || list.length > MAX_ITEMS) return null

	const out: DocxControlItemInput[] = []
	const seen = new Set<string>()

	for (const entry of list as unknown[]) {
		if (typeof entry !== "object" || entry === null || Array.isArray(entry)) return null

		const text = stringArg(entry, "text", MAX_ITEM, 1)
		const itemValue = stringArg(entry, "value", MAX_ITEM, 1)

		// Word refuses two entries with the same value or display text.
		if (text === undefined || itemValue === undefined || hasControl(text) || hasControl(itemValue)) return null
		if (seen.has(`v${itemValue}`) || seen.has(`t${text}`)) return null

		seen.add(`v${itemValue}`)
		seen.add(`t${text}`)
		out.push({ text, value: itemValue })
	}

	return out
}

export const parseFormCommand: DocxCommandParser<FormCommand> = (type, value) => {
	switch (type) {
		case "control.choose": {
			const chosen = stringArg(value, "value", MAX_ITEM)

			return chosen === undefined ? undefined : { type, value: chosen }
		}
		case "control.date": {
			const iso = stringArg(value, "iso", 10, 10)

			if (iso === undefined || !ISO_DATE.test(iso)) return undefined

			const date = new Date(`${iso}T00:00:00Z`)

			return !Number.isNaN(date.getTime()) && date.toISOString().startsWith(iso) ? { type, iso } : undefined
		}
		case "control.text": {
			const text = stringArg(value, "text", MAX_TEXT)

			// A text field is one run of text: line breaks become w:br, other control characters are refused.
			return text === undefined || hasControl(text.replace(/\n/g, "")) ? undefined : { type, text }
		}
		case "control.insert": {
			const kind = choiceArg<DocxControlKind>(value, "kind", KINDS)
			const list = items(value)

			if (kind === undefined || list === null) return undefined
			if (list !== undefined && kind !== "dropDownList" && kind !== "comboBox") return undefined

			return list === undefined ? { type, kind } : { type, kind, items: list }
		}
		default:
			return { type }
	}
}
