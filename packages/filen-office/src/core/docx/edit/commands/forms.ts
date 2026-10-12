import { stringArg, type DocxCommandInfo, type DocxCommandParser } from "./base"

// Content controls (checkbox, list, date, text) at the selection.

export type FormCommand =
	| { readonly type: "control.check" }
	// A list item's value (w:listItem w:value).
	| { readonly type: "control.choose"; readonly value: string }
	// ISO 8601 date ("2026-10-12").
	| { readonly type: "control.date"; readonly iso: string }
	| { readonly type: "control.clear" }

const story = { write: true, scope: "story", group: "control", context: "control" } as const

export const FORM_COMMANDS: readonly DocxCommandInfo<FormCommand["type"]>[] = [
	{ type: "control.check", labelKey: "docx.control.check", icon: "checkbox", ...story },
	{ type: "control.choose", labelKey: "docx.control.choose", icon: "dropdown", input: "choice", ...story },
	{ type: "control.date", labelKey: "docx.control.date", icon: "date", input: "date", ...story },
	{ type: "control.clear", labelKey: "docx.control.clear", icon: "control-clear", ...story }
]

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/

export const parseFormCommand: DocxCommandParser<FormCommand> = (type, value) => {
	switch (type) {
		case "control.choose": {
			// w:value is ST_String; Word caps list entries at 255 characters.
			const chosen = stringArg(value, "value", 255)

			return chosen === undefined ? undefined : { type, value: chosen }
		}
		case "control.date": {
			const iso = stringArg(value, "iso", 10, 10)

			if (iso === undefined || !ISO_DATE.test(iso)) return undefined

			const date = new Date(`${iso}T00:00:00Z`)

			return !Number.isNaN(date.getTime()) && date.toISOString().startsWith(iso) ? { type, iso } : undefined
		}
		default:
			return { type }
	}
}
