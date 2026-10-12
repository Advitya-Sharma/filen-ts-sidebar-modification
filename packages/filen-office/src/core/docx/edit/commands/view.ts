import { boolArg, type DocxCommandInfo, type DocxCommandParser } from "./base"

// How the document is shown. The surface is the app's state; the command is listed so every app labels it
// alike.

export type ViewCommand = { readonly type: "view.reflow"; readonly on: boolean }

export const VIEW_COMMANDS: readonly DocxCommandInfo<ViewCommand["type"]>[] = [
	{ type: "view.reflow", labelKey: "docx.view.reflow", icon: "reflow", write: false, scope: "document", group: "view" }
]

export const parseViewCommand: DocxCommandParser<ViewCommand> = (type, value) => {
	const on = boolArg(value, "on")

	return on === undefined ? undefined : { type, on }
}
