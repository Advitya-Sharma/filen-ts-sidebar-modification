import { boolArg, stringArg, type DocxCommandInfo, type DocxCommandParser } from "./base"

// Equations, entered as a LaTeX subset.

export const MAX_LATEX = 10_000

export type MathCommand =
	| { readonly type: "equation.insert"; readonly latex: string; readonly display: boolean }
	| { readonly type: "equation.edit"; readonly latex: string }

export const MATH_COMMANDS: readonly DocxCommandInfo<MathCommand["type"]>[] = [
	{
		type: "equation.insert",
		labelKey: "docx.equation.insert",
		icon: "equation",
		write: true,
		scope: "story",
		group: "equation",
		input: "equation"
	},
	{
		type: "equation.edit",
		labelKey: "docx.equation.edit",
		icon: "equation-edit",
		write: true,
		scope: "story",
		group: "equation",
		input: "equation"
	}
]

export const parseMathCommand: DocxCommandParser<MathCommand> = (type, value) => {
	const latex = stringArg(value, "latex", MAX_LATEX, 1)

	if (latex === undefined) return undefined
	if (type === "equation.edit") return { type, latex }

	const display = boolArg(value, "display")

	return display === undefined ? undefined : { type, latex, display }
}
