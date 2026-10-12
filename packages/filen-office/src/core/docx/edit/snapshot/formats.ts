import type { DocxHighlightColor, DocxLineSpacingRule } from "../commands/formatting"
import type { SnapshotInput } from "./index"

// Formatting values of the selection (font, size, colours, spacing, indents): one value, "mixed" when the
// selection holds several, null where there is nothing to tell. Counted additively over the selection like
// the toggles, so a growing selection costs its difference.

export interface DocxValuesSnapshot {
	readonly font: string | "mixed" | null
	// Half-points.
	readonly size: number | "mixed" | null
	// "auto", "RRGGBB" or "theme:<name>".
	readonly color: string | "mixed" | null
	readonly highlight: DocxHighlightColor | "mixed" | null
	readonly vertAlign: "baseline" | "superscript" | "subscript" | "mixed" | null
	readonly lineSpacing: { readonly rule: DocxLineSpacingRule; readonly value: number } | "mixed" | null
	// Twips.
	readonly spacing: { readonly before: number; readonly after: number } | "mixed" | null
	readonly indent: { readonly start: number; readonly end: number; readonly firstLine: number; readonly hanging: number } | "mixed" | null
}

export const NO_VALUES: DocxValuesSnapshot = Object.freeze({
	font: null,
	size: null,
	color: null,
	highlight: null,
	vertAlign: null,
	lineSpacing: null,
	spacing: null,
	indent: null
})

// The link at the selection.
export interface DocxLinkContext {
	readonly target: string | null
	readonly anchor: string | null
	readonly text: string
}

export function valuesOf(_input: SnapshotInput): DocxValuesSnapshot {
	return NO_VALUES
}

export function linkOf(_input: SnapshotInput): DocxLinkContext | null {
	return null
}
