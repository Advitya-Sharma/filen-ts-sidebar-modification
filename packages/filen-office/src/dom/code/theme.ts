import { HighlightStyle, syntaxHighlighting } from "@codemirror/language"
import type { Extension } from "@codemirror/state"
import { EditorView } from "@codemirror/view"
import { tags as t, type Tag } from "@lezer/highlight"

// Syntax colours are CSS variables, so the editor follows a theme change (including the OS flipping the
// system appearance) with no reconfiguration. One rule per slot, each slot one tag; a token takes the most
// specific slot its tag falls under. Light is CodeMirror's default highlight style and dark One Dark (from
// @codemirror/language and @codemirror/theme-one-dark, MIT, see THIRD_PARTY_NOTICES.md), each
// resolved per slot tag; the slots are the tags of both styles plus the modified tags where the two styles'
// resolutions would otherwise differ, so both come out exact (codeTheme.dom.test.ts). The light value is
// the variable's fallback, and a dark theme sets every variable (codeSyntaxVariables).

type Pair = readonly [light: string | null, dark: string | null]

interface Slot {
	readonly slot: string
	readonly tag: Tag
	readonly color?: Pair
	readonly fontWeight?: Pair
	readonly fontStyle?: Pair
	readonly textDecoration?: Pair
}

const SLOTS: readonly Slot[] = [
	{ slot: "meta", tag: t.meta, color: ["#404740", "#7d8799"] },
	{ slot: "link", tag: t.link, color: [null, "#7d8799"], textDecoration: ["underline", "underline"] },
	{ slot: "heading", tag: t.heading, color: [null, "#e06c75"], fontWeight: ["bold", "bold"], textDecoration: ["underline", null] },
	{ slot: "emphasis", tag: t.emphasis, fontStyle: ["italic", "italic"] },
	{ slot: "strong", tag: t.strong, fontWeight: ["bold", "bold"] },
	{ slot: "strikethrough", tag: t.strikethrough, textDecoration: ["line-through", "line-through"] },
	{ slot: "keyword", tag: t.keyword, color: ["#708", "#c678dd"], fontWeight: [null, null] },
	{ slot: "atom", tag: t.atom, color: ["#219", "#d19a66"] },
	{ slot: "bool", tag: t.bool, color: ["#219", "#d19a66"] },
	{ slot: "url", tag: t.url, color: ["#219", "#56b6c2"] },
	{ slot: "content-separator", tag: t.contentSeparator, color: ["#219", null] },
	{ slot: "label-name", tag: t.labelName, color: ["#219", "#61afef"] },
	{ slot: "literal", tag: t.literal, color: ["#164", null] },
	{ slot: "inserted", tag: t.inserted, color: ["#164", "#98c379"] },
	{ slot: "string", tag: t.string, color: ["#a11", "#98c379"] },
	{ slot: "deleted", tag: t.deleted, color: ["#a11", "#e06c75"] },
	{ slot: "regexp", tag: t.regexp, color: ["#e40", "#56b6c2"] },
	{ slot: "escape", tag: t.escape, color: ["#e40", "#56b6c2"] },
	{ slot: "special-string", tag: t.special(t.string), color: ["#e40", "#56b6c2"] },
	{ slot: "definition-variable-name", tag: t.definition(t.variableName), color: ["#00f", "#abb2bf"] },
	{ slot: "local-variable-name", tag: t.local(t.variableName), color: ["#30a", "#e06c75"] },
	{ slot: "type-name", tag: t.typeName, color: ["#085", "#e5c07b"] },
	{ slot: "namespace", tag: t.namespace, color: ["#085", "#e5c07b"] },
	{ slot: "class-name", tag: t.className, color: ["#167", "#e5c07b"] },
	{ slot: "special-variable-name", tag: t.special(t.variableName), color: ["#256", "#d19a66"] },
	{ slot: "macro-name", tag: t.macroName, color: ["#256", "#e06c75"] },
	{ slot: "definition-property-name", tag: t.definition(t.propertyName), color: ["#00c", "#e06c75"] },
	{ slot: "comment", tag: t.comment, color: ["#940", "#7d8799"], fontStyle: [null, null] },
	{ slot: "invalid", tag: t.invalid, color: ["#f00", "#ffffff"] },
	{ slot: "name", tag: t.name, color: [null, "#e06c75"] },
	{ slot: "character", tag: t.character, color: ["#a11", "#e06c75"] },
	{ slot: "property-name", tag: t.propertyName, color: [null, "#e06c75"] },
	{ slot: "function-variable-name", tag: t.function(t.variableName), color: [null, "#61afef"] },
	{ slot: "color", tag: t.color, color: ["#164", "#d19a66"] },
	{ slot: "constant-name", tag: t.constant(t.name), color: [null, "#d19a66"] },
	{ slot: "standard-name", tag: t.standard(t.name), color: [null, "#d19a66"] },
	{ slot: "definition-name", tag: t.definition(t.name), color: [null, "#abb2bf"] },
	{ slot: "separator", tag: t.separator, color: [null, "#abb2bf"] },
	{ slot: "number", tag: t.number, color: ["#164", "#e5c07b"] },
	{ slot: "changed", tag: t.changed, color: [null, "#e5c07b"] },
	{ slot: "annotation", tag: t.annotation, color: ["#404740", "#e5c07b"] },
	{ slot: "modifier", tag: t.modifier, color: ["#708", "#e5c07b"] },
	{ slot: "self", tag: t.self, color: ["#708", "#e5c07b"] },
	{ slot: "operator", tag: t.operator, color: [null, "#56b6c2"] },
	{ slot: "operator-keyword", tag: t.operatorKeyword, color: ["#708", "#56b6c2"] },
	{ slot: "processing-instruction", tag: t.processingInstruction, color: ["#404740", "#98c379"] },
	{ slot: "special-character", tag: t.special(t.character), color: ["#e40", "#e06c75"] },
	{ slot: "definition-function-variable-name", tag: t.definition(t.function(t.variableName)), color: ["#00f", "#61afef"] },
	{ slot: "definition-special-variable-name", tag: t.definition(t.special(t.variableName)), color: ["#00f", "#d19a66"] },
	{ slot: "constant-local-variable-name", tag: t.constant(t.local(t.variableName)), color: ["#30a", "#d19a66"] },
	{ slot: "function-local-variable-name", tag: t.function(t.local(t.variableName)), color: ["#30a", "#61afef"] },
	{ slot: "function-special-variable-name", tag: t.function(t.special(t.variableName)), color: ["#256", "#61afef"] },
	{ slot: "standard-local-variable-name", tag: t.standard(t.local(t.variableName)), color: ["#30a", "#d19a66"] },
	{ slot: "local-special-variable-name", tag: t.local(t.special(t.variableName)), color: ["#30a", "#d19a66"] }
]

type Property = "color" | "fontWeight" | "fontStyle" | "textDecoration"

const PROPERTIES: readonly Property[] = ["color", "fontWeight", "fontStyle", "textDecoration"]

const SUFFIX: Readonly<Record<Property, string>> = { color: "", fontWeight: "-weight", fontStyle: "-style", textDecoration: "-decoration" }

// What a theme that leaves a property alone sets, so the light fallback does not show through.
const NEUTRAL: Readonly<Record<Property, string>> = {
	color: "currentColor",
	fontWeight: "normal",
	fontStyle: "normal",
	textDecoration: "none"
}

function variable(slot: Slot, property: Property): string {
	return `--fo-code-${slot.slot}${SUFFIX[property]}`
}

export const codeHighlightStyle: HighlightStyle = HighlightStyle.define(
	SLOTS.map(slot => {
		const spec: { tag: Tag; [property: string]: string | Tag } = { tag: slot.tag }

		for (const property of PROPERTIES) {
			const pair = slot[property]

			if (pair !== undefined) {
				spec[property] = pair[0] === null ? `var(${variable(slot, property)})` : `var(${variable(slot, property)}, ${pair[0]})`
			}
		}

		return spec
	})
)

// The syntax variables a theme sets: "light" is the fallbacks (an app need not set them), "dark" One Dark.
export function codeSyntaxVariables(theme: "light" | "dark"): Readonly<Record<string, string>> {
	const index = theme === "light" ? 0 : 1
	const variables: Record<string, string> = {}

	for (const slot of SLOTS) {
		for (const property of PROPERTIES) {
			const pair = slot[property]

			if (pair === undefined) {
				continue
			}

			const value = pair[index]

			if (value !== null) {
				variables[variable(slot, property)] = value
			} else if (index === 1 && pair[0] !== null) {
				variables[variable(slot, property)] = NEUTRAL[property]
			}
		}
	}

	return variables
}

// Editor chrome, every colour a --fo-code-* variable with a neutral default.
const CHROME = EditorView.theme({
	"&": { backgroundColor: "var(--fo-code-bg, transparent)", color: "var(--fo-code-fg, currentColor)" },
	".cm-gutters": { backgroundColor: "var(--fo-code-gutter-bg, transparent)", color: "var(--fo-code-gutter-fg, #6e7781)", border: "none" },
	".cm-activeLine, .cm-activeLineGutter": { backgroundColor: "var(--fo-code-active-line, rgba(0, 0, 0, 0.04))" },
	".cm-cursor, .cm-dropCursor": { borderLeftColor: "var(--fo-code-caret, currentColor)" },
	"&.cm-focused .cm-selectionBackground, .cm-selectionBackground, .cm-content ::selection": {
		backgroundColor: "var(--fo-code-selection, color-mix(in oklab, currentColor 18%, transparent))"
	},
	".cm-panels": { backgroundColor: "var(--fo-code-panel-bg, #fff)", color: "var(--fo-code-panel-fg, currentColor)" },
	".cm-panels.cm-panels-top": { borderBottom: "1px solid var(--fo-code-border, #d0d7de)" },
	".cm-panels.cm-panels-bottom": { borderTop: "1px solid var(--fo-code-border, #d0d7de)" },
	".cm-search": { display: "flex", flexWrap: "wrap", alignItems: "center", gap: "6px", padding: "8px 36px 8px 12px", fontSize: "13px" },
	".cm-search br": { flexBasis: "100%", height: "0" },
	".cm-textfield": {
		height: "28px",
		margin: "0",
		padding: "0 8px",
		fontSize: "13px",
		color: "var(--fo-code-panel-fg, currentColor)",
		backgroundColor: "transparent",
		border: "1px solid var(--fo-code-input-border, #d0d7de)",
		borderRadius: "var(--fo-code-radius, 8px)",
		outline: "none"
	},
	".cm-textfield:focus": {
		borderColor: "var(--fo-code-focus-ring, #0969da)",
		boxShadow: "0 0 0 2px var(--fo-code-focus-shadow, color-mix(in oklab, #0969da 35%, transparent))"
	},
	".cm-button": {
		height: "28px",
		margin: "0",
		padding: "0 10px",
		fontSize: "13px",
		color: "var(--fo-code-button-fg, currentColor)",
		backgroundImage: "none",
		backgroundColor: "var(--fo-code-button-bg, #f0f1f3)",
		border: "none",
		borderRadius: "var(--fo-code-radius, 8px)",
		cursor: "pointer"
	},
	".cm-button:hover": { backgroundColor: "var(--fo-code-button-hover-bg, #e4e6e9)" },
	".cm-search label": {
		display: "inline-flex",
		alignItems: "center",
		gap: "4px",
		color: "var(--fo-code-muted-fg, #6e7781)",
		fontSize: "13px"
	},
	".cm-search button[name=close]": {
		top: "50%",
		right: "10px",
		transform: "translateY(-50%)",
		fontSize: "18px",
		color: "var(--fo-code-muted-fg, #6e7781)",
		cursor: "pointer"
	},
	".cm-searchMatch": { backgroundColor: "var(--fo-code-match, color-mix(in oklab, #facc15 35%, transparent))", outline: "none" },
	".cm-searchMatch.cm-searchMatch-selected": {
		backgroundColor: "var(--fo-code-match-selected, color-mix(in oklab, #f97316 45%, transparent))"
	}
})

// Chrome plus syntax colours: one value for every editor in both themes.
export const codeTheme: Extension = [CHROME, syntaxHighlighting(codeHighlightStyle)]
