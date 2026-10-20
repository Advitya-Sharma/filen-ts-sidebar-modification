// @vitest-environment happy-dom
import { describe, expect, it } from "vitest"
import { defaultHighlightStyle, HighlightStyle } from "@codemirror/language"
import { Tag, tags as t } from "@lezer/highlight"
import { codeHighlightStyle, codeSyntaxVariables } from "../dom/code/index"

// One Dark's highlight style as @codemirror/theme-one-dark 6.1.3 defines it.
const chalky = "#e5c07b"
const coral = "#e06c75"
const cyan = "#56b6c2"
const ivory = "#abb2bf"
const stone = "#7d8799"
const malibu = "#61afef"
const sage = "#98c379"
const whiskey = "#d19a66"
const violet = "#c678dd"
const ONE_DARK = HighlightStyle.define([
	{ tag: t.keyword, color: violet },
	{ tag: [t.name, t.deleted, t.character, t.propertyName, t.macroName], color: coral },
	{ tag: [t.function(t.variableName), t.labelName], color: malibu },
	{ tag: [t.color, t.constant(t.name), t.standard(t.name)], color: whiskey },
	{ tag: [t.definition(t.name), t.separator], color: ivory },
	{ tag: [t.typeName, t.className, t.number, t.changed, t.annotation, t.modifier, t.self, t.namespace], color: chalky },
	{ tag: [t.operator, t.operatorKeyword, t.url, t.escape, t.regexp, t.link, t.special(t.string)], color: cyan },
	{ tag: [t.meta, t.comment], color: stone },
	{ tag: t.strong, fontWeight: "bold" },
	{ tag: t.emphasis, fontStyle: "italic" },
	{ tag: t.strikethrough, textDecoration: "line-through" },
	{ tag: t.link, color: stone, textDecoration: "underline" },
	{ tag: t.heading, fontWeight: "bold", color: coral },
	{ tag: [t.atom, t.bool, t.special(t.variableName)], color: whiskey },
	{ tag: [t.processingInstruction, t.string, t.inserted], color: sage },
	{ tag: t.invalid, color: "#ffffff" }
])

type Declarations = Record<string, string>

function rules(style: HighlightStyle): Map<string, Declarations> {
	const byClass = new Map<string, Declarations>()

	for (const rule of style.module?.getRules().split("\n") ?? []) {
		const match = /^\.(\S+)\s*\{(.*)\}$/.exec(rule.trim())

		if (match?.[1] !== undefined && match[2] !== undefined) {
			const declarations: Declarations = {}

			for (const declaration of match[2].split(";")) {
				const colon = declaration.indexOf(":")

				if (colon > 0) {
					declarations[declaration.slice(0, colon).trim()] = declaration.slice(colon + 1).trim()
				}
			}

			byClass.set(match[1], declarations)
		}
	}

	return byClass
}

function resolved(style: HighlightStyle, tag: Tag): Declarations {
	const name = style.style([tag])

	return name === null ? {} : (rules(style).get(name) ?? {})
}

// A theme's own value for a property: the variable when set, else its fallback; neutral values mean unset.
const NEUTRAL: Readonly<Record<string, string>> = {
	color: "currentColor",
	"font-weight": "normal",
	"font-style": "normal",
	"text-decoration": "none"
}

function substituted(declarations: Declarations, variables: Readonly<Record<string, string>>): Declarations {
	const out: Declarations = {}

	for (const [property, value] of Object.entries(declarations)) {
		const match = /^var\((--[\w-]+)(?:,\s*(.+))?\)$/.exec(value)
		const actual = match?.[1] === undefined ? value : (variables[match[1]] ?? match[2])

		if (actual !== undefined && actual !== NEUTRAL[property]) {
			out[property] = actual
		}
	}

	return out
}

// Every tag @lezer/highlight defines, alone and under one or two modifiers.
function everyTag(): Tag[] {
	const base = Object.values(t).filter((value): value is Tag => value instanceof Tag)
	const modifiers = [t.definition, t.constant, t.function, t.standard, t.local, t.special]
	const once = modifiers.flatMap(modifier => base.map(tag => modifier(tag)))

	return [...base, ...once, ...modifiers.flatMap(modifier => once.map(tag => modifier(tag)))]
}

describe("code highlight style", () => {
	it("styles every tag in light exactly as CodeMirror's default highlight style", () => {
		for (const tag of everyTag()) {
			expect(substituted(resolved(codeHighlightStyle, tag), {}), String(tag)).toEqual(resolved(defaultHighlightStyle, tag))
		}
	})

	it("styles every tag in dark exactly as One Dark", () => {
		const dark = codeSyntaxVariables("dark")

		for (const tag of everyTag()) {
			expect(substituted(resolved(codeHighlightStyle, tag), dark), String(tag)).toEqual(resolved(ONE_DARK, tag))
		}
	})

	it("leaves the light theme to the fallbacks and sets every fallback in dark", () => {
		const light = codeSyntaxVariables("light")
		const dark = codeSyntaxVariables("dark")

		for (const name of Object.keys(light)) {
			expect(dark[name], name).toBeDefined()
		}
	})

	it("colours through variables only, so a theme change needs no reconfiguration", () => {
		for (const declarations of rules(codeHighlightStyle).values()) {
			for (const value of Object.values(declarations)) {
				expect(value.startsWith("var(--fo-code-")).toBe(true)
			}
		}
	})
})
