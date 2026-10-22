import { describe, expect, it } from "vitest"
import { MAX_HTML_COPY_LENGTH, parseClipboardHtml } from "../core/sheet/clipboard/html"
import { SHEET_MAX_EDIT_CELLS } from "../core/sheet/protocol"
import { cpuMs } from "./timing"

// Clipboard HTML any page can put on the clipboard, at the paste cap (768 KiB), in shapes that cost the square of
// their size to a parse with backtracking expressions or per-cell copies of shared rules. Each parses in about as long
// as a plain table of that size (about 90 ms); a quadratic step takes minutes. The bound leaves room for a loaded machine.

const SIZE = MAX_HTML_COPY_LENGTH
const HALF = SIZE / 2
const BOUND_MS = 2_500
// Shapes whose cost is per character of what a cell holds, at the plain table's own rate.
const TIGHT: Readonly<Record<string, number>> = { libreOfficeCodes: 150, libreOfficeHugeCode: 150 }

const declarations = (count: number): string => Array.from({ length: count }, (_, k) => `x${String(k)}:1`).join(";")
const repeat = (unit: string, length: number): string => unit.repeat(Math.floor(length / unit.length))

function tds(cell: string, room: number): string {
	return `<table><tr>${repeat(cell, room)}</tr></table>`
}

const SHAPES: Record<string, () => string> = {
	plain: () => tds("<td>12345</td>", SIZE - 40),
	newlines: () => `<table><tr><td>a${"\n".repeat(SIZE - 40)}a</td></tr></table>`,
	breaks: () => `<table><tr><td>a${repeat("<br>", SIZE - 40)}a</td></tr></table>`,
	spaces: () => `<table><tr><td>a${repeat(" \t", SIZE - 40)}a</td></tr></table>`,
	border: () => `<table><tr><td style="border:${repeat("1 ", SIZE - 60)})">1</td></tr></table>`,
	background: () => `<table><tr><td style="background:${repeat("1 ", SIZE - 60)})">1</td></tr></table>`,
	selectors: () => `<style>${repeat("td,", HALF)}td{${repeat("color:red;", HALF - 80)}}</style><table><tr><td>1</td></tr></table>`,
	classSelectors: () => {
		const selectors = Array.from({ length: SIZE / 32 }, (_, k) => `.c${String(k)}`).join(",")

		return `<style>${selectors}{${repeat("color:red;", SIZE / 4)}}</style>${tds('<td class="c1">1</td>', SIZE / 2)}`
	},
	comments: () => `<style>${repeat("/*", SIZE - 80)}</style><table><tr><td>1</td></tr></table>`,
	important: () => `<table><tr><td style="color:${repeat("!important ", SIZE - 60)}x">1</td></tr></table>`,
	tdRule: () => {
		const rule = declarations(HALF / 8)

		return `<style>td{${rule}}</style>${tds("<td></td>", SIZE - rule.length - 60)}`
	},
	sharedValue: () => {
		const value = repeat("1 ", SIZE / 16)

		return `<style>td{border:${value};color:${value};font-family:${value};mso-number-format:"${value}";background:${value};text-align:${value}}</style>${tds("<td>1</td>", SIZE / 2)}`
	},
	classes: () => {
		const rule = declarations(HALF / 8)

		return `<style>.a{${rule};color:red}</style>${tds(`<td class="${"a ".repeat(100)}">1</td>`, SIZE - rule.length - 60)}`
	},
	distinctClasses: () => {
		const names = Array.from({ length: SIZE / 26 }, (_, k) => `k${String(k)}`)

		return `<style>${names.map(name => `.${name}{color:red}`).join("")}</style><table><tr><td class="${names.join(" ")}">1</td></tr></table>`
	},
	nestedSpans: () => {
		const open = `<span style="${declarations(4)};color:red">`

		return `<table><tr><td>${repeat(open, HALF)}a${repeat("</span>", HALF - 60)}</td></tr></table>`
	},
	// LibreOffice format codes in a locale's syntax: many long ones, and one huge one.
	libreOfficeCodes: () => tds(`<td sdval="1,5" sdnum="1031;1031;${"0,0".repeat(300)}">1</td>`, SIZE - 40),
	libreOfficeHugeCode: () => `<table><tr><td sdval="1" sdnum="1031;1031;${repeat("TT.MM.JJJJ ", SIZE - 80)}">1</td></tr></table>`,
	nestedBold: () => `<table><tr><td>${repeat("<b>", HALF)}a${repeat("</b>", HALF - 60)}</td></tr></table>`
}

describe("clipboard HTML at the paste cap", () => {
	for (const [name, build] of Object.entries(SHAPES))
		it(`parses ${name} in linear time`, { timeout: 60_000 }, () => {
			const html = build()

			expect(html.length).toBeLessThanOrEqual(SIZE)
			expect(html.length).toBeGreaterThan(SIZE * 0.75)

			const start = cpuMs()
			const table = parseClipboardHtml(html, SHEET_MAX_EDIT_CELLS)

			expect(cpuMs() - start).toBeLessThan(TIGHT[name] ?? BOUND_MS)
			expect(table).not.toBeNull()
		})
})
