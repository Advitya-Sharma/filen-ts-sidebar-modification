import { describe, expect, it } from "vitest"
import { runCss } from "../../core/docx/reflow/css"
import { sessionOf } from "../docxEditFixtures"

// Font names from a document reach the reflowed view's style attributes inside a CSS string: nothing in a
// name may end that string and start declarations of its own.

// Property names of a declaration list as CSS reads it: a line break ends a string (a bad string), so the
// next ";" outside a string starts a new declaration.
function properties(css: string): string[] {
	const names: string[] = []
	let quote = ""
	let start = 0

	for (let at = 0; at <= css.length; at++) {
		const char = css[at] ?? ";"

		if (quote !== "") {
			if (char === "\\") at++
			else if (char === quote || char === "\n" || char === "\r" || char === "\f") quote = ""
		} else if (char === '"' || char === "'") {
			quote = char
		} else if (char === ";") {
			const name = css.slice(start, at).split(":")[0]?.trim() ?? ""

			if (name !== "") names.push(name)

			start = at + 1
		}
	}

	return names
}

const HOSTILE = "X&#10;;position:fixed;inset:0;z-index:2147483647;background:url(https://attacker.invalid/b);x:"

describe("reflow font CSS", () => {
	it.each(["ascii", "eastAsia", "cs", "hAnsi"])("keeps a hostile w:%s font name inside its string", async attribute => {
		const session = await sessionOf({
			body: `<w:p><w:r><w:rPr><w:rFonts w:${attribute}="${HOSTILE}"/><w:sz w:val="24"/></w:rPr><w:t>text</w:t></w:r></w:p>`
		})
		const paragraph = session.state.doc.child(0)
		const run = session.formats.resolvedRun(session.formats.paragraph(paragraph), paragraph.child(0).marks)
		const css = runCss(run)

		expect(
			properties(css).every(name => ["font-family", "font-size"].includes(name)),
			css
		).toBe(true)
		expect([...css].some(char => char.charCodeAt(0) < 0x20 || char.charCodeAt(0) === 0x7f)).toBe(false)
	})
})
