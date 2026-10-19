import { describe, expect, it } from "vitest"
import { macroButtonText, shownFieldText } from "../core/docx/layout/fields"
import { docxWordCount, type DocxWordCount } from "../core/docx/paged/queries"
import { Item } from "../core/docx/model/kinds"
import { buildDocx } from "./docxSynthetic"
import { loadDocx } from "./layoutFixtures"
import { pagedEditor } from "./pagedFixtures"

const instr = (text: string): string => `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`
const char = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`
const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const para = (inner: string): string => `<w:p>${inner}</w:p>`

// A MACROBUTTON with a result, without one, with an empty one, w:fldSimple with and without one, one inside
// another field's result, and one whose display text is quoted.
const BODY = [
	para(char("begin") + instr(' MACROBUTTON  synchronize "Sync contents" ') + char("separate") + run("*") + char("end")),
	para(char("begin") + instr(" MACROBUTTON  synchronize Without result ") + char("end")),
	para(char("begin") + instr(" MACROBUTTON NoMacro [Click here] ") + char("separate") + char("end")),
	para('<w:fldSimple w:instr=" MACROBUTTON NoMacro Simple text "/>'),
	para(`<w:fldSimple w:instr=" MACROBUTTON NoMacro Simple text ">${run("stored")}</w:fldSimple>`),
	para(
		char("begin") +
			instr(" QUOTE x ") +
			char("separate") +
			run("outer ") +
			char("begin") +
			instr(" MACROBUTTON NoMacro Nested ") +
			char("end") +
			char("end")
	),
	para(run("Body"))
].join("")

const SHOWN = ["*", "Without result", "[Click here]", "Simple text", "stored", "outer ", "Body"]

describe("MACROBUTTON fields", () => {
	it("read their display text from the instruction", () => {
		expect(macroButtonText(' MACROBUTTON  synchronize "Sync contents" ')).toBe("Sync contents")
		expect(macroButtonText(" MACROBUTTON  synchronize Without result ")).toBe("Without result")
		expect(macroButtonText('macrobutton "My macro" "say \\"hi\\"" \\* MERGEFORMAT')).toBe('say "hi"')
		expect(macroButtonText(" MACROBUTTON NoMacro ")).toBe("")
		expect(macroButtonText(" MACROBUTTONS x y")).toBeUndefined()
		expect(macroButtonText(" PAGE ")).toBeUndefined()
		expect(shownFieldText("*", " MACROBUTTON a b ")).toBe("*")
		expect(shownFieldText("", " MACROBUTTON a b ")).toBe("b")
		expect(shownFieldText("", " PAGE ")).toBe("")
	})

	it("count and read as text what they show", async () => {
		const bytes = buildDocx({ body: BODY })
		const { model } = await loadDocx(bytes)
		const part = model.parts[model.stories[0]?.part ?? -1]

		if (part === undefined) throw new Error("no body part")

		const texts: string[] = []

		for (let item = 0; item < part.kinds.length; item++) if (part.kinds[item] === Item.Paragraph) texts.push(part.paragraphText(item))

		expect(texts).toEqual(SHOWN)

		// A count as of the same text in plain runs.
		const counted = async (source: Uint8Array): Promise<DocxWordCount> => {
			const { editor } = await pagedEditor(source)
			const count = await docxWordCount(editor)

			editor.dispose()

			return count
		}

		expect(await counted(bytes)).toEqual(await counted(buildDocx({ body: SHOWN.map(text => para(run(text))).join("") })))
	})
})
