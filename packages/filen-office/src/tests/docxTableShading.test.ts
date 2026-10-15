import { unzipSync } from "fflate"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { docxSchema, parseDocxEditorCommand, type DocxEditorCommand } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { saveDocx } from "@filen/office/core/docx/write"
import { buildDocx, p, type SyntheticDocx } from "./docxSynthetic"
import { pageDigest, pagedEditor, savedDigest, settleAll } from "./pagedFixtures"

// A table cell filled with a theme colour is written as Word writes one (ECMA-376 Part 1 §17.3.5 shd): the colour
// it gives in w:fill and the reference in w:themeFill with w:themeFillTint or w:themeFillShade, which a later edit
// of the cell keeps and the picker reads back as that theme colour.

const nodes = docxSchema.nodes
const THEME =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>' +
	'<a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
	'<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
	'<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>' +
	'<a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
	'<a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
	'<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
	'<a:fmtScheme name="T"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>'
const cell = (text: string, tcPr = ""): string => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr>${p(text)}</w:tc>`
const table = (first: string): string =>
	'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
	`<w:tr>${first}${cell("b1")}</w:tr></w:tbl>`

function spec(first: string, theme: boolean): SyntheticDocx {
	return {
		body: p("Before") + table(first) + p("After"),
		...(theme
			? {
					parts: [
						{
							name: "word/theme/theme1.xml",
							contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
							data: THEME
						}
					],
					relationships: [{ id: "rIdTheme", type: "theme", target: "theme/theme1.xml" }]
				}
			: {})
	}
}

function caretIn(editor: PagedEditor, text: string): void {
	let at = -1

	editor.state.doc.descendants((node, pos) => {
		if (at < 0 && node.type === nodes.paragraph && node.textContent === text) at = pos + 1

		return at < 0
	})
	editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)))
}

async function run(editor: PagedEditor, command: DocxEditorCommand): Promise<boolean> {
	return await editor.execute(command)
}

async function saved(opened: Awaited<ReturnType<typeof pagedEditor>>): Promise<{ bytes: Uint8Array; shd: string[] }> {
	await settleAll(opened.editor)

	const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: opened.editor.sessions })

	if (result.bytes === undefined) throw new Error("no bytes")

	expect(pageDigest(opened.editor.layout)).toEqual(await savedDigest(result.bytes))

	const xml = new TextDecoder().decode(unzipSync(result.bytes)["word/document.xml"])

	return { bytes: result.bytes, shd: xml.match(/<w:shd [^>]*\/>/g) ?? [] }
}

const shading = (editor: PagedEditor): string | null | undefined => editor.snapshot().context.table?.shading

describe("theme table shading", () => {
	it("takes a theme colour with a tint or a shade, never both, and only with a fill", () => {
		expect(parseDocxEditorCommand({ type: "table.cell.shading", fill: "DAE3F3", theme: "accent1", tint: 51 })).toEqual({
			type: "table.cell.shading",
			fill: "DAE3F3",
			theme: "accent1",
			tint: 51
		})
		expect(
			parseDocxEditorCommand({ type: "table.cell.shading", fill: "DAE3F3", theme: "accent1", tint: 51, shade: 10 })
		).toBeUndefined()
		expect(parseDocxEditorCommand({ type: "table.cell.shading", fill: "DAE3F3", tint: 51 })).toBeUndefined()
		expect(parseDocxEditorCommand({ type: "table.cell.shading", fill: null, theme: "accent1" })).toBeUndefined()
		expect(parseDocxEditorCommand({ type: "table.cell.shading", fill: "DAE3F3", theme: "accent9" })).toBeUndefined()
	})

	it("writes w:themeFill with the colour the theme gives, reads it back, and keeps it across other edits", async () => {
		const opened = await pagedEditor(buildDocx(spec(cell("a1"), true)))
		const { editor } = opened

		caretIn(editor, "a1")
		// The fill given is the host's guess; the document's theme decides it.
		expect(await run(editor, { type: "table.cell.shading", fill: "000000", theme: "accent1", tint: 51 })).toBe(true)
		expect(shading(editor)).toBe("theme:accent1/tint:51")

		let out = await saved(opened)
		const [shd] = out.shd

		expect(shd).toMatch(/^<w:shd w:val="clear" w:color="auto" w:fill="([0-9A-F]{6})" w:themeFill="accent1" w:themeFillTint="33"\/>$/)
		expect(shd).not.toContain('w:fill="000000"')

		// Reopened, the picker reads the theme colour; aligning the cell leaves its fill as it is.
		const reopened = await pagedEditor(out.bytes)

		caretIn(reopened.editor, "a1")
		expect(shading(reopened.editor)).toBe("theme:accent1/tint:51")
		expect(await run(reopened.editor, { type: "table.cell.vAlign", align: "center" })).toBe(true)
		expect((await saved(reopened)).shd).toEqual([shd])
		reopened.editor.dispose()

		// A standard colour writes no theme reference; a shade writes w:themeFillShade.
		expect(await run(editor, { type: "table.cell.shading", fill: "ffee00" })).toBe(true)
		expect(shading(editor)).toBe("FFEE00")
		out = await saved(opened)
		expect(out.shd).toEqual(['<w:shd w:val="clear" w:color="auto" w:fill="FFEE00"/>'])
		expect(await run(editor, { type: "table.cell.shading", fill: "2F5496", theme: "accent1", shade: 191 })).toBe(true)
		expect((await saved(opened)).shd[0]).toMatch(/w:themeFill="accent1" w:themeFillShade="BF"\/>$/)
		expect(await run(editor, { type: "table.cell.shading", fill: null })).toBe(true)
		expect(shading(editor)).toBeNull()
		editor.dispose()
	})

	it("records a tracked theme fill change with the fill before it, and rejecting it brings that back", async () => {
		const opened = await pagedEditor(
			buildDocx(
				spec(cell("a1", '<w:shd w:val="clear" w:color="auto" w:fill="E2EFD9" w:themeFill="accent6" w:themeFillTint="33"/>'), true)
			)
		)
		const { editor, context } = opened

		context.identity.author = "Ann"
		context.review.tracking = true
		caretIn(editor, "a1")
		expect(shading(editor)).toBe("theme:accent6/tint:51")
		expect(await run(editor, { type: "table.cell.shading", fill: "FFEE00" })).toBe(true)

		const xml = new TextDecoder().decode(unzipSync((await saved(opened)).bytes)["word/document.xml"])

		expect(xml).toMatch(
			/<w:tcPrChange [^>]*><w:tcPr><w:tcW w:w="3000" w:type="dxa"\/><w:shd w:val="clear" w:color="auto" w:fill="E2EFD9" w:themeFill="accent6" w:themeFillTint="33"\/><\/w:tcPr><\/w:tcPrChange>/
		)
		expect(await run(editor, { type: "review.rejectAll" })).toBe(true)
		caretIn(editor, "a1")
		expect(shading(editor)).toBe("theme:accent6/tint:51")
		editor.dispose()
	})

	it("writes the fill given where the document has no theme", async () => {
		const opened = await pagedEditor(buildDocx(spec(cell("a1"), false)))

		caretIn(opened.editor, "a1")
		expect(await run(opened.editor, { type: "table.cell.shading", fill: "DAE3F3", theme: "accent1", tint: 51 })).toBe(true)
		expect((await saved(opened)).shd).toEqual([
			'<w:shd w:val="clear" w:color="auto" w:fill="DAE3F3" w:themeFill="accent1" w:themeFillTint="33"/>'
		])
		opened.editor.dispose()
	})
})
