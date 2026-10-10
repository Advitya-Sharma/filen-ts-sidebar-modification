import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { Item, buildDocxModel, gateReasons, modelStats, openDocxPackage, type DocxModel } from "@filen/office/core/docx"
import { modelProblems } from "./docxIntegrity"

// Small licensed real-world files (see fixtures/docx-corpus/README.md).
const FIXTURES = join(import.meta.dirname, "fixtures", "docx-corpus")

async function load(name: string): Promise<DocxModel> {
	const opened = await openDocxPackage(bytesSource(new Uint8Array(readFileSync(join(FIXTURES, name)))))

	if (!opened.ok) throw new Error(`${name}: ${opened.reason}`)

	const model = await buildDocxModel(opened.package)

	expect(modelProblems(model)).toEqual([])

	return model
}

describe("real documents", () => {
	it("reads a template with nested tables", async () => {
		const model = await load("oxsdk__BingoSheet_TP10220238.dotx")

		expect(model.package.facts).toMatchObject({ variant: "dotx", template: true })
		expect(modelStats(model)).toMatchObject({ tables: 8, nestedTables: 6, maxTableDepth: 1 })
	})

	it("reads Strict documents with SmartArt, charts and equations", async () => {
		const smartArt = await load("docx4j__strict-smartart.docx")
		const chart = await load("docx4j__strict-chart.docx")
		const math = await load("docx4j__strict-math.docx")
		const format = await load("mammoth__strict-format.docx")
		const vertical = await load("oxsdk__textDirection-tbRl.docx")

		for (const model of [smartArt, chart, math, format, vertical]) {
			expect(model.strict).toBe(true)
			expect(gateReasons(model.package.facts)).toContain("strict")
		}

		const diagram = smartArt.drawings.find(drawing => drawing.graphic === "smartArt")

		expect(diagram?.diagram?.data).toBeDefined()
		expect(await smartArt.smartArtDrawingPart(diagram ?? (undefined as never))).toMatch(/^word\/diagrams\/drawing\d*\.xml$/)

		const chartDrawing = chart.drawings.find(drawing => drawing.graphic === "chart")

		expect(chart.chartPart(chartDrawing ?? (undefined as never))).toMatch(/^word\/charts\/chart\d*\.xml$/)
		expect(modelStats(math).kinds).toMatchObject({ MathParagraph: 1 })
		expect(format.body === undefined ? "" : format.partOf(format.body).paragraphText(format.body.item)).not.toBe("")
	})

	it("reads comments with extended data and right-to-left authors", async () => {
		const model = await load("oxsdk__Comment028.docx")

		expect(model.comments).toHaveLength(15)
		expect(model.comments.every(comment => comment.anchor >= 0 && comment.paraId !== undefined)).toBe(true)
		expect(model.comments[1]?.author).toBe("ارمان(ArmanAg)")
	})

	it("reads tracked moves, format changes and a heavily revised document", async () => {
		const moves = await load("oxpt__RP015-MoveFrom-MoveTo.docx")
		const format = await load("oxsdk__trackChange.docx")
		const revised = await load("oxpt__RA001-Tracked-Revisions-02.docx")

		expect(modelStats(moves).revisions).toMatchObject({ MoveFrom: 1, MoveTo: 1, MoveFromRange: 1, MoveToRange: 1 })
		expect(modelStats(format).revisions).toEqual({ RunProperties: 1 })
		expect(modelStats(revised).revisions).toMatchObject({
			Insertion: 36,
			Deletion: 33,
			MoveFrom: 15,
			MoveTo: 15,
			RowInsertion: 10,
			RowDeletion: 10
		})
		expect(revised.package.facts.documentProtection).toMatchObject({ edit: "readOnly", enforced: false })
		expect(gateReasons(revised.package.facts)).toEqual([])
	})

	it("reads shapes and groups with text boxes", async () => {
		const model = await load("poi__shapes-with-text.docx")
		const stats = modelStats(model)

		expect(stats.drawings).toEqual({ shape: 5, group: 2 })
		expect(stats.stories["textbox"]).toBe(5)
		expect(stats.alternateContent).toMatchObject({ total: 7, choice: 7 })
	})

	it("reads a watermark and content controls", async () => {
		const model = await load("oxpt__Watermark-2.docx")

		expect(model.vml).toHaveLength(1)
		expect(modelStats(model).sdts.run).toBe(1)
	})

	it("resolves first-page and even-page headers", async () => {
		const evenOdd = await load("pydocx__doc-odd-even-hdrs.docx")
		const firstPage = await load("poi__DiffFirstPageHeadFoot.docx")
		const [section] = evenOdd.sections

		expect(section?.headers.default).toBeGreaterThanOrEqual(0)
		expect(section?.headers.first).toBeGreaterThanOrEqual(0)
		expect(section?.headers.even).toBeGreaterThanOrEqual(0)
		expect(evenOdd.stories[section?.headers.even ?? -1]?.kind).toBe("header")
		expect(firstPage.sections[0]?.footers.first).toBeGreaterThanOrEqual(0)
	})

	it("keeps altChunk as a block", async () => {
		const model = await load("oxpt__CA009-altChunk.docx")

		expect(modelStats(model).kinds["AltChunk"]).toBe(1)
	})

	it("reads footnotes and endnotes", async () => {
		const model = await load("oxsdk__ftEdNote.docx")
		const body = model.body

		expect(model.footnotes.size).toBe(4)
		expect(model.endnotes.size).toBe(4)

		if (body === undefined) throw new Error("no body")

		const part = model.partOf(body)
		let references = 0

		for (let item = 0; item < part.count; item++) {
			const kind = part.kinds[item]

			if (kind === Item.FootnoteReference || kind === Item.EndnoteReference) {
				references++
				expect((kind === Item.FootnoteReference ? model.footnotes : model.endnotes).has(String(part.aux[item]))).toBe(true)
			}
		}

		expect(references).toBeGreaterThan(0)
	})

	it("reports signatures, macros and glossaries", async () => {
		const signed = await load("poi__hello-world-signed.docx")
		const macro = await load("poi__SimpleMacro.docm")
		const normal = await load("docx4j__Normal.dotm")

		expect(gateReasons(signed.package.facts)).toEqual(["signed"])
		expect(macro.package.facts).toMatchObject({ variant: "docm", macros: true })
		expect(gateReasons(macro.package.facts)).toEqual([])
		expect(normal.package.facts).toMatchObject({ variant: "dotm", macros: true, glossary: true })
	})

	it("reads a complicated document end to end", async () => {
		const model = await load("oxpt__HC031-Complicated-Document.docx")
		const stats = modelStats(model, { fieldTypes: true })

		expect(stats.stories).toMatchObject({ body: 1, header: 3, footer: 3, textbox: 3, comment: 1 })
		expect(stats.sections).toBe(4)
		expect(stats.fieldTypes).toMatchObject({ TOC: 1, PAGEREF: 4 })
		expect(stats.sdts.block).toBeGreaterThan(0)
		expect(model.sections[3]?.effectiveHeaders.default).toBe(model.sections[0]?.headers.default)
	})

	it("reports password-protected files without parsing them", async () => {
		const opened = await openDocxPackage(
			bytesSource(new Uint8Array(readFileSync(join(FIXTURES, "oxsdk__document_with_password.docx"))))
		)

		expect(opened).toMatchObject({ ok: false, reason: "encrypted" })
	})
})
