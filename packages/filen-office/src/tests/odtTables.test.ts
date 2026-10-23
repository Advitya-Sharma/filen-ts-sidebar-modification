import { describe, expect, it } from "vitest"
import { documentXml } from "./odtFixtures"

function gridColumns(xml: string): number[] {
	return [...xml.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map(match => Number(match[1]))
}

function tableWidth(xml: string): number {
	return Number(/<w:tblW w:w="(\d+)"/.exec(xml)?.[1] ?? "NaN")
}

const ROW =
	"<table:table-row><table:table-cell><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row>"

describe("ODT tables", () => {
	it("fills the text width less its margins when aligned to the margins", async () => {
		const xml = await documentXml({
			automatic:
				'<style:style style:name="T" style:family="table"><style:table-properties style:width="5cm" table:align="margins" fo:margin-left="1cm" fo:margin-right="1cm"/></style:style><style:style style:name="A" style:family="table-column"><style:table-column-properties style:rel-column-width="1*"/></style:style><style:style style:name="B" style:family="table-column"><style:table-column-properties style:rel-column-width="3*"/></style:style>',
			body: `<table:table table:style-name="T"><table:table-column table:style-name="A"/><table:table-column table:style-name="B"/>${ROW}</table:table>`
		})
		// A4 less 2 cm margins is 17 cm of text; the table's own margins leave 15 cm (8504 twips), shared 1:3.
		const width = Math.round((15 / 2.54) * 1440)

		expect(tableWidth(xml)).toBe(width)
		expect(gridColumns(xml)).toEqual([Math.round(width / 4), Math.round((width * 3) / 4)])
		expect(xml).toContain('<w:tblInd w:w="567" w:type="dxa"/>')
	})

	it("sizes a relative-width table against the text width", async () => {
		const xml = await documentXml({
			automatic:
				'<style:style style:name="T" style:family="table"><style:table-properties style:width="9cm" style:rel-width="50%" table:align="center"/></style:style>',
			body: `<table:table table:style-name="T"><table:table-column table:number-columns-repeated="2"/>${ROW}</table:table>`
		})

		expect(tableWidth(xml)).toBe(Math.round((8.5 / 2.54) * 1440))
		expect(xml).toContain('<w:jc w:val="center"/>')
	})

	it("gives a zero-width table LibreOffice's minimum width per column, shared equally", async () => {
		const xml = await documentXml({
			automatic:
				'<style:style style:name="T" style:family="table"><style:table-properties style:width="0in" table:align="left"/></style:style><style:style style:name="A" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style><style:style style:name="B" style:family="table-column"><style:table-column-properties style:column-width="4in"/></style:style>',
			body: `<table:table table:style-name="T"><table:table-column table:style-name="A"/><table:table-column table:style-name="B"/>${ROW}</table:table>`
		})

		// Measured over UNO on a Word-written ODT: each column 23 twips (0.04 cm), the table their sum.
		expect(gridColumns(xml)).toEqual([23, 23])
		expect(tableWidth(xml)).toBe(46)
	})

	it("merges spanned cells and drops the cells they cover", async () => {
		const xml = await documentXml({
			body: '<table:table><table:table-column table:number-columns-repeated="3"/><table:table-row><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2"><text:p>wide</text:p></table:table-cell><table:covered-table-cell/><table:table-cell><text:p>c</text:p></table:table-cell></table:table-row><table:table-row><table:covered-table-cell/><table:covered-table-cell/><table:table-cell><text:p>f</text:p></table:table-cell></table:table-row></table:table>'
		})
		const rows = xml.split("<w:tr>").slice(1)

		expect(rows[0]).toContain('<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>')
		expect((rows[0]?.match(/<w:tc>/g) ?? []).length).toBe(2)
		expect(rows[1]).toContain('<w:gridSpan w:val="2"/><w:vMerge/>')
		expect((rows[1]?.match(/<w:tc>/g) ?? []).length).toBe(2)
	})

	it("leaves a cell side without a border unset, so the neighbour's border shows between them", async () => {
		const xml = await documentXml({
			automatic:
				'<style:style style:name="A" style:family="table-cell"><style:table-cell-properties fo:border-left="0.5pt solid #000000" fo:border-right="none"/></style:style><style:style style:name="B" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #000000"/></style:style>',
			body: '<table:table><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell table:style-name="A"><text:p>a</text:p></table:table-cell><table:table-cell table:style-name="B"><text:p>b</text:p></table:table-cell></table:table-row></table:table>'
		})

		expect(xml).not.toContain('w:val="nil"')
		expect(xml).toContain('<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
	})

	it("repeats header rows on every page", async () => {
		const xml = await documentXml({
			body: `<table:table><table:table-column/><table:table-header-rows>${ROW}</table:table-header-rows>${ROW}</table:table>`
		})

		expect((xml.match(/<w:tblHeader\/>/g) ?? []).length).toBe(1)
	})
})
