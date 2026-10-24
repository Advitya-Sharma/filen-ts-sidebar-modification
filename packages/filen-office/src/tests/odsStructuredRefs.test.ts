import { describe, expect, it } from "vitest"
import { resolveStructured, type SheetTable } from "../core/ods/convert/structured"
import { rewriteStructured, specifierColumns } from "../core/xlsx/refs/index"

// A structured reference's column is the table column of that name exactly, letter case aside (ECMA-376 Part 1
// §18.17.6): a bracketed name keeps the spaces it has, so "Total " and "Total" are two columns.

const TABLE: SheetTable = {
	name: "Data",
	sheet: "S",
	ref: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 },
	headerRows: 1,
	totalsRows: 0,
	columns: ["Total", "Total ", "Funded"]
}
const TABLES = new Map([["DATA", TABLE]])
const HERE = { sheet: "S", row: 2, col: 5 }

describe("a structured reference's column", () => {
	it("is the column named with the same spaces", () => {
		expect(resolveStructured("SUM(Data[[Total ]])", TABLES, HERE)).toBe("SUM($B$2:$B$4)")
		expect(resolveStructured("SUM(Data[[Total]])", TABLES, HERE)).toBe("SUM($A$2:$A$4)")
		expect(resolveStructured('SUMIF(Data[Funded],"Yes",Data[[Total ]])', TABLES, HERE)).toBe('SUMIF($C$2:$C$4,"Yes",$B$2:$B$4)')
	})

	it("matches the name in any letter case", () => {
		expect(resolveStructured("SUM(Data[[TOTAL ]])", TABLES, HERE)).toBe("SUM($B$2:$B$4)")
		expect(resolveStructured("SUM(Data[total])", TABLES, HERE)).toBe("SUM($A$2:$A$4)")
	})

	it("takes the spaces just inside a reference's own brackets as spacing", () => {
		expect(resolveStructured("SUM(Data[ Funded ])", TABLES, HERE)).toBe("SUM($C$2:$C$4)")
		expect(resolveStructured("SUM(Data[ [Total ] ])", TABLES, HERE)).toBe("SUM($B$2:$B$4)")
	})

	it("names no column when no name matches with its spaces", () => {
		expect(resolveStructured("SUM(Data[[Funded ]])", TABLES, HERE)).toBeNull()
	})

	it("is this row's cell of the column, with a bracketed name after @", () => {
		expect(resolveStructured("Data[@[Total ]]*2", TABLES, HERE)).toBe("$B$3*2")
		expect(resolveStructured("Data[@Funded]*2", TABLES, HERE)).toBe("$C$3*2")
		expect(resolveStructured("[@[Total]]", TABLES, { sheet: "S", row: 1, col: 1 })).toBe("$A$2")
	})

	it("breaks on an edit only when the column of that exact name went", () => {
		const tables = new Map([["DATA", { deleted: false as const, columns: new Set(["TOTAL "]) }]])

		expect(specifierColumns("[Total ]")).toEqual(["Total "])
		expect(specifierColumns("[#Headers],[Total ]:[Funded]")).toEqual(["Total ", "Funded"])
		expect(specifierColumns(" Funded ")).toEqual(["Funded"])
		expect(rewriteStructured("SUM(Data[[Total ]])+SUM(Data[[Total]])", null, tables)).toBe("SUM(#REF!)+SUM(Data[[Total]])")
	})
})
