import { describe, expect, it } from "vitest"
import { layerKeyFor, sizesInFile } from "@/features/spreadsheet/lib/sizeRouting.logic"

describe("sizesInFile", () => {
	// A workbook not yet proven keeps its sizes in the file too: a resize waits for the proof like any edit.
	it("is only a workbook this viewer may edit", () => {
		expect(sizesInFile("xlsx", true)).toBe(true)
		expect(sizesInFile("xlsx", false)).toBe(false)
		expect(sizesInFile("csv", true)).toBe(false)
		expect(sizesInFile("xls", false)).toBe(false)
	})
})

describe("layerKeyFor", () => {
	it("keys by the stable id, and by the document for the session without one", () => {
		expect(layerKeyFor("stable-1", "doc-1")).toEqual({ kind: "stable", id: "stable-1" })
		expect(layerKeyFor(undefined, "doc-1")).toEqual({ kind: "session", id: "doc-1" })
	})
})
