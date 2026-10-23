import { describe, expect, it } from "vitest"
import { docxFormatOf } from "@/features/preview/components/docxViewer.logic"

describe("docxFormatOf", () => {
	it("opens OpenDocument text and its templates as odt, everything else as docx", () => {
		expect(docxFormatOf("Report.ODT")).toBe("odt")
		expect(docxFormatOf("letter.ott")).toBe("odt")
		expect(docxFormatOf("a.docx")).toBe("docx")
		expect(docxFormatOf("a.odt.docx")).toBe("docx")
		expect(docxFormatOf("noextension")).toBe("docx")
	})

	it("opens RTF as rtf (read in the docx worker, edited and saved as RTF)", () => {
		expect(docxFormatOf("Letter.RTF")).toBe("rtf")
		expect(docxFormatOf("notes.rtf")).toBe("rtf")
	})
})
