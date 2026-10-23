import { describe, expect, it } from "vitest"
import { convertedCopyName } from "@/features/drive/lib/copyName"

describe("convertedCopyName", () => {
	it("takes the file's name with the new extension", async () => {
		expect(await convertedCopyName("report.doc", "docx", () => Promise.resolve(false))).toBe("report.docx")
		expect(await convertedCopyName("Template.DOT", "docx", () => Promise.resolve(false))).toBe("Template.docx")
		expect(await convertedCopyName("noextension", "docx", () => Promise.resolve(false))).toBe("noextension.docx")
	})

	it("counts on past taken names, never reusing one", async () => {
		const taken = new Set(["report.docx", "report (2).docx"])

		expect(await convertedCopyName("report.doc", "docx", name => Promise.resolve(taken.has(name)))).toBe("report (3).docx")
	})
})
