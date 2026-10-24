import { describe, expect, it } from "vitest"
import { openOdtAsDocx } from "../core/odt/index"
import { bytesSource } from "../core/opc/index"
import { odtBytes, partText, translated } from "./odtFixtures"

// An OpenDocument text's own title (meta.xml dc:title) reaches the package it reads as, as its core properties: the
// PDF export and a .docx copy show it; a document naming none has none (the export titles it from its file name).

const meta = (title: string): { name: string; data: Uint8Array } => ({
	name: "meta.xml",
	data: new TextEncoder().encode(
		`<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.3"><office:meta><dc:title>${title}</dc:title></office:meta></office:document-meta>`
	)
})

describe("an OpenDocument text's title", () => {
	it("is the package's core title", async () => {
		const translation = await translated({
			body: "<text:p>Hi</text:p>",
			extra: [meta("Minutes &amp; notes ")]
		})

		expect(partText(translation, "docProps/core.xml")).toContain("<dc:title>Minutes &amp; notes</dc:title>")
		expect(partText(translation, "_rels/.rels")).toContain('Target="docProps/core.xml"')

		const opened = await openOdtAsDocx(bytesSource(odtBytes({ body: "<text:p>Hi</text:p>", extra: [meta("Minutes")] })))

		expect(opened.ok).toBe(true)
	})

	it("is none when meta.xml names none or a blank one", async () => {
		for (const extra of [[], [meta("  ")]]) {
			const translation = await translated({
				body: "<text:p>Hi</text:p>",
				extra
			})

			expect(translation.parts.has("docProps/core.xml")).toBe(false)
			expect(partText(translation, "_rels/.rels")).not.toContain("core-properties")
		}
	})
})
