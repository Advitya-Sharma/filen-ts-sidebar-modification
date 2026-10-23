import { describe, expect, it } from "vitest"
import { convertText } from "../core/convert/text/index"
import { bytesSource, ZipArchive } from "../core/opc/index"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { editRtf, latin1, saveRtf } from "./rtfWriteFixtures"

// Text in text boxes reaches the .docx: an RTF shape's \shptxt (alone or in a \shpgrp group) as positioned
// (framed) paragraphs, as the Word 6/95 \do text boxes already are.

async function body(bytes: Uint8Array): Promise<string> {
	const archive = await ZipArchive.open(bytesSource(bytes))
	const entry = archive.get("word/document.xml")

	return entry === undefined ? "" : decodeUtf8(await archive.read(entry))
}

function rtf(text: string): Uint8Array {
	return encodeUtf8(`{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}${text}}`)
}

const BOX = (text: string, geometry = "\\shpleft1000\\shptop500\\shpright4000\\shpbottom1500") =>
	`{\\shp{\\*\\shpinst${geometry}\\shpbxcolumn\\shpbypara\\shpwr3{\\sp{\\sn shapeType}{\\sv 202}}{\\shptxt \\pard ${text}\\par}}}`

describe("text box text", () => {
	it("writes an RTF text box's paragraphs as framed paragraphs", async () => {
		const converted = await convertText(rtf(`\\pard Before${BOX("Boxed text")}\\par After\\par`), "rtf", "docx")

		if (!converted.ok) throw new Error(converted.refusal.message)

		const xml = await body(converted.bytes)

		expect(xml).toMatch(/<w:framePr [^>]*w:w="3000"[^>]*\/>[^]*?Boxed text/u)
		expect(xml).toMatch(/Before[^]*After/u)
	})

	it("relates a picture in an RTF text box from the part the box is written to", async () => {
		const png =
			"89504e470d0a1a0a0000000d4948445200000001000000010806000000" +
			"1f15c4890000000d4944415478da63fccfc0f01f0005fe02fe0dd8b0a10000000049454e44ae426082"
		const converted = await convertText(
			rtf(`\\pard A${BOX(`{\\pict\\pngblip\\picw1\\pich1\\picwgoal240\\pichgoal240 ${png}}`)}\\par`),
			"rtf",
			"docx"
		)

		if (!converted.ok) throw new Error(converted.refusal.message)

		const archive = await ZipArchive.open(bytesSource(converted.bytes))
		const rels = decodeUtf8(await archive.read(archive.get("word/_rels/document.xml.rels") ?? archive.entries[0]!))
		const ids = [...(await body(converted.bytes)).matchAll(/r:embed="([^"]+)"/gu)].map(match => match[1] ?? "")

		expect(ids.length).toBe(1)
		expect(rels).toContain(`Id="${ids[0] ?? ""}"`)
	})

	it("writes RTF WordArt's text as a framed paragraph", async () => {
		const wordArt =
			"{\\shp{\\*\\shpinst\\shpleft0\\shptop0\\shpright5000\\shpbottom2000\\shpbxmargin\\shpbymargin\\shpwr3{\\sp{\\sn shapeType}{\\sv 136}}{\\sp{\\sn gtextUNICODE}{\\sv DRAFT}}}}"
		const converted = await convertText(rtf(`\\pard Body${wordArt}\\par`), "rtf", "docx")

		if (!converted.ok) throw new Error(converted.refusal.message)

		expect(await body(converted.bytes)).toMatch(/<w:framePr [^>]*\/>[^]*?DRAFT/u)
	})

	it("writes the text boxes of an RTF shape group", async () => {
		const group = `{\\shpgrp{\\*\\shpinst\\shpleft0\\shptop0\\shpright6000\\shpbottom3000{\\sp{\\sn shapeType}{\\sv 0}}${BOX("First box")}${BOX("Second box")}}}`
		const converted = await convertText(rtf(`\\pard Text${group}\\par`), "rtf", "docx")

		if (!converted.ok) throw new Error(converted.refusal.message)

		const xml = await body(converted.bytes)

		expect(xml).toContain("First box")
		expect(xml).toContain("Second box")
	})

	// The editor opens an RTF only when its proof reads every block back: a stand-alone WordArt shape's paragraph is
	// one of them, and an untouched save is the file's bytes.
	it("opens an RTF with a stand-alone WordArt shape for editing, an untouched save its bytes", async () => {
		const geometry = "\\shpleft1000\\shptop500\\shpright4000\\shpbottom1500\\shpbxcolumn\\shpbypara\\shpwr3"
		const art = `{\\shp{\\*\\shpinst${geometry}{\\sp{\\sn shapeType}{\\sv 136}}{\\sp{\\sn gtextUNICODE}{\\sv ArtWords}}}`

		for (const text of [
			`\\pard Before${art}{\\shprslt{\\*\\do\\dobxcolumn\\dobypara\\dodhgt8192\\dprect\\dpx0\\dpy0\\dpxsize100\\dpysize100}}}\\par After\\par`,
			`\\pard Before\\par ${art}}\\par After\\par`
		]) {
			const bytes = latin1(`{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}${text}}`)
			const editing = await editRtf(bytes)

			expect(await body(editing.source.read.docx)).toContain("ArtWords")
			expect((await saveRtf(editing)).bytes).toEqual(bytes)
		}
	})
})
