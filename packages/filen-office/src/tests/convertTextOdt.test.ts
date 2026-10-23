import { describe, expect, it } from "vitest"
import { convertText } from "../core/convert/text/index"
import { bytesSource, ZipArchive } from "../core/opc/index"
import { translateOdt } from "../core/odt/index"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { buildDocx, p, type SyntheticDocx } from "./docxSynthetic"

const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"
// A 1×1 PNG.
const PNG = Uint8Array.from(
	atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="),
	char => char.charCodeAt(0)
)

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`

const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\uF0B7"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%1)"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`

function spec(body: string, extra: Partial<SyntheticDocx> = {}): SyntheticDocx {
	return {
		body,
		parts: [
			{ name: "word/styles.xml", contentType: `${WML}.styles+xml`, data: STYLES },
			{ name: "word/numbering.xml", contentType: `${WML}.numbering+xml`, data: NUMBERING },
			{ name: "word/media/image1.png", data: PNG },
			...(extra.parts ?? [])
		],
		relationships: [
			{ id: "rIdS", type: "styles", target: "styles.xml" },
			{ id: "rIdN", type: "numbering", target: "numbering.xml" },
			{ id: "rIdImg", type: "image", target: "media/image1.png" },
			{ id: "rIdLink", type: "hyperlink", target: "https://filen.io/", external: true },
			...(extra.relationships ?? [])
		],
		...Object.fromEntries(Object.entries(extra).filter(([key]) => key !== "parts" && key !== "relationships"))
	}
}

async function odtParts(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
	const archive = await ZipArchive.open(bytesSource(bytes))
	const out = new Map<string, Uint8Array>()

	for (const entry of archive.entries) out.set(entry.name, await archive.read(entry))

	return out
}

// The text of each body paragraph as the ODT translator reads the written file back.
async function readBack(bytes: Uint8Array): Promise<string[]> {
	const translated = await translateOdt(bytesSource(bytes), { strict: true })

	if (!translated.ok) throw new Error(`${translated.reason}: ${translated.message}`)

	const document = translated.translation.parts.get("word/document.xml")
	const xml = document instanceof Uint8Array ? decodeUtf8(document) : ""

	return [...xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/gu)].map(match =>
		[...(match[1] ?? "").matchAll(/<w:(t|delText)\b[^>]*>([^<]*)<\/w:\1>|<w:tab\/>/gu)]
			.map(run => (run[0] === "<w:tab/>" ? "\t" : (run[2] ?? "")))
			.join("")
	)
}

const PICTURE =
	'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1" descr="A dot"/>' +
	'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="dot.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
	'<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'

describe("convertText docx → odt", () => {
	it("writes an ODF 1.3 text whose text reads back as the source's", async () => {
		const body =
			p("Title", '<w:pStyle w:val="Heading1"/>') +
			'<w:p><w:r><w:t xml:space="preserve">Plain </w:t></w:r><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:hyperlink r:id="rIdLink"><w:r><w:t>a link</w:t></w:r></w:hyperlink><w:r><w:footnoteReference w:id="1"/></w:r></w:p>' +
			p("Bullet one", '<w:pStyle w:val="ListBullet"/>') +
			p("Bullet two", '<w:pStyle w:val="ListBullet"/>') +
			p("Letter", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>') +
			'<w:p><w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>' +
			`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>${p("Wide")}</w:tc></w:tr><w:tr><w:tc>${p("A")}</w:tc><w:tc>${p("B")}</w:tc></w:tr></w:tbl>` +
			`<w:p>${PICTURE}<w:ins w:id="5" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>` +
			`<w:sdt><w:sdtPr><w:alias w:val="Box"/></w:sdtPr><w:sdtContent>${p("In a control")}</w:sdtContent></w:sdt>` +
			'<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
		const source = buildDocx(
			spec(body, {
				headers: {
					rIdH: '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
				},
				footnotes:
					'<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> A note.</w:t></w:r></w:p></w:footnote>'
			})
		)
		const converted = await convertText(source, "docx", "odt")

		if (!converted.ok) throw new Error(converted.refusal.message)

		const parts = await odtParts(converted.bytes)
		const content = decodeUtf8(parts.get("content.xml") ?? new Uint8Array())
		const styles = decodeUtf8(parts.get("styles.xml") ?? new Uint8Array())
		const manifest = decodeUtf8(parts.get("META-INF/manifest.xml") ?? new Uint8Array())

		expect([...parts.keys()][0]).toBe("mimetype")
		expect(decodeUtf8(parts.get("mimetype") ?? new Uint8Array())).toBe("application/vnd.oasis.opendocument.text")
		expect(content).toContain('office:version="1.3"')
		// No internal placeholders or WordprocessingML reach the ODF streams.
		expect(content.includes("\u0002") || content.includes("\u0003")).toBe(false)
		expect(styles.includes("\u0002") || styles.includes("\u0003")).toBe(false)
		expect(content).not.toMatch(/<w:|w:val/u)
		expect(styles).not.toMatch(/<w:/u)
		// The heading's style, the bullet list (the Symbol bullet as Unicode) and the letter list (aa, bb...).
		expect(styles).toMatch(
			/<style:style style:name="Heading1" style:display-name="heading 1" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:default-outline-level="1">/u
		)
		expect(styles).toContain('text:bullet-char="•"')
		expect(styles).not.toContain('style:font-name="Symbol"')
		expect(styles).toContain('style:num-format="a" style:num-letter-sync="true" style:num-prefix="(" style:num-suffix=")"')
		expect(content).toMatch(
			/<text:list text:style-name="WWNum0"[^>]*><text:list-item><text:p[^>]*>Bullet one<\/text:p><\/text:list-item><text:list-item><text:p[^>]*>Bullet two/u
		)
		// The page break starts a paragraph of its own on a new page.
		expect(content).toMatch(/>Before<\/text:p><text:p text:style-name="(P\d+)">after<\/text:p>/u)
		expect(content).toContain('<text:note text:note-class="footnote" text:id="filen-footnote-1">')
		expect(content).toMatch(
			/<text:tracked-changes text:track-changes="false"><text:changed-region [^>]*><text:insertion><office:change-info><dc:creator>Ann<\/dc:creator><dc:date>2024-01-02T03:04:05<\/dc:date>/u
		)
		expect(content).toMatch(
			/<draw:frame [^>]*text:anchor-type="as-char" svg:width="2.54cm" svg:height="1.27cm"><draw:image xlink:href="Pictures\/1.png"[^>]*\/><svg:desc>A dot<\/svg:desc><\/draw:frame>/u
		)
		expect(manifest).toContain('manifest:full-path="Pictures/1.png" manifest:media-type="image/png"')
		expect(parts.get("Pictures/1.png")).toEqual(PNG)
		expect(styles).toMatch(
			/<style:header><text:p[^>]*>Page <text:page-number text:select-page="current">1<\/text:page-number><\/text:p><\/style:header>/u
		)
		expect(content).toContain('<text:a xlink:type="simple" xlink:href="https://filen.io/">a link</text:a>')

		expect(await readBack(converted.bytes)).toEqual([
			"Title",
			"Plain bold and a link",
			"Bullet one",
			"Bullet two",
			"Letter",
			"Before",
			"after",
			"Wide",
			"A",
			"B",
			"inserted",
			"In a control"
		])
		expect(converted.lossy.features).toEqual([{ kind: "contentControl", count: 1 }])
	})

	it("adds no empty paragraph for a page break at the start or end of a paragraph (as LibreOffice reads it)", async () => {
		const body =
			p("A") +
			'<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>B</w:t></w:r></w:p>' +
			'<w:p><w:r><w:t>C</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>' +
			p("D") +
			'<w:p><w:r><w:t>E</w:t><w:br w:type="page"/><w:t>F</w:t></w:r></w:p>'
		const converted = await convertText(buildDocx(spec(body)), "docx", "odt")

		if (!converted.ok) throw new Error(converted.refusal.message)

		const content = decodeUtf8((await odtParts(converted.bytes)).get("content.xml") ?? new Uint8Array())
		const breaks = new Set(
			[...content.matchAll(/<style:style style:name="([^"]+)"[^>]*><style:paragraph-properties fo:break-before="page"\/>/gu)].map(
				match => match[1]
			)
		)
		const paragraphs = [...content.matchAll(/<text:p(?: text:style-name="([^"]*)")?(?:\/>|>([^<]*)<\/text:p>)/gu)].map(
			match => `${breaks.has(match[1]) ? "|" : ""}${match[2] ?? ""}`
		)

		expect(paragraphs).toEqual(["A", "|B", "C", "|D", "E", "|F"])
	})

	it("keeps a block's text where the emitter cannot write it, and counts it", async () => {
		const cells = Array.from({ length: 64 }, (_, k) => `<w:tc>${p(`c${String(k)}`)}</w:tc>`).join("")
		const grid = '<w:gridCol w:w="100"/>'.repeat(64)
		const converted = await convertText(
			buildDocx(spec(`<w:tbl><w:tblGrid>${grid}</w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`)),
			"docx",
			"odt"
		)

		if (!converted.ok) throw new Error(converted.refusal.message)

		expect(converted.lossy.features).toContainEqual({ kind: "blockFormat", count: 1 })

		const content = decodeUtf8((await odtParts(converted.bytes)).get("content.xml") ?? new Uint8Array())

		// Every cell's text (the ODT reader shows 63 columns at most, so the written file is read directly).
		expect([...content.matchAll(/<text:p>(c\d+)<\/text:p>/gu)].map(match => match[1])).toEqual(
			Array.from({ length: 64 }, (_, k) => `c${String(k)}`)
		)
	})

	it("counts what it leaves out: macros, run properties, drawings (their text kept), fields kept as results", async () => {
		const body =
			'<w:p><w:r><w:rPr><w:kern w:val="28"/></w:rPr><w:t>kerned</w:t></w:r></w:p>' +
			'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE \\@ "d MMMM yyyy" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2 January 2024</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
			`<w:p><w:r><w:pict><v:shape style="width:100pt;height:50pt"><v:textbox><w:txbxContent>${p("Boxed text")}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`
		const converted = await convertText(
			buildDocx(spec(body, { mainContentType: "application/vnd.ms-word.document.macroEnabled.main+xml" })),
			"docx",
			"odt"
		)

		if (!converted.ok) throw new Error(converted.refusal.message)

		expect(converted.lossy.features).toEqual([
			{ kind: "drawing", count: 1 },
			{ kind: "field", count: 1 },
			{ kind: "macros", count: 1 },
			{ kind: "runFormat", count: 1 }
		])
		expect(await readBack(converted.bytes)).toEqual(["kerned", "2 January 2024", "", "Boxed text"])
	})

	it("refuses what is not a .docx, and an aborted job", async () => {
		expect(await convertText(encodeUtf8("not a zip"), "docx", "odt")).toMatchObject({ ok: false, refusal: { reason: "notSource" } })
		expect(
			await convertText(buildDocx(spec(p("x"))), "docx", "odt", {
				turns: { signal: { aborted: true }, slicing: { pause: () => Promise.resolve(), sliceMs: 0, now: () => 0 } }
			})
		).toMatchObject({ ok: false, refusal: { reason: "aborted" } })
	})
})
