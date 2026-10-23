import { createReadStream } from "node:fs"
import { Readable } from "node:stream"
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs"
import { RevisionView } from "@filen/office/core/docx"
import { writeDocxPdf, type DocxPdfOptions, type DocxPdfResult } from "../core/docx/pdf/write"
import type { DocxPdfDocument } from "../core/docx/pdf/types"
import type { PdfByteStream, PdfFontAssets } from "../core/pdfwrite/fonts/assets"
import { docxPdfSnapshot } from "../dom/docx/pdfSnapshot"
import { buildDocx, type SyntheticDocx } from "./docxSynthetic"
import { layOut } from "./layoutFixtures"

// DOCX → PDF in node: the synthetic document laid out with real metrics (layoutFixtures), snapshotted as the
// export does, painted, and read back with pdf.js.

const ASSETS = new URL("../../assets/fonts/", import.meta.url)

export const pdfAssets: PdfFontAssets = {
	open: file => Promise.resolve(Readable.toWeb(createReadStream(new URL(file, ASSETS))) as unknown as PdfByteStream)
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
</w:styles>`

const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`

const CORE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Quarterly report</dc:title></cp:coreProperties>`

export function paragraph(text: string, style?: string): string {
	return `<w:p>${style === undefined ? "" : `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
}

// A document with every structure the export tags: headings, a list, a table with a header row, a footnote,
// a header and footer, an external link and an internal one to a bookmark.
export function structuredDocx(options: { readonly title?: boolean; readonly extraBody?: string } = {}): Uint8Array {
	const spec: SyntheticDocx = {
		body:
			paragraph("Introduction", "Heading1") +
			'<w:p><w:r><w:t xml:space="preserve">See the </w:t></w:r><w:hyperlink r:id="rIdLink"><w:r><w:t>website</w:t></w:r></w:hyperlink>' +
			'<w:r><w:t xml:space="preserve"> or </w:t></w:r><w:hyperlink w:anchor="details"><w:r><w:t>the details</w:t></w:r></w:hyperlink>' +
			'<w:r><w:t xml:space="preserve"> below.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>' +
			'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First item</w:t></w:r></w:p>' +
			'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Second item</w:t></w:r></w:p>' +
			paragraph("Details", "Heading2")
				.replace("<w:r>", '<w:bookmarkStart w:id="0" w:name="details"/><w:r>')
				.replace("</w:p>", '<w:bookmarkEnd w:id="0"/></w:p>') +
			'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
			"<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>" +
			"<w:tr><w:tc><w:p><w:r><w:t>Alpha</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>One</w:t></w:r></w:p></w:tc></w:tr>" +
			"</w:tbl>" +
			paragraph("Closing words.") +
			(options.extraBody ?? "") +
			'<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>',
		headers: { rIdH: paragraph("Running header") },
		footers: { rIdF: paragraph("Running footer") },
		footnotes:
			'<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
			'<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
			'<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> A note.</w:t></w:r></w:p></w:footnote>',
		relationships: [
			{ id: "rIdLink", type: "hyperlink", target: "https://example.com/report", external: true },
			{ id: "rIdStyles", type: "styles", target: "styles.xml" },
			{ id: "rIdNumbering", type: "numbering", target: "numbering.xml" }
		],
		parts: [
			{
				name: "word/styles.xml",
				contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
				data: STYLES
			},
			{
				name: "word/numbering.xml",
				contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
				data: NUMBERING
			},
			...(options.title === false
				? []
				: [{ name: "docProps/core.xml", contentType: "application/vnd.openxmlformats-package.core-properties+xml", data: CORE }])
		],
		...(options.title === false
			? {}
			: {
					packageRelationships: [
						{
							id: "rIdCore",
							type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
							target: "docProps/core.xml"
						}
					]
				})
	}

	return buildDocx(spec)
}

// The document laid out (real metrics) and snapshotted as the export does.
export async function snapshotOf(
	bytes: Uint8Array,
	options: { readonly title?: string; readonly copy?: boolean } = {}
): Promise<DocxPdfDocument> {
	const layout = await layOut(bytes, { view: RevisionView.Final })
	const { model, resolver, store, fonts } = layout.loaded

	return docxPdfSnapshot(
		{
			model,
			layout,
			fonts,
			resolver,
			store,
			view: RevisionView.Final,
			objects: ids => Promise.resolve(ids.map(id => ({ kind: "placeholder", id, reason: "test" }) as const)),
			media: () => Promise.resolve(null)
		},
		options
	)
}

export async function pdfOf(
	document: DocxPdfDocument,
	options: Partial<DocxPdfOptions> = {}
): Promise<DocxPdfResult & { bytes: Uint8Array }> {
	const result = await writeDocxPdf(document, { assets: pdfAssets, deterministic: true, ...options })
	const chunks: Uint8Array[] = []

	for (const part of result.parts) chunks.push(part instanceof Uint8Array ? part : new Uint8Array(await part.arrayBuffer()))

	const bytes = new Uint8Array(result.size)
	let at = 0

	for (const chunk of chunks) {
		bytes.set(chunk, at)
		at += chunk.length
	}

	return { ...result, bytes }
}

export function openPdf(bytes: Uint8Array, password?: string): Promise<pdfjs.PDFDocumentProxy> {
	return pdfjs.getDocument({ data: new Uint8Array(bytes), verbosity: 0, ...(password === undefined ? {} : { password }) }).promise
}

export async function pageText(pdf: pdfjs.PDFDocumentProxy, page: number): Promise<string> {
	const content = await (await pdf.getPage(page)).getTextContent()

	return content.items.map(item => ("str" in item ? item.str : "")).join("")
}

// Structure element roles of a page's tree, depth first.
export async function structRoles(pdf: pdfjs.PDFDocumentProxy, page: number): Promise<string[]> {
	const tree = (await (await pdf.getPage(page)).getStructTree()) as unknown as { role?: string; children?: unknown[] } | null
	const out: string[] = []
	const visit = (node: unknown): void => {
		if (typeof node !== "object" || node === null) return

		const element = node as { role?: string; children?: unknown[] }

		if (typeof element.role === "string") out.push(element.role)

		for (const child of element.children ?? []) visit(child)
	}

	visit(tree)

	return out
}
