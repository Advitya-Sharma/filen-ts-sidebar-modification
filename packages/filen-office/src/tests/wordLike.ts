// A document.xml shaped like Word's own output (rsids, paragraph ids, full run properties, proofing marks,
// bookmarks), for tests and benchmarks that need realistic multi-megabyte parts without committing them.

const NAMESPACES =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
	'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
	'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"'

const WORDS = [
	"lorem",
	"ipsum",
	"dolor",
	"sit",
	"amet",
	"consectetur",
	"adipiscing",
	"elit",
	"sed",
	"do",
	"eiusmod",
	"tempor",
	"Größe",
	"&amp;",
	"naïve"
]

function rsid(k: number): string {
	return (0x00a00000 + (k % 4096) * 7919).toString(16).toUpperCase().padStart(8, "0")
}

function run(k: number, words: number): string {
	const text: string[] = []

	for (let w = 0; w < words; w++) text.push(WORDS[(k * 7 + w * 3) % WORDS.length] ?? "")

	return (
		`<w:proofErr w:type="spellStart"/><w:r w:rsidRPr="${rsid(k)}"><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>` +
		`<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB"/></w:rPr><w:t xml:space="preserve">${text.join(" ")} </w:t></w:r>`
	)
}

export function wordLikeDocument(paragraphs: number): Uint8Array {
	const parts: string[] = [`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document ${NAMESPACES}><w:body>`]

	for (let p = 0; p < paragraphs; p++) {
		parts.push(
			`<w:p w14:paraId="${rsid(p * 4)}" w14:textId="77777777" w:rsidR="${rsid(p * 4 + 1)}" w:rsidRDefault="${rsid(p * 4 + 2)}" w:rsidP="${rsid(p * 4 + 3)}">` +
				'<w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>' +
				(p % 50 === 0 ? `<w:bookmarkStart w:id="${p}" w:name="_Toc${p}"/>` : "") +
				run(p, 12) +
				run(p + 1, 9) +
				run(p + 2, 14) +
				(p % 50 === 0 ? `<w:bookmarkEnd w:id="${p}"/>` : "") +
				"</w:p>"
		)
	}

	parts.push(
		'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
	)

	return new TextEncoder().encode(parts.join(""))
}
