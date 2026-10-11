import type { Node } from "prosemirror-model"
import { bytesSource } from "@filen/office/core/opc"
import {
	Item,
	RevisionView,
	buildDocxModel,
	openDocxPackage,
	type DocxModel,
	type DocxPackage,
	type StoryPart
} from "@filen/office/core/docx"
import {
	DocxDocBuilder,
	DocxEditorSession,
	docxSchema,
	markerKey,
	openDocxEditContext,
	survivorsOf,
	type DocxEditContext,
	type DocxEditorOptions,
	type MarkerList,
	type StoryDoc
} from "@filen/office/core/docx/edit"
import { openSynthetic, p, type SyntheticDocx } from "./docxSynthetic"

// Opening documents for editing in tests, and a structural comparison of an editing document with the
// model it was built from.

export interface Opened {
	readonly pkg: DocxPackage
	readonly model: DocxModel
	readonly context: DocxEditContext
}

export async function openBytes(bytes: Uint8Array): Promise<Opened> {
	const opened = await openDocxPackage(bytesSource(bytes))

	if (!opened.ok) throw new Error(`${opened.reason}: ${opened.message}`)

	const model = await buildDocxModel(opened.package)

	return { pkg: opened.package, model, context: await openDocxEditContext(opened.package, model) }
}

export async function openSpec(spec: SyntheticDocx): Promise<Opened> {
	const pkg = await openSynthetic(spec)
	const model = await buildDocxModel(pkg)

	return { pkg, model, context: await openDocxEditContext(pkg, model) }
}

export async function sessionOf(spec: SyntheticDocx, options: DocxEditorOptions = {}): Promise<DocxEditorSession> {
	return new DocxEditorSession((await openSpec(spec)).context, options)
}

export async function bodyOf(spec: SyntheticDocx): Promise<StoryDoc> {
	const { context, model } = await openSpec(spec)

	if (model.body === undefined) throw new Error("no body")

	return new DocxDocBuilder(context).build(model.body)
}

export const STYLES_PART = (styles: string): { name: string; contentType: string; data: string } => ({
	name: "word/styles.xml",
	contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
	data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles}</w:styles>`
})

export const NUMBERING_PART = (numbering: string): { name: string; contentType: string; data: string } => ({
	name: "word/numbering.xml",
	contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
	data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${numbering}</w:numbering>`
})

// Every range marker and comment reference key in a document, sorted (with repeats).
export function markerKeys(doc: Node): string[] {
	const keys: string[] = []
	const lists = (node: Node): void => {
		for (const name of ["before", "end"]) {
			const list: unknown = node.attrs[name]

			if (Array.isArray(list)) for (const marker of list as MarkerList) keys.push(markerKey(marker))
		}
	}

	lists(doc)
	doc.descendants(node => {
		for (const marker of survivorsOf(node)) keys.push(markerKey(marker))

		if (node.isBlock) lists(node)

		return true
	})

	return keys.sort()
}

// Text of a paragraph node as StoryPart.paragraphText (Markup view) gives it for its source.
export function nodeText(paragraph: Node): string {
	let out = ""

	paragraph.forEach(child => {
		switch (child.type.name) {
			case "text":
				out += child.text ?? ""

				break
			case "tab":
				out += "\t"

				break
			case "break":
				out += child.attrs["type"] === "textWrapping" ? "\n" : "\f"

				break
			case "hyphen":
				out += child.attrs["soft"] === true ? "­" : "‑"

				break
			case "symbol": {
				const char = child.attrs["char"] as number

				if (char > 0 && char <= 0x10ffff) out += String.fromCodePoint(char)

				break
			}
			case "field":
				out += child.attrs["result"] as string

				break
			case "run_object":
			case "opaque_inline":
				out += child.attrs["text"] as string

				break
		}
	})

	return out
}

// Structural differences between an editing document and its model: every source block sits where its
// item does (same order per container, markers aside), kinds match, and editable paragraphs read as their
// source text. Returns at most `limit` problems.
export function structureProblems(story: StoryDoc, limit: number = 5): string[] {
	const problems: string[] = []
	const part = story.part
	const report = (problem: string): void => {
		if (problems.length < limit) problems.push(problem)
	}
	const expectedBlocks = (container: number): number[] => {
		const out: number[] = []

		for (let item = container + 1; item < part.end(container); item = part.end(item)) {
			if (part.kinds[item] !== Item.SectionProperties && !isMarker(part, item)) out.push(item)
		}

		return out
	}
	const visit = (node: Node, container: number): void => {
		const items = expectedBlocks(container)
		const children: Node[] = []

		node.forEach(child => {
			children.push(child)
		})

		if (items.length !== children.length && !(items.length === 0 && children.length === 1)) {
			report(`container ${container}: ${children.length} blocks for ${items.length} items`)

			return
		}

		children.forEach((child, index) => {
			const item = items[index]

			if (item === undefined) return

			const src = child.attrs["src"] as number

			if (src !== item) {
				report(`block ${index} of ${container}: src ${src}, expected ${item}`)

				return
			}

			switch (child.type) {
				case docxSchema.nodes.paragraph: {
					const expected = part.paragraphText(item, RevisionView.Markup)
					const actual = nodeText(child)

					if (expected !== actual)
						report(`paragraph ${item}: ${JSON.stringify(actual.slice(0, 80))} != ${JSON.stringify(expected.slice(0, 80))}`)

					break
				}
				case docxSchema.nodes.table:
					child.forEach((row, _offset, r) => {
						const rowItem = row.attrs["src"] as number

						if (part.kinds[rowItem] !== Item.Row) report(`row ${r} of table ${item} has src ${rowItem}`)

						row.forEach(cell => {
							const cellItem = cell.attrs["src"] as number

							if (part.kinds[cellItem] !== Item.Cell) report(`cell of row ${rowItem} has src ${cellItem}`)
							else visit(cell, cellItem)
						})
					})

					break
				case docxSchema.nodes.sdt:
					visit(child, item)

					break
			}
		})
	}

	visit(story.doc, story.story.item)

	return problems
}

function isMarker(part: StoryPart, item: number): boolean {
	const kind = part.kinds[item] ?? 0

	return kind >= Item.BookmarkStart && kind <= Item.CustomXmlRangeMarker
}

const FIELD = (instruction: string, result: string): string =>
	'<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
	`<w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r>` +
	'<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
	`<w:r><w:t>${result}</w:t></w:r>` +
	'<w:r><w:fldChar w:fldCharType="end"/></w:r>'

// A body with one of everything the editing schema maps.
export const RICH_BODY =
	'<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>' +
	'<w:bookmarkStart w:id="0" w:name="start"/>' +
	'<w:r w:rsidR="00A1"><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>' +
	'<w:r w:rsidR="00B2"><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> more</w:t></w:r>' +
	'<w:r><w:t>plain</w:t><w:tab/><w:br w:type="page"/><w:sym w:font="Wingdings" w:char="F04A"/><w:softHyphen/><w:lastRenderedPageBreak/></w:r>' +
	'<w:bookmarkEnd w:id="0"/>' +
	'<w:hyperlink r:id="rIdL"><w:r><w:t>link</w:t></w:r></w:hyperlink>' +
	'<w:ins w:id="5" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins>' +
	'<w:del w:id="6" w:author="Bob"><w:r><w:delText>gone</w:delText></w:r></w:del>' +
	FIELD(" PAGE ", "3") +
	'<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple>' +
	'<w:r><w:footnoteReference w:id="1"/></w:r>' +
	'<w:commentRangeStart w:id="3"/><w:r><w:t>c</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r>' +
	"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>" +
	"</w:p>" +
	'<w:bookmarkStart w:id="1" w:name="table"/>' +
	'<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
	'<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>' +
	'<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc><w:tc><w:p/><w:bookmarkEnd w:id="1"/></w:tc></w:tr>' +
	"</w:tbl>" +
	'<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>end of section</w:t></w:r></w:p>' +
	'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Entry 1</w:t></w:r></w:p>' +
	'<w:p><w:r><w:t>Entry 2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
	'<w:sdt><w:sdtPr><w:lock w:val="sdtLocked"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>in control</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
	'<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>locked control</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
	'<w:altChunk r:id="rIdA"/>' +
	p("last") +
	'<w:bookmarkStart w:id="2" w:name="end"/><w:bookmarkEnd w:id="2"/>' +
	'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'

export const RICH = {
	body: RICH_BODY,
	relationships: [{ id: "rIdL", type: "hyperlink", target: "https://example.com/", external: true }]
}
