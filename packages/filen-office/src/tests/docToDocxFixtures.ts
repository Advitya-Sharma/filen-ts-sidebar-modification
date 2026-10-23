import { bytesSource } from "@filen/office/core/opc"
import { buildDocxModel, Item, openDocxPackage, RevisionView, type DocxModel } from "@filen/office/core/docx"
import { docxPackage, type DocConversionResult } from "../core/doc/toDocx/index"

// A converted package read back through the docx reader.

export async function reopen(result: DocConversionResult): Promise<DocxModel> {
	const opened = await openDocxPackage(bytesSource(await docxPackage(result.parts)))

	if (!opened.ok) throw new Error(`the package does not open: ${opened.reason}`)

	return buildDocxModel(opened.package)
}

// The body's paragraphs (not those of text boxes) as the docx reader reads them, tracked changes shown.
export function docxBodyText(model: DocxModel): string {
	const body = model.body === undefined ? undefined : model.partOf(model.body)
	const out: string[] = []

	if (body === undefined) return ""

	for (let item = 0; item < body.count; item++) {
		if (body.kinds[item] !== Item.Paragraph) continue

		let nested = false

		for (let at = body.parent(item); at >= 0 && !nested; at = body.parent(at))
			nested = body.kinds[at] === Item.Textbox || body.kinds[at] === Item.Drawing

		if (!nested) out.push(body.paragraphText(item, RevisionView.Markup))
	}

	return out.join("\n")
}
