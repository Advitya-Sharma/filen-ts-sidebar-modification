import { layoutDocument } from "@filen/office/core/docx"
import { HmtxMeasurer, loadDocx } from "../layoutFixtures"

// Page counts for the layout report (never a gate: LibreOffice lays out differently from Word, and so may we).

// Page objects in a PDF LibreOffice wrote (its page dictionaries are never inside compressed object streams).
export function pdfPageCount(bytes: Uint8Array): number {
	const text = new TextDecoder("latin1").decode(bytes)
	let count = 0

	for (const _ of text.matchAll(/\/Type\s*\/Page(?![s\w])/g)) count++

	return count
}

// Word's own page count when it last saved the file: the page breaks it rendered in the body, plus one.
export function wordPageCount(documentXml: string): number {
	let count = 1

	for (const _ of documentXml.matchAll(/<w:lastRenderedPageBreak\b/g)) count++

	return count
}

export async function ourPageCount(bytes: Uint8Array): Promise<number> {
	const loaded = await loadDocx(bytes)
	const layout = layoutDocument(loaded.model, loaded.resolver, loaded.fonts, new HmtxMeasurer(), {
		sliceMs: Infinity,
		store: loaded.store
	})

	await layout.done

	return layout.pages.length
}
