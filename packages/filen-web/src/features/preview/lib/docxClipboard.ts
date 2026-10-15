import { log } from "@/lib/log"
import { htmlHoldsOnlyImages } from "@/features/preview/lib/docxImages.logic"

// The system clipboard for the DOCX editor's own Cut, Copy and Paste commands (menus, touch): the selection
// written as HTML and text (as a copy event writes it), read back as both, or a picture copied alone.
// Keyboard shortcuts and the browser's own menu go through clipboard events instead and never get here.

export interface ClipboardContent {
	html?: string
	text?: string
	// Pictures on the clipboard without text (a screenshot, a copied image).
	files?: File[]
}

// False when the browser refused (no permission, no user activation). Without `html`, text alone.
export async function writeEditorClipboard(content: { html?: string; text: string }): Promise<boolean> {
	try {
		if (content.html !== undefined && typeof ClipboardItem === "function") {
			await navigator.clipboard.write([
				new ClipboardItem({
					"text/html": new Blob([content.html], { type: "text/html" }),
					"text/plain": new Blob([content.text], { type: "text/plain" })
				})
			])
		} else {
			await navigator.clipboard.writeText(content.text)
		}

		return true
	} catch (error) {
		log.warn("docx", "clipboard write refused", { error })

		return false
	}
}

async function readItems(): Promise<ClipboardContent> {
	const out: ClipboardContent = {}

	for (const item of await navigator.clipboard.read()) {
		if (out.html === undefined && item.types.includes("text/html")) {
			out.html = await (await item.getType("text/html")).text()
		}

		if (out.text === undefined && item.types.includes("text/plain")) {
			out.text = await (await item.getType("text/plain")).text()
		}

		// A picture is read only from an item without text, or whose HTML is only pictures (a copied image):
		// copied content keeps its own pictures.
		const image =
			item.types.includes("text/plain") && !htmlHoldsOnlyImages(out.html)
				? undefined
				: item.types.find(type => type.startsWith("image/"))

		if (image !== undefined) {
			const blob = await item.getType(image)

			out.files = [...(out.files ?? []), new File([blob], `image.${image.slice(6).replace(/\W.*$/, "")}`, { type: image })]
		}
	}

	return out
}

// HTML and text where the browser lets the page read both, else text alone; null when it refused.
export async function readEditorClipboard(): Promise<ClipboardContent | null> {
	try {
		return await readItems()
	} catch {
		try {
			return { text: await navigator.clipboard.readText() }
		} catch (error) {
			log.warn("docx", "clipboard read refused", { error })

			return null
		}
	}
}
