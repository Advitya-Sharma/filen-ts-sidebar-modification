import { useEffect, type RefObject } from "react"
import { MAX_HTML_COPY_CELLS } from "@filen/office/core/sheet"
import type { SheetGridModel } from "@filen/office/react/sheet"
import { useLatestRef } from "@/lib/useLatestRef"
import { MAX_PASTE_HTML, sanitizeSheetClipboard } from "@/features/spreadsheet/next/clipboard/sheetClipboardSanitizer"

// Copy, cut and paste while the grid has focus, as tab-separated text and HTML (a table with values, formulas and
// formats, which Excel, Google Sheets and LibreOffice read). Listened for on the document: Firefox sends clipboard
// events to the body when the focused element holds no text selection, as the grid never does. The session holds the
// cells, so a copy's text and HTML arrive after the event: they go to the clipboard as pending items. A paste of the
// session's own latest copy names it, so formulas move and formats come along (a cut moves its cells); other HTML is
// sanitised and read by the session, else the text is pasted.
export function useSheetClipboard({
	model,
	container,
	canEdit,
	onRefused
}: {
	model: SheetGridModel
	container: RefObject<HTMLElement | null>
	canEdit: boolean
	// A copy too large, or a clipboard the browser would not write.
	onRefused: (reason: "tooLarge" | "failed") => void
}): void {
	const latest = useLatestRef({ model, canEdit, onRefused })

	useEffect(() => {
		let lastClip: number | null = null

		function focusedGrid(): boolean {
			const active = document.activeElement

			return active !== null && active.getAttribute("role") === "grid" && container.current?.contains(active) === true
		}

		function copy(cut: boolean): void {
			const { model: grid, onRefused: refused } = latest.current
			const { sheet, selection } = grid.getSnapshot()
			const range = selection.ranges[selection.ranges.length - 1]

			if (range === undefined) {
				return
			}

			const text = grid.session.query("copy", { sheet, range, cut }).then(answer => {
				if ("refused" in answer) {
					throw new RangeError("tooLarge")
				}

				lastClip = answer.clip

				return answer.text
			})
			const cells = (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1)
			// Decided now, inside the gesture: an item the clipboard is promised must arrive.
			const html =
				cells <= MAX_HTML_COPY_CELLS
					? grid.session.query("copyHtml", { sheet, range }).then(answer => {
							if (answer.html === null) {
								throw new RangeError("tooLarge")
							}

							return sanitizeSheetClipboard(answer.html)
						})
					: null

			const written =
				typeof ClipboardItem === "function"
					? navigator.clipboard
							.write([
								new ClipboardItem({
									"text/plain": text.then(value => new Blob([value], { type: "text/plain" })),
									...(html !== null && { "text/html": html.then(value => new Blob([value], { type: "text/html" })) })
								})
							])
							// HTML too long for the clipboard: the text alone (a text that failed fails again here).
							.catch(() => text.then(value => navigator.clipboard.writeText(value)))
					: text.then(value => navigator.clipboard.writeText(value))

			written.catch((error: unknown) => {
				refused(error instanceof RangeError ? "tooLarge" : "failed")
			})
		}

		function paste(data: DataTransfer | null): void {
			const { model: grid } = latest.current
			const { sheet, selection } = grid.getSnapshot()
			const range = selection.ranges[selection.ranges.length - 1]

			if (range === undefined) {
				return
			}

			const text = data?.getData("text/plain") ?? ""
			const html = data?.getData("text/html") ?? ""
			const at = { row: range.startRow, col: range.startCol }
			const clip = lastClip === null ? {} : { clip: lastClip }

			void grid.intent()
			void grid.command(
				html === "" || html.length > MAX_PASTE_HTML
					? { type: "clipboard.paste", sheet, at, text, ...clip }
					: { type: "clipboard.pasteHtml", sheet, at, text, html: sanitizeSheetClipboard(html), ...clip }
			)
		}

		function listener(event: ClipboardEvent): void {
			if (event.defaultPrevented || !focusedGrid()) {
				return
			}

			event.preventDefault()

			switch (event.type) {
				case "copy":
					copy(false)
					break
				case "cut":
					if (latest.current.canEdit) {
						copy(true)
					}
					break
				case "paste":
					if (latest.current.canEdit) {
						paste(event.clipboardData)
					}
					break
			}
		}

		document.addEventListener("copy", listener)
		document.addEventListener("cut", listener)
		document.addEventListener("paste", listener)

		return () => {
			document.removeEventListener("copy", listener)
			document.removeEventListener("cut", listener)
			document.removeEventListener("paste", listener)
		}
	}, [latest, container])
}
