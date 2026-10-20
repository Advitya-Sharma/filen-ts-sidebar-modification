import { useEffect, type RefObject } from "react"
import type { SheetGridModel } from "@filen/office/react/sheet"
import { useLatestRef } from "@/lib/useLatestRef"

// Copy, cut and paste while the grid has focus, as tab-separated text. Listened for on the document: Firefox sends
// clipboard events to the body when the focused element holds no text selection, as the grid never does. The
// session holds the cells, so a copy's text arrives after the event: it goes to the clipboard as a pending item.
// A paste of the session's own latest copy names it, so formulas move and formats come along (a cut moves its
// cells).
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

			const written =
				typeof ClipboardItem === "function"
					? navigator.clipboard.write([
							new ClipboardItem({ "text/plain": text.then(value => new Blob([value], { type: "text/plain" })) })
						])
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

			void grid.intent()
			void grid.command({
				type: "clipboard.paste",
				sheet,
				at: { row: range.startRow, col: range.startCol },
				text: data?.getData("text/plain") ?? "",
				...(lastClip !== null && { clip: lastClip })
			})
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
