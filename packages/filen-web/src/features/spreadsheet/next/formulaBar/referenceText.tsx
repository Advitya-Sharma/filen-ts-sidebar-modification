import { useLayoutEffect, useRef } from "react"
import type { HighlightedReference } from "@filen/office/core/sheet"

// An entry's text with its references in their colours (the ones their outlines take on the grid), as the formula
// bar and the in-cell editor draw it under their own clear text. The text stays one run, coloured by a background
// clipped to its glyphs with a layer on each reference's boxes as laid out: an element around each reference would
// round its edges and drift the glyphs off the editor's, a pixel every few dozen references.
export function ReferenceText({
	text,
	references,
	colors
}: {
	text: string
	references: readonly HighlightedReference[]
	colors: readonly string[]
}) {
	const ref = useRef<HTMLDivElement>(null)

	useLayoutEffect(() => {
		const element = ref.current
		const node = element?.firstChild

		if (element === null || !(node instanceof Text)) return

		const origin = element.getBoundingClientRect()
		const range = document.createRange()
		const images: string[] = []
		const positions: string[] = []
		const sizes: string[] = []

		for (const reference of references) {
			const color = colors[reference.color % colors.length] ?? "currentColor"

			range.setStart(node, reference.start)
			range.setEnd(node, reference.end)

			for (const box of range.getClientRects()) {
				images.push(`linear-gradient(${color}, ${color})`)
				positions.push(`${String(box.left - origin.left)}px ${String(box.top - origin.top)}px`)
				sizes.push(`${String(box.width)}px ${String(box.height)}px`)
			}
		}

		// The rest of the text in the editor's own colour, under the references.
		element.style.backgroundImage = [...images, "linear-gradient(currentColor, currentColor)"].join(", ")
		element.style.backgroundPosition = [...positions, "0 0"].join(", ")
		element.style.backgroundSize = [...sizes, "100% 100%"].join(", ")
	})

	return (
		<div
			ref={ref}
			className="w-max min-w-full [background-clip:text] bg-no-repeat [-webkit-background-clip:text] [-webkit-text-fill-color:transparent]"
		>
			{text}
		</div>
	)
}
