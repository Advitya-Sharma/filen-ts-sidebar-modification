import type { DrawingImage } from "../../../core/drawingml/index"
import type { DocxObjectData, DocxSessionApi } from "../../../dom/docx/protocol"
import type { DocxMediaCache } from "../media"
import { drawingImages, imageMime, paintDrawing } from "./drawing"
import { paintMath } from "./math"
import type { PaintedObject } from "./page"

// Fills a painted page's object boxes: drawings as SVG (their images read from the package first), equations
// as MathML, separators as rules; anything that cannot be drawn becomes a placeholder box. Returns the
// cleanup that releases the page's images.

let drawingCounter = 0

function placeholder(element: HTMLElement): void {
	element.classList.add("fo-docx-placeholder")
}

function mediaKey(image: DrawingImage): { partName: string; type: string } | undefined {
	const partName = image.svgPartName ?? image.partName
	const type = image.svgPartName !== undefined ? "image/svg+xml" : imageMime(image)

	return partName === undefined || type.length === 0 ? undefined : { partName, type }
}

export interface ObjectPaintOptions {
	readonly document: Document
	readonly session: DocxSessionApi
	readonly media: DocxMediaCache
	// Stops painting when the page goes away first.
	readonly signal: { readonly aborted: boolean }
}

interface Holder {
	// Records an acquired image; false (and released at once) when the page already went away.
	hold(key: { partName: string; type: string }): boolean
}

async function paintDrawingObject(
	object: PaintedObject,
	data: Extract<DocxObjectData, { kind: "drawing" }>,
	options: ObjectPaintOptions,
	holder: Holder
): Promise<void> {
	const images = drawingImages(data.drawing)
	const urls = new Map<string, string | undefined>()

	await Promise.all(
		images.map(async image => {
			const key = mediaKey(image)

			if (key === undefined || urls.has(`${key.type}\u0000${key.partName}`)) return

			urls.set(`${key.type}\u0000${key.partName}`, undefined)

			const url = await options.media.acquire(key.partName, key.type)

			if (holder.hold(key)) urls.set(`${key.type}\u0000${key.partName}`, url)
		})
	)

	if (options.signal.aborted) return

	const svg = paintDrawing(data.drawing, {
		document: options.document,
		idPrefix: `fod${++drawingCounter}-`,
		imageUrl(image) {
			const key = mediaKey(image)

			return key === undefined ? undefined : urls.get(`${key.type}\u0000${key.partName}`)
		}
	})

	object.element.appendChild(svg)
}

export function paintObjects(objects: readonly PaintedObject[], options: ObjectPaintOptions): () => void {
	const held: { partName: string; type: string }[] = []
	let released = false
	const release = (): void => {
		if (released) return

		released = true

		for (const key of held) options.media.release(key.partName, key.type)

		held.length = 0
	}
	const holder: Holder = {
		hold(key) {
			if (released) {
				options.media.release(key.partName, key.type)

				return false
			}

			held.push(key)

			return true
		}
	}

	if (objects.length > 0) {
		const byId = new Map<number, PaintedObject[]>()

		for (const object of objects) {
			const list = byId.get(object.id)

			if (list === undefined) byId.set(object.id, [object])
			else list.push(object)
		}

		void options.session
			.objects([...byId.keys()])
			.then(async results => {
				const work: Promise<void>[] = []

				for (const result of results) {
					for (const object of byId.get(result.id) ?? []) {
						if (options.signal.aborted) return

						switch (result.kind) {
							case "separator":
								object.element.classList.add("fo-docx-separator")

								break
							case "math":
								try {
									const math = paintMath(result.tree, options.document)

									object.element.style.fontSize = `${Math.min(Math.max(result.fontSize, 1), 1000)}px`
									object.element.classList.add("fo-docx-math")
									object.element.appendChild(math)
								} catch {
									placeholder(object.element)
								}

								break
							case "drawing":
								work.push(
									paintDrawingObject(object, result, options, holder).catch(() => {
										placeholder(object.element)
									})
								)

								break
							case "placeholder":
								placeholder(object.element)

								break
						}
					}
				}

				await Promise.all(work)
			})
			.catch(() => {
				for (const object of objects) placeholder(object.element)
			})
	}

	return release
}
