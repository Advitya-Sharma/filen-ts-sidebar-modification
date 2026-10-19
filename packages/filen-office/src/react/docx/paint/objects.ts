import { shadeRaster, type DrawingImage, type Paint } from "../../../core/drawingml/index"
import { encodePng } from "../../../core/metafile/dib"
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

// Advance widths through a canvas of the document (text a drawing wraps); NaN without one (the painter
// then estimates).
function textMeasurer(document: Document): (text: string, font: string) => number {
	let context: CanvasRenderingContext2D | null | undefined

	return (text, font) => {
		if (context === undefined) {
			try {
				context = document.createElement("canvas").getContext("2d")
			} catch {
				context = null
			}
		}

		if (context === null) return Number.NaN

		context.font = font

		return context.measureText(text).width
	}
}

export interface ObjectPaintOptions {
	readonly document: Document
	readonly session: DocxSessionApi
	readonly media: DocxMediaCache
	// Stops painting when the page goes away first.
	readonly signal: { readonly aborted: boolean }
}

// Pixels of one shaded raster, and of all a page's shaded rasters over its life.
const MAX_SHADE_PIXELS = 2048 * 2048
const MAX_SHADE_TOTAL = 32 * 1024 * 1024

interface ShadeEntry {
	readonly object: PaintedObject
	readonly drawingWidth: number
	readonly image: SVGImageElement
	readonly paint: Extract<Paint, { kind: "shaded" }>
	readonly localScale: number
	bucket: number
	key: string
}

// A page's shaded paints, rasterised at the device resolution they are drawn at (zoom × device pixel ratio,
// rounded up to a power of two) and again when the zoom crosses a 2× step; each raster is made once per
// paint and step, and every URL is revoked with the page.
class Shader {
	private readonly media: DocxMediaCache
	private readonly entries: ShadeEntry[] = []
	private readonly urls = new Map<string, { url: string; refs: number }>()
	private readonly ids = new WeakMap<object, number>()
	private nextId = 0
	private pixelsLeft = MAX_SHADE_TOTAL
	private observer: MutationObserver | null = null
	private resize: (() => void) | null = null
	private scheduled = false
	private disposed = false

	constructor(media: DocxMediaCache) {
		this.media = media
	}

	register(
		object: PaintedObject,
		drawingWidth: number,
		image: SVGImageElement,
		paint: Extract<Paint, { kind: "shaded" }>,
		localScale: number
	): void {
		if (this.disposed) return

		const entry: ShadeEntry = { object, drawingWidth, image, paint, localScale, bucket: 0, key: "" }

		this.entries.push(entry)
		this.update(entry)
		this.watch(object.element)
	}

	// Device pixels per local unit of the entry's paint as drawn now.
	private deviceScale(entry: ShadeEntry): number {
		const element = entry.object.element
		const view = element.ownerDocument.defaultView
		const width = element.getBoundingClientRect().width || entry.object.width
		const ratio = view !== null && view.devicePixelRatio > 0 ? view.devicePixelRatio : 1

		return entry.drawingWidth > 0 ? (width / entry.drawingWidth) * ratio * entry.localScale : ratio
	}

	private update(entry: ShadeEntry): void {
		const scale = this.deviceScale(entry)
		const bucket = 2 ** Math.ceil(Math.log2(Number.isFinite(scale) && scale > 0 ? Math.min(Math.max(scale, 1 / 64), 256) : 1))

		if (bucket === entry.bucket) return

		let id = this.ids.get(entry.paint)

		if (id === undefined) {
			id = ++this.nextId
			this.ids.set(entry.paint, id)
		}

		const key = `${id}@${bucket}`
		let held = this.urls.get(key)

		if (held === undefined) {
			const paint = entry.paint
			const raster =
				this.pixelsLeft > 0
					? shadeRaster(paint.points, paint.colors, paint.ramp, bucket, Math.min(MAX_SHADE_PIXELS, this.pixelsLeft))
					: undefined

			if (raster === undefined) return

			this.pixelsLeft -= raster.width * raster.height

			const url = this.media.createBytesUrl(encodePng(raster.rgba, raster.width, raster.height), "image/png")

			if (url === undefined) return

			held = { url, refs: 0 }
			this.urls.set(key, held)
		}

		held.refs++
		entry.image.setAttribute("href", held.url)
		this.drop(entry.key)
		entry.bucket = bucket
		entry.key = key
	}

	private drop(key: string): void {
		const held = this.urls.get(key)

		if (held === undefined || --held.refs > 0) return

		this.urls.delete(key)
		this.media.revokeBytesUrl(held.url)
	}

	// Zoom scales the page with a CSS transform, and the device pixel ratio changes with the window: either
	// re-checks the entries once per frame.
	private watch(element: HTMLElement): void {
		if (this.resize !== null) return

		const view = element.ownerDocument.defaultView
		const schedule = (): void => {
			if (this.scheduled || this.disposed) return

			this.scheduled = true

			const run = (): void => {
				this.scheduled = false

				if (!this.disposed) for (const entry of this.entries) this.update(entry)
			}

			if (view !== null && typeof view.requestAnimationFrame === "function") view.requestAnimationFrame(run)
			else void Promise.resolve().then(run)
		}
		const scaled = element.closest(".fo-docx-page-scale")

		this.resize = schedule
		view?.addEventListener("resize", schedule)

		if (scaled !== null && typeof MutationObserver === "function") {
			this.observer = new MutationObserver(schedule)
			this.observer.observe(scaled, { attributes: true, attributeFilter: ["style"] })
		}
	}

	dispose(): void {
		if (this.disposed) return

		this.disposed = true
		this.observer?.disconnect()

		const view = this.entries[0]?.object.element.ownerDocument.defaultView

		if (this.resize !== null) view?.removeEventListener("resize", this.resize)

		for (const held of this.urls.values()) this.media.revokeBytesUrl(held.url)

		this.urls.clear()
		this.entries.length = 0
	}
}

interface Holder {
	// Records an acquired image; false (and released at once) when the page already went away.
	hold(key: { partName: string; type: string }): boolean
	// Records an object URL made for bytes the drawing carries, revoked with the page's images.
	holdUrl(url: string): boolean
	readonly shader: Shader
}

// Raster formats a drawing may carry inline (bitmaps of a metafile); anything else is not shown.
function inlineMime(image: DrawingImage): string | undefined {
	if (!(image.bytes instanceof Uint8Array) || image.partName !== undefined) return undefined

	switch (image.format) {
		case "png":
		case "jpeg":
		case "gif":
		case "bmp":
			return imageMime(image)
		default:
			return undefined
	}
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

	// Text the painter wraps is measured: wait for the faces the engine registered for it.
	if (Array.isArray(data.drawing.texts) && data.drawing.texts.some(text => text?.kind === "block"))
		await options.media.fontsReady().catch(() => undefined)

	if (options.signal.aborted) return

	const inline = new Map<DrawingImage, string>()
	// One URL per encoded bitmap: entries that differ only in crop or effects share their bytes.
	const byBytes = new Map<Uint8Array, string>()

	for (const image of Array.isArray(data.drawing.images) ? data.drawing.images : []) {
		const type = image === null || typeof image !== "object" ? undefined : inlineMime(image)

		if (type === undefined || image.bytes === undefined || inline.has(image)) continue

		let url = byBytes.get(image.bytes)

		if (url === undefined) {
			const created = options.media.createBytesUrl(image.bytes, type)

			if (created === undefined || !holder.holdUrl(created)) continue

			url = created
			byBytes.set(image.bytes, url)
		}

		inline.set(image, url)
	}

	const svg = paintDrawing(data.drawing, {
		document: options.document,
		idPrefix: `fod${++drawingCounter}-`,
		measureText: textMeasurer(options.document),
		shade(image, paint, localScale) {
			holder.shader.register(object, data.drawing.width, image, paint, localScale)
		},
		imageUrl(image) {
			const own = inline.get(image)

			if (own !== undefined) return own

			const key = mediaKey(image)

			return key === undefined ? undefined : urls.get(`${key.type}\u0000${key.partName}`)
		}
	})

	object.element.appendChild(svg)
}

export function paintObjects(objects: readonly PaintedObject[], options: ObjectPaintOptions): () => void {
	const held: { partName: string; type: string }[] = []
	const heldUrls: string[] = []
	const shader = new Shader(options.media)
	let released = false
	const release = (): void => {
		if (released) return

		released = true
		shader.dispose()

		for (const key of held) options.media.release(key.partName, key.type)
		for (const url of heldUrls) options.media.revokeBytesUrl(url)

		held.length = 0
		heldUrls.length = 0
	}
	const holder: Holder = {
		shader,
		hold(key) {
			if (released) {
				options.media.release(key.partName, key.type)

				return false
			}

			held.push(key)

			return true
		},
		holdUrl(url) {
			if (released) {
				options.media.revokeBytesUrl(url)

				return false
			}

			heldUrls.push(url)

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
							case "checkbox":
								object.element.classList.add("fo-docx-checkbox")

								if (result.checked) object.element.dataset["foChecked"] = ""

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
								// A picture with alternative text is an image to assistive technology, named by it (the
								// drawing's own SVG stays hidden from it).
								if (typeof result.alt === "string" && result.alt.length > 0) {
									object.element.setAttribute("role", "img")
									object.element.setAttribute("aria-label", result.alt)
								}

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
