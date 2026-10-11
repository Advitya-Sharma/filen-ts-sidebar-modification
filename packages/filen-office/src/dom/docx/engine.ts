import {
	DocxDrawings,
	Item,
	Layer,
	OBJECT_INTS,
	ObjectKind,
	PropertyResolver,
	PropsStore,
	RevisionView,
	SOURCE_INTS,
	SourceInt,
	buildDocxModel,
	layoutDocument,
	openDocxPackage,
	gateReasons,
	parseNumbering,
	parseStyles,
	parseTheme,
	styleSettings,
	type DocumentLayout,
	type DocxModel,
	type DocxPackage,
	type ObjectSizer,
	type PageDisplayList
} from "../../core/docx/index"
import type { DrawingList } from "../../core/drawingml/index"
import {
	FontManager,
	cssFont,
	declareDocxEmbeddedFonts,
	readDocxFontTable,
	type FontAssets,
	type FontFaceDescriptor,
	type FontHost
} from "../../core/fonts/index"
import { OpcError, ZipError, type ByteSource, type ContentTypes, type OpcPackage } from "../../core/opc/index"
import {
	FRAGMENT_FLOATS,
	FRAGMENT_INTS,
	FRAGMENT_KIND_MASK,
	FragmentKind,
	LINE_FLOATS,
	LINE_INTS,
	LineInt,
	SMALL_CAPS_SCALE,
	type TextMeasurer
} from "../../core/text/index"
import { XmlError, type XmlDocument } from "../../core/xml/index"
import { findMatches, type FindablePage } from "./find"
import { bookmarkIndex, paragraphLinks, paragraphOf, type LinkRange } from "./links"
import {
	LINK_INTS,
	type DocxErrorReason,
	type DocxEvent,
	type DocxFindOptions,
	type DocxFontFace,
	type DocxLinkTarget,
	type DocxLocation,
	type DocxMedia,
	type DocxObjectData,
	type DocxOpenOptions,
	type DocxPageData,
	type DocxSessionApi,
	type DocxTextStyle,
	type DrawingData
} from "./protocol"

// Runs one document: package → model → styles/numbering/theme → fonts → incremental layout, and hands the
// renderer compacted pages as they are laid out. The same class runs in a worker (web) and time-sliced on
// a UI thread (an Expo DOM component): the host supplies fonts, measurement and how to yield.

export interface DocxEngineServices {
	// Registers faces in the realm that measures. `replay`: also send each face to the renderer (a worker
	// measures in its own realm; an in-thread engine shares the renderer's).
	readonly fontHost: FontHost
	readonly replayFonts: boolean
	readonly fontAssets: FontAssets
	readonly measurer: TextMeasurer
	// Gives the host a turn between layout slices (a macrotask, so requests are answered meanwhile).
	readonly yield: () => Promise<void>
	readonly sliceMs?: number
	readonly now?: () => number
}

export type DocxEmit = (event: DocxEvent, transfer: readonly ArrayBuffer[]) => void

// Images larger than this (inflated) are not shown.
const MAX_MEDIA_BYTES = 64 * 1024 * 1024
const COLOR = /^[0-9A-Fa-f]{6}$/
const IMAGE_TYPES = new Set(["image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp", "image/svg+xml"])

function isDocumentError(error: unknown): boolean {
	return error instanceof XmlError || error instanceof OpcError || error instanceof ZipError
}

function messageOf(error: unknown): string {
	return error instanceof Error ? error.message : String(error)
}

async function optionalXml(pkg: DocxPackage, kind: string): Promise<XmlDocument | undefined> {
	const name = pkg.part(kind)

	if (name === undefined) return undefined

	try {
		return await pkg.xml(name)
	} catch (error) {
		if (isDocumentError(error)) return undefined

		throw error
	}
}

function bufferOf(array: ArrayBufferView): ArrayBuffer | undefined {
	const buffer = array.buffer

	return buffer instanceof ArrayBuffer ? buffer : undefined
}

function drawingData(list: DrawingList, transfer: ArrayBuffer[]): DrawingData {
	const compact = list.compact()
	const data: DrawingData = {
		width: compact.width,
		height: compact.height,
		commands: compact.commands,
		coords: compact.coords,
		itemCount: compact.itemCount,
		itemFloats: compact.itemFloats,
		itemInts: compact.itemInts,
		paints: compact.paints,
		strokes: compact.strokes,
		images: compact.images,
		texts: compact.texts,
		shadows: compact.shadows
	}

	for (const array of [data.commands, data.coords, data.itemFloats, data.itemInts]) {
		const buffer = bufferOf(array)

		if (buffer !== undefined) transfer.push(buffer)
	}

	return data
}

// Records what the layout realm registered so a renderer in another realm can register the same aliases.
function replayingHost(inner: FontHost, sink: (face: DocxFontFace) => void): FontHost {
	return {
		async addBytes(face: FontFaceDescriptor, data: Uint8Array): Promise<boolean> {
			// `data` is only lent for the call: copy before the platform's async load.
			const copy = data.slice()
			const ok = await inner.addBytes(face, data)

			if (ok) sink({ kind: "bytes", face, bytes: copy })

			return ok
		},
		async addLocal(face: FontFaceDescriptor, name: string): Promise<boolean> {
			const ok = await inner.addLocal(face, name)

			if (ok) sink({ kind: "local", face, name })

			return ok
		},
		remove(family: string): void {
			inner.remove(family)
		}
	}
}

export class DocxEngine implements DocxSessionApi {
	private readonly source: ByteSource
	private readonly services: DocxEngineServices
	private readonly emit: DocxEmit
	private readonly view: RevisionView
	private readonly signal = { aborted: false }
	private model: DocxModel | undefined = undefined
	private layout: DocumentLayout | undefined = undefined
	private fonts: FontManager | undefined = undefined
	private drawings: Promise<DocxDrawings> | undefined = undefined
	private bookmarks: Map<string, { part: number; item: number }> | undefined = undefined
	private readonly retain: boolean
	// The zip directory and content types: images are read from them after the model is gone.
	private pkg: { readonly opc: OpcPackage; readonly contentTypes: ContentTypes } | undefined = undefined
	private pages: readonly PageDisplayList[] = []
	private mainPartBytes: number | undefined = undefined
	// Objects the shipped pages show; prepared before the model is released.
	private readonly objectIds = new Set<number>()
	private objectCache: Map<number, DocxObjectData> | undefined = undefined
	private bookmarkLocations: Map<string, DocxLocation> | undefined = undefined
	private releasing: Promise<void> | undefined = undefined
	private readonly mathSizes = new Map<number, number>()
	private readonly pendingFaces: DocxFontFace[] = []
	private readonly links: DocxLinkTarget[] = []
	private readonly linkKeys = new Map<string, number>()
	private readonly paragraphLinkCache = new Map<number, readonly LinkRange[]>()
	private shippedPages = 0
	private shippedStyles = 0
	private shippedDecorations = 0
	private shippedLinks = 0

	constructor(source: ByteSource, services: DocxEngineServices, emit: DocxEmit, options: DocxOpenOptions = {}) {
		this.source = source
		this.services = services
		this.emit = emit
		this.view = options.view === "original" ? RevisionView.Original : RevisionView.Final
		this.retain = options.retain === true
	}

	get closed(): boolean {
		return this.signal.aborted
	}

	// Opens and lays out the whole document; resolves when done, failed or closed. Never throws.
	async run(): Promise<void> {
		try {
			await this.open()
		} catch (error) {
			if (this.closed) return

			this.fail(isDocumentError(error) ? "malformed" : "failed", messageOf(error))
		}
	}

	private fail(reason: DocxErrorReason, message: string): void {
		this.emit(
			{ type: "error", reason, message, ...(this.mainPartBytes === undefined ? {} : { mainPartBytes: this.mainPartBytes }) },
			[]
		)
		this.close()
	}

	private async open(): Promise<void> {
		const opened = await openDocxPackage(this.source)

		if (this.closed) return

		if (!opened.ok) {
			this.fail(opened.reason, opened.message)

			return
		}

		const pkg = opened.package

		this.pkg = { opc: pkg.opc, contentTypes: pkg.contentTypes }
		this.mainPartBytes = pkg.opc.entry(pkg.mainPart)?.uncompressedSize ?? 0

		const model = await buildDocxModel(pkg)

		if (this.closed) return

		this.model = model

		const store = new PropsStore()
		const stylesXml = await optionalXml(pkg, "styles")
		const styles = stylesXml === undefined ? undefined : parseStyles(stylesXml, store)
		const numberingXml = await optionalXml(pkg, "numbering")
		const numbering = numberingXml === undefined ? undefined : parseNumbering(numberingXml, store, styles)
		const themeXml = await optionalXml(pkg, "theme")
		const theme = themeXml === undefined ? undefined : parseTheme(themeXml)
		const fontTablePart = pkg.part("fontTable")
		const fontTable = await optionalXml(pkg, "fontTable")

		if (this.closed) return

		const services = this.services
		const host = services.replayFonts ? replayingHost(services.fontHost, face => this.pendingFaces.push(face)) : services.fontHost
		const fonts = new FontManager(host, services.fontAssets)

		this.fonts = fonts

		if (fontTable !== undefined && fontTablePart !== undefined) {
			const declarations = readDocxFontTable(fontTable)

			fonts.declare(declarations)
			declareDocxEmbeddedFonts(fonts, pkg.opc, fontTablePart, declarations)
		}

		const resolver = new PropertyResolver({ styles, numbering, theme, settings: styleSettings(model.settings) })
		const body = model.body === undefined ? undefined : model.partOf(model.body)
		let breaks = 0

		if (body !== undefined) for (const kind of body.kinds) if (kind === Item.LastRenderedPageBreak) breaks++

		const firstSection = model.sections[0]
		const section =
			body === undefined || firstSection === undefined || firstSection.sectPr < 0
				? undefined
				: resolver.section(store.section(body.document, firstSection.sectPr))
		// Twips to px (1440 per inch, 96 px per inch); Word's Letter default without a section.
		const pageWidth = section === undefined ? 816 : section.pageWidth / 15
		const pageHeight = section === undefined ? 1056 : section.pageHeight / 15

		this.emit(
			{
				type: "opened",
				info: {
					variant: pkg.facts.variant,
					gate: gateReasons(pkg.facts, model, styles, numbering),
					estimatedPages: breaks + 1,
					pageWidth: pageWidth > 0 && Number.isFinite(pageWidth) ? pageWidth : 816,
					pageHeight: pageHeight > 0 && Number.isFinite(pageHeight) ? pageHeight : 1056,
					mainPartBytes: this.mainPartBytes
				}
			},
			[]
		)

		const mathSizes = this.mathSizes
		const sizer: ObjectSizer = {
			size(object) {
				// Only the equation's run size is recorded (the renderer sizes MathML with it).
				if (object.kind === ObjectKind.MATH) mathSizes.set(object.part.index * 0x4000000 + object.item, object.fontSize)

				return undefined
			}
		}
		const layout = layoutDocument(model, resolver, fonts, services.measurer, {
			store,
			sizer,
			view: this.view,
			signal: this.signal,
			sliceMs: services.sliceMs ?? 12,
			yield: services.yield,
			...(services.now === undefined ? {} : { now: services.now })
		})

		this.layout = layout
		this.pages = layout.pages

		const unsubscribe = layout.subscribe(event => {
			// The final batch goes out once, below, with the pages laid out again for page count fields.
			if (!this.closed && !event.final) this.ship(false)
		})

		try {
			await layout.done
		} finally {
			unsubscribe()
		}

		if (this.closed || layout.cancelled) return

		// Pages laid out again for page count fields replace what was sent.
		const dirty = [...layout.dirty].filter(index => index < this.shippedPages)

		this.ship(true, dirty)

		if (!this.retain) {
			this.releasing = this.release()
			await this.releasing
		}
	}

	private ship(final: boolean, replace: readonly number[] = []): void {
		const layout = this.layout

		if (layout === undefined) return

		const pages: DocxPageData[] = []
		const transfer: ArrayBuffer[] = []

		for (const index of replace) {
			const page = layout.pages[index]

			if (page !== undefined) pages.push(this.pageData(page, transfer))
		}

		for (; this.shippedPages < layout.pages.length; this.shippedPages++) {
			const page = layout.pages[this.shippedPages]

			if (page !== undefined) pages.push(this.pageData(page, transfer))
		}

		if (pages.length === 0 && !final) return

		this.flushFaces()

		const styles = this.newStyles(layout)
		const decorations = layout.decorations.slice(this.shippedDecorations)

		this.shippedDecorations = layout.decorations.length

		const links = this.links.slice(this.shippedLinks)

		this.shippedLinks = this.links.length
		this.emit(
			{ type: "pages", pages, styles, decorations, links, final: final && layout.final, pageCount: layout.pages.length },
			transfer
		)
	}

	private flushFaces(): void {
		if (this.pendingFaces.length === 0) return

		const faces = this.pendingFaces.splice(0)
		const transfer: ArrayBuffer[] = []

		for (const face of faces) {
			if (face.kind === "bytes") {
				const buffer = bufferOf(face.bytes)

				if (buffer !== undefined) transfer.push(buffer)
			}
		}

		this.emit({ type: "fonts", faces }, transfer)
	}

	private newStyles(layout: DocumentLayout): DocxTextStyle[] {
		const table = layout.styles
		const info = layout.styleInfo
		const measurer = this.services.measurer
		const out: DocxTextStyle[] = []

		for (let id = this.shippedStyles; id < table.count; id++) {
			const style = table.get(id)
			const font = cssFont(style.font, style.size)
			const extents = measurer.fontExtents(font)
			const color = info[id]?.color

			out.push({
				font,
				smallCapsFont: cssFont(style.font, style.size * SMALL_CAPS_SCALE),
				color: color !== undefined && COLOR.test(color) ? color : "",
				kerning: style.kerning ?? true,
				letterSpacing: table.letterSpacing[id] ?? 0,
				scaleX: table.scaleX[id] ?? 1,
				shift: table.shift[id] ?? 0,
				ascent: extents.ascent,
				descent: extents.descent,
				revision: info[id]?.revision ?? 0
			})
		}

		this.shippedStyles = table.count

		return out
	}

	private linkIndex(target: DocxLinkTarget): number {
		const key =
			target.kind === "external" ? `e${target.href}\u0000${target.tooltip ?? ""}` : `i${target.anchor}\u0000${target.tooltip ?? ""}`
		let index = this.linkKeys.get(key)

		if (index === undefined) {
			index = this.links.length
			this.links.push(target)
			this.linkKeys.set(key, index)
		}

		return index
	}

	private sourceLinks(story: number, item: number, text: string): readonly LinkRange[] {
		const model = this.model

		if (model === undefined || story < 0 || item < 0) return []

		const key = story * 0x4000000 + item
		let ranges = this.paragraphLinkCache.get(key)

		if (ranges === undefined) {
			const storyInfo = model.stories[story]
			const part = storyInfo === undefined ? undefined : model.parts[storyInfo.part]

			ranges = part === undefined || part.kinds[item] !== Item.Paragraph ? [] : paragraphLinks(part, item, text, this.view)
			this.paragraphLinkCache.set(key, ranges)
		}

		return ranges
	}

	// The page's placed objects plus its inline ones, which the text layout carries as object fragments
	// on their lines (box: the fragment's x and width, the object's ascent above the baseline).
	private inlineObjects(page: PageDisplayList): { count: number; floats: Float32Array; ints: Int32Array; kinds: Uint8Array } {
		const table = this.layout?.objects
		const text = page.text
		const floats = Array.from(page.objectFloats)
		const ints = Array.from(page.objectInts)

		for (let line = 0; line < text.lineCount; line++) {
			const baseline = text.lineFloats[line * LINE_FLOATS + 2] ?? 0
			const first = text.lineInts[line * LINE_INTS + LineInt.FIRST_FRAGMENT] ?? 0
			const count = text.lineInts[line * LINE_INTS + LineInt.FRAGMENT_COUNT] ?? 0

			for (let k = first; k < first + count; k++) {
				if (((text.fragmentInts[k * FRAGMENT_INTS + 3] ?? 0) & FRAGMENT_KIND_MASK) !== FragmentKind.OBJECT) continue

				const id = text.fragmentInts[k * FRAGMENT_INTS + 4] ?? -1
				const height = table?.heights[id] ?? 0

				if (id < 0 || !(height > 0)) continue

				floats.push(
					text.fragmentFloats[k * FRAGMENT_FLOATS] ?? 0,
					baseline - (table?.ascents[id] ?? height),
					text.fragmentFloats[k * FRAGMENT_FLOATS + 1] ?? table?.widths[id] ?? 0,
					height
				)
				ints.push(id, Layer.TEXT, text.lineInts[line * LINE_INTS + LineInt.SOURCE] ?? -1)
			}
		}

		const total = ints.length / OBJECT_INTS
		const kinds = new Uint8Array(total)

		for (let o = 0; o < total; o++) this.objectIds.add(ints[o * OBJECT_INTS] ?? -1)

		for (let o = 0; o < total; o++) kinds[o] = table?.kinds[ints[o * OBJECT_INTS] ?? -1] ?? 0

		return { count: total, floats: Float32Array.from(floats), ints: Int32Array.from(ints), kinds }
	}

	private pageData(page: PageDisplayList, transfer: ArrayBuffer[]): DocxPageData {
		const objects = this.inlineObjects(page)

		const linkInts: number[] = []

		for (let source = 0; source < page.sourceCount; source++) {
			const at = source * SOURCE_INTS
			const ranges = this.sourceLinks(
				page.sourceInts[at + SourceInt.STORY] ?? -1,
				page.sourceInts[at + SourceInt.ITEM] ?? -1,
				page.sourceTexts[source] ?? ""
			)

			for (const range of ranges) linkInts.push(source, range.start, range.end, this.linkIndex(range.target))
		}

		const text = page.text
		// Copies: the layout keeps its pages (anchors, page count re-layout); the copies are transferred.
		const data: DocxPageData = {
			index: page.index,
			number: page.number,
			numberText: page.numberText,
			section: page.section,
			width: page.width,
			height: page.height,
			lineCount: text.lineCount,
			fragmentCount: text.fragmentCount,
			decorationCount: text.decorationCount,
			lineFloats: text.lineFloats.slice(),
			lineInts: text.lineInts.slice(),
			fragmentFloats: text.fragmentFloats.slice(),
			fragmentInts: text.fragmentInts.slice(),
			decorationFloats: text.decorationFloats.slice(),
			decorationInts: text.decorationInts.slice(),
			graphicCount: page.graphicCount,
			graphicFloats: page.graphicFloats.slice(),
			graphicInts: page.graphicInts.slice(),
			objectCount: objects.count,
			objectFloats: objects.floats,
			objectInts: objects.ints,
			objectKinds: objects.kinds,
			sourceCount: page.sourceCount,
			sourceInts: page.sourceInts.slice(),
			sourceTexts: page.sourceTexts,
			linkCount: linkInts.length / LINK_INTS,
			linkInts: Int32Array.from(linkInts)
		}

		for (const array of [
			data.lineFloats,
			data.lineInts,
			data.fragmentFloats,
			data.fragmentInts,
			data.decorationFloats,
			data.decorationInts,
			data.graphicFloats,
			data.graphicInts,
			data.objectFloats,
			data.objectInts,
			data.objectKinds,
			data.sourceInts,
			data.linkInts
		]) {
			const buffer = bufferOf(array)

			if (buffer !== undefined && buffer.byteLength > 0) transfer.push(buffer)
		}

		return data
	}

	private drawingLayer(model: DocxModel): Promise<DocxDrawings> {
		this.drawings ??= DocxDrawings.open(model)

		return this.drawings
	}

	// One object as the renderer draws it, from the live model.
	private async objectData(id: number, transfer: ArrayBuffer[]): Promise<DocxObjectData> {
		const model = this.model
		const table = this.layout?.objects

		if (model === undefined || table === undefined) return { kind: "placeholder", id, reason: "closed" }

		const kind = table.kinds[id]
		const index = table.indices[id] ?? -1

		if (kind === ObjectKind.SEPARATOR || kind === ObjectKind.CONTINUATION_SEPARATOR) return { kind: "separator", id }

		if (kind === ObjectKind.MATH) {
			const part = table.parts[id] ?? -1
			const item = table.items[id] ?? -1
			const node = model.parts[part]?.node(item) ?? -1
			const tree = node < 0 ? undefined : (await this.drawingLayer(model)).math(part, node)

			return tree === undefined
				? { kind: "placeholder", id, reason: "math" }
				: { kind: "math", id, tree, fontSize: this.mathSizes.get(part * 0x4000000 + item) ?? 14.667 }
		}

		if (kind === ObjectKind.DRAWING || kind === ObjectKind.VML) {
			const drawings = await this.drawingLayer(model)
			const drawingInfo = kind === ObjectKind.DRAWING ? model.drawings[index] : undefined
			const vmlInfo = kind === ObjectKind.VML ? model.vml[index] : undefined
			const result =
				drawingInfo !== undefined
					? await drawings.drawing(drawingInfo)
					: vmlInfo !== undefined
						? drawings.vml(vmlInfo)
						: ({ kind: "placeholder", reason: "missing" } as const)

			return result.kind === "drawing"
				? { kind: "drawing", id, drawing: drawingData(result.list, transfer) }
				: { kind: "placeholder", id, reason: result.reason }
		}

		return { kind: "placeholder", id, reason: "unsupported" }
	}

	async objects(ids: readonly number[], transfer: ArrayBuffer[] = []): Promise<readonly DocxObjectData[]> {
		if (this.releasing !== undefined) await this.releasing

		if (this.closed) return []

		const cache = this.objectCache
		const out: DocxObjectData[] = []

		for (const id of ids) {
			// Kept objects are cloned by the message, never transferred: the cache still needs them.
			const kept = cache?.get(id)

			out.push(kept ?? (cache === undefined ? await this.objectData(id, transfer) : { kind: "placeholder", id, reason: "missing" }))
		}

		return out
	}

	async media(partName: string): Promise<DocxMedia | null> {
		const pkg = this.pkg

		if (pkg === undefined || this.closed) return null

		const entry = pkg.opc.entry(partName)

		if (entry === undefined || entry.uncompressedSize > MAX_MEDIA_BYTES) return null

		try {
			const bytes = await pkg.opc.read(partName)
			const type = (pkg.contentTypes.typeOf(partName) ?? "").toLowerCase()

			return { bytes, type: IMAGE_TYPES.has(type) ? type : "" }
		} catch (error) {
			if (isDocumentError(error)) return null

			throw error
		}
	}

	async find(query: string, options: DocxFindOptions = {}): Promise<Int32Array> {
		if (this.closed) return new Int32Array(0)

		return Promise.resolve(findMatches(this.findablePages(), query, options))
	}

	private *findablePages(): Iterable<FindablePage> {
		for (const page of this.pages) {
			yield {
				index: page.index,
				sourceCount: page.sourceCount,
				sourceTexts: page.sourceTexts,
				lineCount: page.text.lineCount,
				lineInts: page.text.lineInts
			}
		}
	}

	// Where every laid-out paragraph starts: "part:paragraph" -> its first line.
	private paragraphStarts(model: DocxModel): Map<string, DocxLocation> {
		const out = new Map<string, DocxLocation>()

		for (const page of this.pages) {
			const text = page.text

			for (let line = 0; line < text.lineCount; line++) {
				const source = text.lineInts[line * LINE_INTS + LineInt.SOURCE] ?? -1
				const at = source * SOURCE_INTS
				const story = model.stories[page.sourceInts[at + SourceInt.STORY] ?? -1]
				const key = `${story?.part ?? -1}:${page.sourceInts[at + SourceInt.ITEM] ?? -1}`

				if (story !== undefined && !out.has(key)) out.set(key, { page: page.index, y: text.lineFloats[line * LINE_FLOATS] ?? 0 })
			}
		}

		return out
	}

	private bookmarkLocation(model: DocxModel, anchor: string, starts: Map<string, DocxLocation>): DocxLocation | null {
		this.bookmarks ??= bookmarkIndex(model)

		const target = this.bookmarks.get(anchor)
		const part = target === undefined ? undefined : model.parts[target.part]

		if (target === undefined || part === undefined) return null

		return starts.get(`${target.part}:${paragraphOf(part, target.item)}`) ?? null
	}

	async locate(anchor: string): Promise<DocxLocation | null> {
		if (anchor === "_top") return this.closed ? null : { page: 0, y: 0 }

		if (this.releasing !== undefined) await this.releasing

		const kept = this.bookmarkLocations

		if (kept !== undefined) return this.closed ? null : (kept.get(anchor) ?? null)

		const model = this.model
		const layout = this.layout

		if (model === undefined || layout === undefined || this.closed) return null

		const now = this.bookmarkLocation(model, anchor, this.paragraphStarts(model))

		if (now !== null || layout.final) return now

		await layout.done.catch(() => undefined)

		if (this.releasing !== undefined) await this.releasing

		if (this.bookmarkLocations !== undefined) return this.closed ? null : (this.bookmarkLocations.get(anchor) ?? null)

		return this.closed || this.model === undefined ? null : this.bookmarkLocation(this.model, anchor, this.paragraphStarts(this.model))
	}

	// After the last page: prepares every object and bookmark the pages can ask for, then drops the model,
	// its XML and the worker's font faces. What remains is the pages (find), the zip directory (images) and
	// those prepared answers.
	private async release(): Promise<void> {
		const model = this.model

		if (model === undefined) return

		const cache = new Map<number, DocxObjectData>()
		const services = this.services
		let slice = (services.now ?? Date.now)()

		for (const id of this.objectIds) {
			cache.set(id, await this.objectData(id, []))

			if (this.closed) return

			const now = (services.now ?? Date.now)()

			if (now - slice > (services.sliceMs ?? 12)) {
				await services.yield()

				if (this.closed) return

				slice = (services.now ?? Date.now)()
			}
		}

		const starts = this.paragraphStarts(model)
		const bookmarks = new Map<string, DocxLocation>()

		this.bookmarks ??= bookmarkIndex(model)

		for (const name of this.bookmarks.keys()) {
			const location = this.bookmarkLocation(model, name, starts)

			if (location !== null) bookmarks.set(name, location)
		}

		this.objectCache = cache
		this.bookmarkLocations = bookmarks

		// A worker's faces only served measuring; an in-thread engine's are the page's own, drawn until close.
		if (this.services.replayFonts) {
			this.fonts?.dispose()
			this.fonts = undefined
		}

		this.layout = undefined
		this.model = undefined
		this.drawings = undefined
		this.bookmarks = undefined
		this.paragraphLinkCache.clear()
		this.mathSizes.clear()
		this.objectIds.clear()
	}

	close(): void {
		if (this.signal.aborted) return

		this.signal.aborted = true
		this.layout?.cancel()
		this.fonts?.dispose()
		this.fonts = undefined
		this.layout = undefined
		this.model = undefined
		this.pkg = undefined
		this.pages = []
		this.drawings = undefined
		this.bookmarks = undefined
		this.objectCache = undefined
		this.bookmarkLocations = undefined
		this.paragraphLinkCache.clear()
		this.mathSizes.clear()
		this.objectIds.clear()
		this.pendingFaces.length = 0
	}
}
