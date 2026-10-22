import {
	CellFlag,
	tileKey,
	type SheetId,
	type SheetLayout,
	type SheetRange,
	type SheetRenderStyle,
	type SheetTile
} from "../../../core/sheet/protocol"
import { parseHeaderFooter, type HeaderFooterValues } from "../../../core/sheet/print/headerFooter"
import type { PrintPage } from "../../../core/sheet/print/paginate"
import type { SheetPageSetup } from "../../../core/sheet/print/setup"
import type { SheetHeaderImage, SheetPrintNote } from "../../../core/sheet/print/xlsx"
import { columnName } from "../../../core/sheet/refs/cellRef"
import type { SheetDrawingObject } from "../../../core/xlsx/drawings/feature"
import { mediaContentType } from "../../../core/xlsx/drawings/mediaTypes"
import type { SheetDrawingData } from "../../../core/xlsx/drawings/render"
import { DocxMediaCache } from "../../docx/media"
import { drawingImages, imageMime, inlineImageUrls, paintDrawing } from "../../docx/paint/drawing"
import { anchorBox, type SheetBox } from "../drawings/geometry"
import { SheetGridModel } from "../model"
import type { SheetGridTheme } from "../theme"
import { framePage, pageTiles, pageView, printPageSvg, printSheet, type BandPictures, type PrintedSheet } from "./printPage"
import { noteBox, printRanges, wholeAxes } from "./printRanges"
import { escapeSvgText } from "./svgCanvas"

// A print as a list of pages over the grid model's session: each sheet cut into pages (and the pages listing its
// notes, when it prints them at the end), then each page drawn as SVG once the tiles it shows are held, with the
// sheet's objects (pictures, shapes, charts) and shown notes placed by their anchors and clipped at the page's edge,
// and its header and footer pictures.
//
// Tiles a page needs and the grid does not hold are asked of the session and put in the grid model's own cache, as if
// the grid had fetched them: the session holds every tile it answered until its client releases it, and the cache's
// eviction releases them like any other (a release of our own could race a fetch of the grid's for the same tile).
// Styles the answer brings reach the model the same way. A page itself is drawn from a model of its own holding the
// tiles as printed (errors as the sheet prints them, black and white), which never releases anything.

export interface PrintSheetRequest {
	readonly sheet: SheetId
	readonly name: string
	readonly setup: SheetPageSetup
	// What prints: the selection's ranges, else the print area's, else the used range (null: nothing).
	readonly ranges: readonly SheetRange[]
	readonly used: SheetRange | null
}

// A page listing notes: each note's cell and its text, in lines as they fit the page.
export interface NotesPage {
	readonly lines: readonly { readonly text: string; readonly bold: boolean }[]
}

export interface PrintJobSheet extends PrintedSheet {
	readonly notesAtEnd: readonly NotesPage[]
	// Notes shown all the time, drawn where they stand (cellComments asDisplayed).
	readonly shownNotes: readonly SheetPrintNote[]
	readonly objects: readonly SheetDrawingObject[]
	readonly headerImages: readonly SheetHeaderImage[]
}

export interface PrintJob {
	readonly model: SheetGridModel
	readonly sheets: readonly PrintJobSheet[]
	// The print was laid out up to its page cap and more pages would follow.
	readonly more: boolean
	readonly measure: (font: string, text: string) => number
	// Pictures as object URLs for the print's lifetime.
	readonly media: DocxMediaCache | null
	// Drawn objects, once per object.
	readonly drawn: Map<string, Promise<string | null>>
	dispose(): void
}

const NOTE_FONT_PX = 12
const NOTE_FONT = `${String(NOTE_FONT_PX)}px Tahoma, Arial, sans-serif`
const NOTE_LINE_PX = 15
const LIST_FONT_PT = 10
const LIST_LINE_PT = 14

// `text` broken into lines `width` wide at spaces (a long word between characters); line breaks kept.
function wrap(text: string, width: number, measure: (text: string) => number): string[] {
	const out: string[] = []

	for (const paragraph of text.split(/\r?\n|\r/u)) {
		let line = ""

		for (const word of paragraph.match(/\S+\s*|\s+/gu) ?? [""]) {
			if (line !== "" && measure(line + word.trimEnd()) > width) {
				out.push(line.trimEnd())
				line = ""
			}

			if (measure(word.trimEnd()) <= width || line !== "") {
				line += word

				continue
			}

			for (const char of word) {
				if (line !== "" && measure(line + char) > width) {
					out.push(line)
					line = ""
				}

				line += char
			}
		}

		out.push(line.trimEnd())
	}

	return out
}

// The notes listed after a sheet, as Excel prints them: a "Cell:" line and a "Note:" line each, a blank line between.
function notesPages(notes: readonly SheetPrintNote[], printed: PrintedSheet, measure: (font: string, text: string) => number): NotesPage[] {
	if (notes.length === 0) return []

	const { area } = printed.pagination
	const font = `${String(LIST_FONT_PT)}pt Calibri, Carlito, Arial, sans-serif`
	const perPage = Math.max(1, Math.floor(area.height / LIST_LINE_PT))
	const lines: { text: string; bold: boolean }[] = []

	for (const note of notes) {
		if (lines.length > 0) lines.push({ text: "", bold: false })

		lines.push({ text: `Cell: ${columnName(note.col)}${String(note.row + 1)}`, bold: true })

		const body = note.author === "" || note.text.startsWith(note.author) ? note.text : `${note.author}: ${note.text}`

		// Measured in px: the list's points are 4/3 as many px.
		for (const text of wrap(`Note: ${body}`, (area.width * 4) / 3, piece => measure(font, piece))) lines.push({ text, bold: false })
	}

	const pages: NotesPage[] = []

	for (let k = 0; k < lines.length; k += perPage) pages.push({ lines: lines.slice(k, k + perPage) })

	return pages
}

// Whether any header or footer text of the sheet shows a picture (&G).
function picturesUsed(setup: SheetPageSetup): boolean {
	const { odd, even, first } = setup.headerFooter

	return [odd.header, odd.footer, even.header, even.footer, first.header, first.footer].some(text =>
		Object.values(parseHeaderFooter(text)).some((runs: readonly { kind: string; field?: string }[]) =>
			runs.some(run => run.kind === "field" && run.field === "picture")
		)
	)
}

// `maxPages`: the sheets' pages laid out at most, all sheets together (what a browser can print), so a huge range
// is not cut into pages nothing shows; the job says when more would follow.
export async function layOutPrint(
	model: SheetGridModel,
	sheets: readonly PrintSheetRequest[],
	measure: (font: string, text: string) => number = (_font, text) => text.length * 7,
	maxPages?: number
): Promise<PrintJob> {
	const out: PrintJobSheet[] = []
	let left = maxPages ?? Number.POSITIVE_INFINITY
	let more = false
	const xlsx = model.kind !== "csv" && model.kind !== "tsv" && model.kind !== "tab"
	let media: DocxMediaCache | null = null

	for (const request of sheets) {
		// The session's own layout: it carries what features read beside the sheet (drawings) as they now are.
		const layout = await model.session.layout(request.sheet)
		const { setup } = request
		// Draft quality prints no graphics.
		const objects = setup.draft ? [] : (layout.extensions?.drawings?.objects.filter(object => !object.hidden) ?? [])
		const notes = xlsx && setup.cellComments !== "none" ? await model.session.query("print.notes", { sheet: request.sheet }) : []
		const shownNotes = setup.cellComments === "asDisplayed" ? notes.filter(note => note.shown) : []
		const ranges = printRanges(request.ranges, request.used, objects, shownNotes, layout)
		const headerImages =
			xlsx && !setup.draft && picturesUsed(setup) ? await model.session.query("print.headerImages", { sheet: request.sheet }) : []
		const notesAtEnd = setup.cellComments === "atEnd" ? notes : []

		if (ranges.length === 0 && notesAtEnd.length === 0) continue

		if (left <= 0) {
			more = true

			break
		}

		const exact = xlsx ? await model.session.query("print.rows", { sheet: request.sheet }) : null
		const printed = printSheet(request.sheet, request.name, layout, setup, ranges, Number.isFinite(left) ? left : undefined, exact)

		left -= printed.pagination.pages.length
		more ||= printed.pagination.more

		if (objects.length > 0 || headerImages.length > 0)
			media ??= new DocxMediaCache({ media: partName => model.session.query("drawing.media", { partName }) })

		out.push({
			...printed,
			notesAtEnd: notesPages(notesAtEnd, printed, measure),
			shownNotes,
			objects,
			headerImages
		})
	}

	const drawn = new Map<string, Promise<string | null>>()

	return {
		model,
		sheets: out,
		more,
		measure,
		media,
		drawn,
		dispose: () => {
			media?.dispose()
			drawn.clear()
		}
	}
}

function sheetPages(sheet: PrintJobSheet): number {
	return sheet.pagination.pages.length + sheet.notesAtEnd.length
}

export function pageCount(job: PrintJob): number {
	return job.sheets.reduce((total, sheet) => total + sheetPages(sheet), 0)
}

// The nth page of a print (0-based) with its sheet, its number (&P) and whether it is its sheet's first: a page of
// cells, or one listing notes.
export function printPageAt(
	job: PrintJob,
	index: number
): {
	readonly sheet: PrintJobSheet
	readonly page: PrintPage | null
	readonly notes: NotesPage | null
	readonly number: number
	readonly first: boolean
} | null {
	let before = 0
	let number = 0

	for (const sheet of job.sheets) {
		const pages = sheet.pagination.pages
		const count = sheetPages(sheet)
		const start = sheet.setup.firstPageNumber ?? number + 1

		if (index < before + count) {
			const at = index - before
			const page = at < pages.length ? (pages[at] ?? null) : null
			const notes = at >= pages.length ? (sheet.notesAtEnd[at - pages.length] ?? null) : null

			return { sheet, page, notes, number: start + at, first: at === 0 }
		}

		before += count
		number = start + count - 1
	}

	return null
}

// --- Tiles and the page's own model ---

async function holdTiles(model: SheetGridModel, sheet: SheetId, coordinates: readonly number[]): Promise<SheetTile[]> {
	const missing: number[] = []
	const out: SheetTile[] = []

	for (let i = 0; i + 1 < coordinates.length; i += 2) {
		const tileRow = coordinates[i] ?? 0
		const tileCol = coordinates[i + 1] ?? 0
		const held = model.tiles.peek(tileKey(sheet, tileRow, tileCol))

		if (held === undefined) missing.push(tileRow, tileCol)
		else out.push(held)
	}

	if (missing.length === 0) return out

	const answer = await model.session.tiles(sheet, Uint32Array.from(missing))

	if (answer.styles !== undefined)
		model.apply({ state: model.getSnapshot().state, patches: [], invalidated: [], layouts: [], styles: answer.styles })

	for (const tile of answer.tiles) {
		const key = tileKey(tile.sheet, tile.tileRow, tile.tileCol)

		out.push(model.tiles.peek(key) ?? tile)

		if (!model.tiles.has(key)) model.tiles.put(tile)
	}

	return out
}

const ERROR_TEXTS: Readonly<Record<SheetPageSetup["errors"], string | null>> = { displayed: null, blank: "", dash: "--", NA: "#N/A" }

// A tile with its error values as the sheet prints them (pageSetup errors, §18.18.60).
export function printedTile(tile: SheetTile, errors: SheetPageSetup["errors"]): SheetTile {
	const replacement = ERROR_TEXTS[errors]

	if (replacement === null) return tile

	let texts: Int32Array | null = null
	let strings: string[] | null = null
	let index = -1

	for (let entry = 0; entry < tile.count; entry++) {
		if (((tile.flags[entry] ?? 0) & CellFlag.ERROR) === 0) continue

		texts ??= tile.texts.slice()

		if (replacement === "") {
			texts[entry] = -1

			continue
		}

		if (strings === null) {
			strings = [...tile.strings, replacement]
			index = strings.length - 1
		}

		texts[entry] = index
	}

	return texts === null ? tile : { ...tile, texts, strings: strings ?? tile.strings }
}

// A look as the sheet prints it in black and white: no fills, automatic (black) text and lines.
export function blackAndWhite(style: SheetRenderStyle): SheetRenderStyle {
	const out: { -readonly [K in keyof SheetRenderStyle]: SheetRenderStyle[K] } = { ...style }

	delete out.fill
	delete out.pattern
	delete out.color

	for (const edge of ["top", "right", "bottom", "left", "diagonalUp", "diagonalDown"] as const) {
		const line = style[edge]

		if (line !== undefined) out[edge] = { ...line, color: 0 }
	}

	return out
}

// The highest render style index the tiles and the sheet's rows and columns use.
function highestStyle(tiles: readonly SheetTile[], layout: SheetLayout): number {
	let high = 0

	for (const tile of tiles) {
		for (const style of tile.styles) high = Math.max(high, style)
		for (const style of tile.layers?.runs?.runStyles ?? []) high = Math.max(high, style)
	}

	for (const style of layout.rows.styles) high = Math.max(high, style)
	for (const style of layout.cols.styles) high = Math.max(high, style)

	return high
}

// A model of one page: the tiles as printed and the live model's looks (in black and white when the sheet says).
function pageModel(live: SheetGridModel, sheet: PrintJobSheet, tiles: readonly SheetTile[]): SheetGridModel {
	const snapshot = live.getSnapshot()
	const count = highestStyle(tiles, sheet.layout) + 1
	const styles: SheetRenderStyle[] = []

	for (let index = 0; index < count; index++)
		styles.push(sheet.setup.blackAndWhite ? blackAndWhite(live.style(index)) : live.style(index))

	const model = new SheetGridModel({
		session: live.session,
		info: {
			kind: live.kind,
			sheets: snapshot.sheets,
			activeSheet: sheet.sheet,
			activeLayout: sheet.layout,
			styles: { from: 0, styles },
			writability: snapshot.writability,
			state: snapshot.state,
			date1904: live.date1904,
			protection: null,
			fileSharing: null
		}
	})

	model.setShowFormulas(snapshot.showFormulas)

	for (const tile of tiles) model.tiles.put(printedTile(tile, sheet.setup.errors))

	return model
}

// --- Objects, notes and header pictures ---

let painted = 0

// A display list as SVG markup stretched over `width` by `height` (px at 100 %, or points in a band), its pictures
// lent by the job.
async function drawingMarkup(job: PrintJob, drawing: SheetDrawingData, width: number, height: number): Promise<string> {
	const urls = new Map<string, string | undefined>()
	const media = job.media

	if (media !== null) {
		await Promise.all(
			drawingImages(drawing).map(async image => {
				const partName = image.svgPartName ?? image.partName
				const type = imageMime(image)

				if (partName !== undefined && type !== "") urls.set(partName, await media.acquire(partName, type))
			})
		)
	}

	// Bitmaps a metafile carries: object URLs the job's media revokes when the print is disposed.
	const inline = media === null ? null : inlineImageUrls(drawing, (bytes, type) => media.createBytesUrl(bytes, type))

	painted += 1

	const svg = paintDrawing(drawing, {
		document,
		idPrefix: `fop${String(painted)}-`,
		measureText: (text, font) => job.measure(font, text),
		imageUrl: image => inline?.get(image) ?? urls.get(image.svgPartName ?? image.partName ?? "")
	})

	svg.setAttribute("width", String(width))
	svg.setAttribute("height", String(height))
	svg.setAttribute("preserveAspectRatio", "none")

	return new XMLSerializer().serializeToString(svg)
}

// An object drawn once per print: a chart through the charts feature (chart.render, laid out at the object's size), a
// picture or shape through the drawings feature (drawing.render). Null when it does not draw.
function objectMarkup(job: PrintJob, sheet: PrintJobSheet, object: SheetDrawingObject, box: SheetBox): Promise<string | null> {
	const key = `${String(sheet.sheet)}:${String(object.key)}:${object.stamp}`
	let made = job.drawn.get(key)

	if (made === undefined) {
		made = (async () => {
			try {
				if (object.chartPart !== undefined) {
					const chart = await job.model.session.query("chart.render", {
						part: object.chartPart,
						width: Math.max(1, Math.round(box.width)) * 9525,
						height: Math.max(1, Math.round(box.height)) * 9525
					})

					return chart.kind === "drawing" ? await drawingMarkup(job, chart.drawing, box.width, box.height) : null
				}

				const rendered = await job.model.session.query("drawing.render", {
					sheet: sheet.sheet,
					key: object.key,
					stamp: object.stamp
				})

				return rendered?.kind === "drawing" ? await drawingMarkup(job, rendered.drawing, box.width, box.height) : null
			} catch {
				return null
			}
		})()
		job.drawn.set(key, made)
	}

	return made
}

function shownNoteMarkup(note: SheetPrintNote, box: SheetBox, measure: (font: string, text: string) => number): string {
	const lines = wrap(note.author === "" ? note.text : `${note.author}:\n${note.text}`, Math.max(1, box.width - 8), text =>
		measure(NOTE_FONT, text)
	)
	const fitting = lines.slice(0, Math.max(1, Math.floor((box.height - 4) / NOTE_LINE_PX)))
	const text = fitting
		.map(
			(line, k) =>
				`<text x="${String(box.x + 4)}" y="${String(box.y + 3 + NOTE_FONT_PX + k * NOTE_LINE_PX)}"${k === 0 && note.author !== "" ? ' font-weight="bold"' : ""}>${escapeSvgText(line)}</text>`
		)
		.join("")

	return (
		`<rect x="${String(box.x + 0.5)}" y="${String(box.y + 0.5)}" width="${String(box.width - 1)}" height="${String(box.height - 1)}" fill="#ffffe1" stroke="#000000"/>` +
		`<g style="font:${NOTE_FONT}" fill="#000000">${text}</g>`
	)
}

// What a page draws over its cells: objects and shown notes whose boxes reach into its own rows and columns, each
// placed by its anchor and clipped at the page's edge (an object across a page break prints in part on each page).
async function overMarkup(job: PrintJob, sheet: PrintJobSheet, page: PrintPage, id: string): Promise<string> {
	if (sheet.objects.length === 0 && sheet.shownNotes.length === 0) return ""

	const view = pageView(sheet, page)
	const axes = wholeAxes(sheet.layout, sheet.rowSizes)
	const left = axes.cols.offset(page.cols.first)
	const top = axes.rows.offset(page.rows.first)
	const right = axes.cols.offset(page.cols.last + 1)
	const bottom = axes.rows.offset(page.rows.last + 1)
	const paneX = view.cols.header + view.cols.frozenSize
	const paneY = view.rows.header + view.rows.frozenSize
	const mirrored = sheet.layout.rightToLeft
	const place = (box: SheetBox): SheetBox => {
		const x = paneX + box.x - left

		return { x: mirrored ? view.width - x - box.width : x, y: paneY + box.y - top, width: box.width, height: box.height }
	}
	const meets = (box: SheetBox): boolean => box.x < right && box.x + box.width > left && box.y < bottom && box.y + box.height > top
	let out = ""

	for (const object of sheet.objects) {
		const box = anchorBox(object.anchor, axes)

		if (!meets(box)) continue

		const markup = await objectMarkup(job, sheet, object, box)

		if (markup === null) continue

		const at = place(box)

		out += `<g transform="translate(${String(at.x)} ${String(at.y)})">${markup}</g>`
	}

	for (const note of sheet.shownNotes) {
		const box = noteBox(note, axes)

		if (meets(box)) out += shownNoteMarkup(note, place(box), job.measure)
	}

	if (out === "") return ""

	const clip = `${id}-over`
	const clipX = mirrored ? view.width - paneX - (right - left) : paneX

	return `<clipPath id="${clip}"><rect x="${String(clipX)}" y="${String(paneY)}" width="${String(right - left)}" height="${String(bottom - top)}"/></clipPath><g clip-path="url(#${clip})">${out}</g>`
}

const SECTIONS: Readonly<Record<string, "left" | "center" | "right">> = { L: "left", C: "center", R: "right" }

// The header and footer pictures a page shows: the sheet's first-page, even-page or odd-page set (shape ids LH, CF,
// RHFIRST, …).
async function bandPictures(
	job: PrintJob,
	sheet: PrintJobSheet,
	number: number,
	first: boolean
): Promise<{ header: BandPictures; footer: BandPictures }> {
	const header: { -readonly [K in keyof BandPictures]: BandPictures[K] } = {}
	const footer: { -readonly [K in keyof BandPictures]: BandPictures[K] } = {}
	const { headerFooter } = sheet.setup
	const variant = first && headerFooter.differentFirst ? "FIRST" : headerFooter.differentOddEven && number % 2 === 0 ? "EVEN" : ""

	for (const [index, image] of sheet.headerImages.entries()) {
		const section = SECTIONS[image.position[0] ?? ""]

		if (section === undefined || image.position.slice(2) !== variant) continue

		const markup = await headerPictureMarkup(job, sheet, index, image)

		if (markup !== null) (image.position[1] === "H" ? header : footer)[section] = { markup, width: image.width, height: image.height }
	}

	return { header, footer }
}

// A header or footer picture drawn once per print, whichever pages show it.
function headerPictureMarkup(job: PrintJob, sheet: PrintJobSheet, index: number, image: SheetHeaderImage): Promise<string | null> {
	const key = `${String(sheet.sheet)}:header:${String(index)}`
	let made = job.drawn.get(key)

	if (made === undefined) {
		made = (async () => {
			if (image.image.kind === "drawing") return await drawingMarkup(job, image.image.drawing, image.width, image.height)
			if (job.media === null) return null

			const url = await job.media.acquire(image.image.partName, mediaContentType(image.image.partName))

			return url === undefined
				? null
				: `<image href="${escapeSvgText(url)}" width="${String(image.width)}" height="${String(image.height)}" preserveAspectRatio="none"/>`
		})().catch(() => null)
		job.drawn.set(key, made)
	}

	return made
}

function notesPageMarkup(sheet: PrintJobSheet, notes: NotesPage): string {
	const { area } = sheet.pagination

	return notes.lines
		.map(
			(line, k) =>
				`<text x="${String(area.x)}" y="${String(area.y + LIST_FONT_PT + k * LIST_LINE_PT)}" font-size="${String(LIST_FONT_PT)}"${line.bold ? ' font-weight="bold"' : ""}>${escapeSvgText(line.text)}</text>`
		)
		.join("")
}

export interface PrintRenderOptions {
	readonly job: PrintJob
	// The file's name and path (&F, &Z) and the date and time (&D, &T) as the app formats them.
	readonly values: Omit<HeaderFooterValues, "page" | "pages" | "sheet">
	// Keeps the SVGs' ids apart from any other print's in the document.
	readonly id: string
	readonly theme?: SheetGridTheme
}

// The nth page as SVG markup, or null past the last.
export async function renderPrintPage(options: PrintRenderOptions, index: number): Promise<string | null> {
	const { job } = options
	const found = printPageAt(job, index)

	if (found === null) return null

	const id = `${options.id}-${String(index)}`
	const common = {
		printed: found.sheet,
		id,
		number: found.number,
		pages: pageCount(job),
		first: found.first,
		values: { ...options.values, sheet: found.sheet.name },
		pictures:
			found.sheet.headerImages.length === 0
				? { header: {}, footer: {} }
				: await bandPictures(job, found.sheet, found.number, found.first),
		...(options.theme !== undefined && { theme: options.theme })
	}

	if (found.page === null)
		return framePage({ ...common, model: job.model }, found.notes === null ? "" : notesPageMarkup(found.sheet, found.notes))

	const over = await overMarkup(job, found.sheet, found.page, id)
	const tiles = await holdTiles(job.model, found.sheet.sheet, pageTiles(found.page))
	// Drawn in the turn the tiles landed: no frame of the grid evicts them first.
	const model = pageModel(job.model, found.sheet, tiles)
	const svg = printPageSvg({ ...common, model, page: found.page, measure: job.measure, over })

	model.dispose()
	// The grid trims its cache to what it shows on its next frame.
	job.model.repaint()

	return svg
}
