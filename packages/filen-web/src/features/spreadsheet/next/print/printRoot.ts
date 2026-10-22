import { printPageAt, renderPrintPage, type PrintRenderOptions } from "@filen/office/react/sheet/print"
import "@/features/spreadsheet/next/print/sheetPrint.css"

// The printable root: the pages as SVG elements in a root of their own beside the app, shown only to the printer
// (sheetPrint.css hides the rest of the document there), each page on a sheet of its size (named @page rules), then
// the browser's print dialog. The root goes away once the dialog closes.

const ROOT_CLASS = "fo-sheet-print-root"
const SVG = "http://www.w3.org/2000/svg"
const XLINK = "http://www.w3.org/1999/xlink"
// How long the print waits for its pictures at most: one that never decodes still prints, without it.
const DECODE_WAIT_MS = 10_000

export interface PrintedPage {
	readonly svg: string
	// Points.
	readonly width: number
	readonly height: number
}

// One measure for the print's text, from a canvas as the grid measures with.
export function canvasMeasure(): (font: string, text: string) => number {
	const context = document.createElement("canvas").getContext("2d")
	let current = ""

	return (font, text) => {
		if (context === null) return text.length * 7

		if (current !== font) {
			context.font = font
			current = font
		}

		return context.measureText(text).width
	}
}

// A page's markup as an element. The markup is our own, every text in it escaped; it is read as XML, so anything
// malformed fails here rather than reaching the document.
export function svgElement(markup: string): SVGSVGElement | null {
	const parsed = new DOMParser().parseFromString(markup, "image/svg+xml")
	const root = parsed.documentElement

	if (root.namespaceURI !== SVG || root.localName !== "svg" || parsed.getElementsByTagName("parsererror").length > 0) return null

	return document.importNode(root, true) as unknown as SVGSVGElement
}

function pageName(width: number, height: number): string {
	return `fo-sheet-${String(Math.round(width * 100))}x${String(Math.round(height * 100))}`
}

// The first `total` pages of a print, drawn one after another (each a few milliseconds on the main thread).
export async function renderPages(options: PrintRenderOptions, total: number, progress: (done: number) => void): Promise<PrintedPage[]> {
	const out: PrintedPage[] = []

	for (let k = 0; k < total; k++) {
		const svg = await renderPrintPage(options, k)
		const at = printPageAt(options.job, k)

		if (svg !== null && at !== null) out.push({ svg, width: at.sheet.pagination.paper.width, height: at.sheet.pagination.paper.height })

		progress(k + 1)
	}

	return out
}

// The pictures the pages show, decoded before the print starts: a browser prints only what has loaded by then.
async function picturesDecoded(root: Element): Promise<void> {
	const urls = new Set<string>()

	for (const image of root.querySelectorAll("image")) {
		const url = image.getAttribute("href") ?? image.getAttributeNS(XLINK, "href")

		if (url !== null && url !== "") {
			urls.add(url)
		}
	}

	if (urls.size === 0) {
		return
	}

	let timer: number | undefined

	await Promise.race([
		Promise.all(
			[...urls].map(url => {
				const picture = new Image()

				picture.src = url

				return picture.decode().catch(() => undefined)
			})
		),
		new Promise<void>(resolve => {
			timer = window.setTimeout(resolve, DECODE_WAIT_MS)
		})
	])
	window.clearTimeout(timer)
}

// Prints `pages`; resolves once the print dialog closed.
export async function printPages(pages: readonly PrintedPage[]): Promise<void> {
	const root = document.createElement("div")
	const style = document.createElement("style")
	const sizes = new Map<string, PrintedPage>()

	root.className = ROOT_CLASS
	root.setAttribute("aria-hidden", "true")

	for (const page of pages) {
		const element = svgElement(page.svg)

		if (element === null) continue

		const name = pageName(page.width, page.height)
		const wrapper = document.createElement("div")

		sizes.set(name, page)
		wrapper.className = "fo-sheet-print-page"
		wrapper.style.setProperty("page", name)
		wrapper.style.width = `${String(page.width)}pt`
		wrapper.style.height = `${String(page.height)}pt`
		wrapper.append(element)
		root.append(wrapper)
	}

	const first = pages[0]

	style.textContent =
		(first === undefined ? "" : `@page { size: ${String(first.width)}pt ${String(first.height)}pt; margin: 0 }`) +
		[...sizes].map(([name, page]) => `@page ${name} { size: ${String(page.width)}pt ${String(page.height)}pt; margin: 0 }`).join("")
	document.head.append(style)
	document.body.append(root)
	await picturesDecoded(root)

	return new Promise(resolve => {
		let done = false
		const finish = (): void => {
			if (done) return

			done = true
			window.removeEventListener("afterprint", finish)
			root.remove()
			style.remove()
			resolve()
		}

		window.addEventListener("afterprint", finish)
		window.print()
		// Browsers whose print() returns before the dialog closes still fire afterprint; one that never does gets the
		// root removed after a while.
		window.setTimeout(finish, 60_000)
	})
}
