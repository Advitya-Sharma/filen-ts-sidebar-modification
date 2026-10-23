import { parseHeaderFooter, type HeaderFooterRun, type HeaderFooterStyle } from "../../sheet/print/headerFooter"
import { paperSizeKnown } from "../../sheet/print/paper"
import { readPageSetup, readTitles, type SheetPageSetup } from "../../sheet/print/setup"
import { escapeAttribute } from "../../xml/index"
import { CellRefused, paragraphText } from "../write/cells"
import type { TitleSpan } from "../write/table"

// A worksheet's print settings (ECMA-376 Part 1 §18.3.1.62 pageMargins, §18.3.1.63 pageSetup, §18.3.1.70
// printOptions, §18.3.1.46 headerFooter, sheetPr's fitToPage) as an OpenDocument page layout and master page (ODF 1.3
// Part 3 §16.5 style:page-layout, §16.9 style:master-page, §17.2 style:page-layout-properties, §16.10–16.12 headers
// and footers with the fields of §7.3). What neither holds is lost: error cells printed otherwise than shown, black and
// white, draft quality, copies, notes printed where shown, a page number offset, header pictures, headers not aligned
// with the margins or not scaled with the sheet, an unknown paper.

export interface SheetPageMarkup {
	readonly pageSetup: string
	readonly pageMargins: string
	readonly printOptions: string
	readonly headerFooter: string
	readonly rowBreaks: string
	readonly colBreaks: string
	readonly sheetPr: string
}

export interface SheetPage {
	// The master page the sheet's table style names.
	readonly masterPage: string
	readonly rowBreaks: ReadonlySet<number>
	readonly colBreaks: ReadonlySet<number>
	readonly titleRows: TitleSpan | null
	readonly titleCols: TitleSpan | null
	// Something the page does not hold.
	readonly lost: boolean
}

function inches(value: number): string {
	return `${String(Math.round(Math.max(0, value) * 10_000) / 10_000)}in`
}

function points(value: number): string {
	return `${String(Math.round(value * 1000) / 1000)}pt`
}

function attributes(pairs: readonly (readonly [string, string | undefined])[]): string {
	return pairs.map(([name, value]) => (value === undefined ? "" : ` ${name}="${escapeAttribute(value)}"`)).join("")
}

// Page layouts, master pages and the text styles of their headers, for styles.xml (each the same once).
export class PageStyles {
	private readonly layouts = new Map<string, string>()
	private readonly masters = new Map<string, string>()
	private readonly texts = new Map<string, string>()
	private readonly layoutMarkup: string[] = []
	private readonly masterMarkup: string[] = []
	private readonly textMarkup: string[] = []

	// styles.xml's office:automatic-styles children and office:master-styles children.
	get automatic(): string {
		return `${this.textMarkup.join("")}${this.layoutMarkup.join("")}`
	}

	get master(): string {
		return this.masterMarkup.join("")
	}

	sheet(markup: SheetPageMarkup, printTitles: string | null): SheetPage {
		const setup = readPageSetup(markup, { printArea: null, printTitles })
		const sections = this.headers(setup)
		const layout = this.layout(setup, markup.pageSetup.trim() !== "", sections.header !== "", sections.footer !== "")
		const key = `${layout}\u0000${sections.markup}`
		let masterPage = this.masters.get(key)

		if (masterPage === undefined) {
			masterPage = `Sheet${String(this.masters.size + 1)}`
			this.masters.set(key, masterPage)
			this.masterMarkup.push(
				`<style:master-page${attributes([
					["style:name", masterPage],
					["style:page-layout-name", layout]
				])}>${sections.markup}</style:master-page>`
			)
		}

		const titles = printTitles === null ? { rows: null, cols: null } : readTitles(printTitles)

		return {
			masterPage,
			rowBreaks: new Set(setup.rowBreaks),
			colBreaks: new Set(setup.colBreaks),
			titleRows: titles.rows,
			titleCols: titles.cols,
			lost: sections.lost || lostSetup(markup, setup)
		}
	}

	private layout(setup: SheetPageSetup, paperGiven: boolean, header: boolean, footer: boolean): string {
		const margins = setup.margins
		const landscape = setup.orientation === "landscape"
		const { width, height } = setup.paper
		// The paper as the file names it; a sheet without pageSetup prints on the reader's own.
		const paper = paperGiven ? (landscape ? [height, width] : [width, height]) : null
		const print = ["objects", "charts", "drawings", "zero-values"]

		if (setup.cellComments !== "none") print.push("annotations")
		if (setup.gridLines) print.push("grid")
		if (setup.headings) print.push("headers")

		const centering =
			setup.horizontalCentered && setup.verticalCentered
				? "both"
				: setup.horizontalCentered
					? "horizontal"
					: setup.verticalCentered
						? "vertical"
						: "none"
		// A header sits at its margin from the top, the sheet at the top margin below it (§16.10 fo:min-height); one not
		// aligned with the margins (alignWithMargins) spans the paper, its margins the page's taken back.
		const aligned = setup.headerFooter.alignWithMargins
		const band = (edge: number, at: number): string =>
			`<style:header-footer-properties${attributes([
				["fo:min-height", inches(edge - at)],
				["fo:margin-left", aligned ? "0in" : `-${inches(margins.left)}`],
				["fo:margin-right", aligned ? "0in" : `-${inches(margins.right)}`],
				["fo:margin-top", "0in"],
				["fo:margin-bottom", "0in"]
			])}/>`
		const properties = attributes([
			["fo:page-width", paper === null ? undefined : points(paper[0] ?? 0)],
			["fo:page-height", paper === null ? undefined : points(paper[1] ?? 0)],
			["style:print-orientation", landscape ? "landscape" : "portrait"],
			["fo:margin-top", inches(header ? margins.header : margins.top)],
			["fo:margin-bottom", inches(footer ? margins.footer : margins.bottom)],
			["fo:margin-left", inches(margins.left)],
			["fo:margin-right", inches(margins.right)],
			["style:print", print.join(" ")],
			["style:print-page-order", setup.pageOrder === "overThenDown" ? "ltr" : "ttb"],
			["style:first-page-number", setup.firstPageNumber === null ? undefined : String(Math.max(1, setup.firstPageNumber))],
			...(setup.fitToPage
				? ([
						["style:scale-to-X", setup.fitToWidth > 0 ? String(setup.fitToWidth) : undefined],
						["style:scale-to-Y", setup.fitToHeight > 0 ? String(setup.fitToHeight) : undefined]
					] as const)
				: ([["style:scale-to", `${String(setup.scale)}%`]] as const)),
			["style:table-centering", centering],
			["style:writing-mode", "lr-tb"]
		])
		const markup =
			`<style:page-layout-properties${properties}/>` +
			`<style:header-style>${header ? band(margins.top, margins.header) : ""}</style:header-style>` +
			`<style:footer-style>${footer ? band(margins.bottom, margins.footer) : ""}</style:footer-style>`
		let name = this.layouts.get(markup)

		if (name === undefined) {
			name = `pm${String(this.layouts.size + 1)}`
			this.layouts.set(markup, name)
			this.layoutMarkup.push(`<style:page-layout style:name="${name}">${markup}</style:page-layout>`)
		}

		return name
	}

	// The master page's headers and footers: odd pages', even pages' (style:header-left), the first page's
	// (style:header-first); none where the file has none.
	private headers(setup: SheetPageSetup): { header: string; footer: string; markup: string; lost: boolean } {
		const hf = setup.headerFooter
		let lost = false
		const one = (element: string, text: string): string => {
			const sections = parseHeaderFooter(text)
			const region = (name: string, runs: readonly HeaderFooterRun[]): string => {
				// Each line's pieces, those of one look in one span.
				const lines: { style: string | undefined; body: string }[][] = [[]]
				const add = (style: HeaderFooterStyle, body: string): void => {
					const line = lines[lines.length - 1] ?? []
					const own = this.textStyle(style)
					const last = line[line.length - 1]

					if (last !== undefined && last.style === own) last.body += body
					else line.push({ style: own, body })
				}

				for (const run of runs) {
					if (run.kind === "field") {
						if (run.field === "picture" || run.offset !== 0) lost = true

						add(run.style, fieldMarkup(run))

						continue
					}

					run.text.split(/\r?\n/).forEach((part, k) => {
						if (k > 0) lines.push([])
						if (part !== "")
							add(
								run.style,
								textOf(part, () => (lost = true))
							)
					})
				}

				const paragraph = (line: readonly { style: string | undefined; body: string }[]): string =>
					line
						.map(piece =>
							piece.style === undefined || piece.body === ""
								? piece.body
								: `<text:span text:style-name="${piece.style}">${piece.body}</text:span>`
						)
						.join("")

				return `<style:${name}>${lines.map(line => `<text:p>${paragraph(line)}</text:p>`).join("")}</style:${name}>`
			}

			return `<style:${element}>${region("region-left", sections.left)}${region("region-center", sections.center)}${region("region-right", sections.right)}</style:${element}>`
		}
		const odd = hf.odd
		const even = hf.differentOddEven ? hf.even : odd
		const first = hf.differentFirst ? hf.first : odd
		const kind = (pick: (texts: { header: string; footer: string }) => string, name: "header" | "footer"): string => {
			if (pick(odd) === "" && pick(even) === "" && pick(first) === "") return ""

			return (
				one(name, pick(odd)) +
				(hf.differentOddEven ? one(`${name}-left`, pick(even)) : "") +
				(hf.differentFirst ? one(`${name}-first`, pick(first)) : "")
			)
		}
		const header = kind(texts => texts.header, "header")
		const footer = kind(texts => texts.footer, "footer")

		// Text not scaled with the sheet has no home; without any text it scales nothing.
		if (!hf.scaleWithDoc && (header !== "" || footer !== "")) lost = true

		return { header, footer, markup: header + footer, lost }
	}

	// A look's text style, undefined for plain text.
	private textStyle(style: HeaderFooterStyle): string | undefined {
		const properties = attributes([
			["fo:font-family", style.font ?? undefined],
			["fo:font-size", style.size === null ? undefined : `${String(style.size)}pt`],
			["fo:font-weight", style.bold ? "bold" : undefined],
			["fo:font-style", style.italic ? "italic" : undefined],
			["style:text-underline-style", style.underline === null ? undefined : "solid"],
			["style:text-underline-type", style.underline === "double" ? "double" : style.underline === null ? undefined : "single"],
			["style:text-underline-color", style.underline === null ? undefined : "font-color"],
			["style:text-line-through-style", style.strike ? "solid" : undefined],
			["style:text-position", style.script === "superscript" ? "super 58%" : style.script === "subscript" ? "sub 58%" : undefined],
			["fo:color", style.color === null ? undefined : `#${style.color.toString(16).padStart(6, "0")}`]
		])

		if (properties === "") return undefined

		let name = this.texts.get(properties)

		if (name === undefined) {
			name = `HF${String(this.texts.size + 1)}`
			this.texts.set(properties, name)
			this.textMarkup.push(
				`<style:style style:name="${name}" style:family="text"><style:text-properties${properties}/></style:style>`
			)
		}

		return name
	}
}

function textOf(text: string, lose: () => void): string {
	try {
		return paragraphText(text)
	} catch (error) {
		if (!(error instanceof CellRefused)) throw error

		lose()

		return paragraphText(
			[...text]
				.filter(c => {
					const code = c.charCodeAt(0)

					return c === "\t" || (code >= 0x20 && code !== 0xfffe && code !== 0xffff)
				})
				.join("")
		)
	}
}

// A field as ODF's (§7.3.4 page number, §7.3.5 page count, §7.3.2 date, §7.3.3 time, §7.3.6 file name, §7.3.7 sheet
// name); a picture has none.
function fieldMarkup(run: Extract<HeaderFooterRun, { kind: "field" }>): string {
	switch (run.field) {
		case "page":
			return `<text:page-number${run.offset === 0 ? "" : ` text:page-adjust="${String(run.offset)}"`}>1</text:page-number>`
		case "pages":
			return "<text:page-count>1</text:page-count>"
		case "date":
			return "<text:date/>"
		case "time":
			return "<text:time/>"
		case "file":
			return '<text:file-name text:display="name-and-extension"/>'
		case "path":
			return '<text:file-name text:display="path"/>'
		case "sheet":
			return "<text:sheet-name/>"
		case "picture":
			return ""
	}
}

function lostSetup(markup: SheetPageMarkup, setup: SheetPageSetup): boolean {
	return (
		setup.blackAndWhite ||
		setup.draft ||
		setup.errors !== "displayed" ||
		setup.cellComments === "asDisplayed" ||
		(!setup.customPaper && markup.pageSetup.trim() !== "" && !paperSizeKnown(setup.paperSize)) ||
		/\scopies="(?!1")[^"]*"/.test(markup.pageSetup)
	)
}
