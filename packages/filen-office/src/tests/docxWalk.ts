import { readFileSync } from "node:fs"
import { NumberingCounters, parseNumbering, type ListMarker, type Numbering } from "../core/docx/numbering/index"
import { PropsReader, PropsStore } from "../core/docx/props/index"
import type { ResolvedParagraphProps, ResolvedRunProps } from "../core/docx/props/types"
import {
	PropertyResolver,
	parseStyles,
	readStyleSettings,
	type CellFormat,
	type DocxStyles,
	type ParagraphFormat,
	type StyleSettings
} from "../core/docx/styles/index"
import { parseTheme, type Theme } from "../core/docx/theme/index"
import { OpcPackage, bytesSource } from "../core/opc/index"
import type { XmlDocument } from "../core/xml/index"

// A minimal stand-in for the document model in tests: opens a package, builds the resolver, and walks the
// body's paragraphs (tables and content controls included) resolving every paragraph and run.

export interface OpenedDocx {
	readonly document: XmlDocument
	readonly store: PropsStore
	readonly styles: DocxStyles
	readonly numbering: Numbering | undefined
	readonly theme: Theme | undefined
	readonly settings: StyleSettings
	readonly resolver: PropertyResolver
}

async function optionalXml(pkg: OpcPackage, name: string): Promise<XmlDocument | undefined> {
	return pkg.has(name) ? pkg.xml(name) : undefined
}

export async function openDocx(path: string): Promise<OpenedDocx> {
	const pkg = await OpcPackage.open(bytesSource(new Uint8Array(readFileSync(path))))
	const document = await pkg.xml("word/document.xml")
	const store = new PropsStore()
	const stylesXml = await optionalXml(pkg, "word/styles.xml")
	const styles = stylesXml === undefined ? parseStyles(document, store) : parseStyles(stylesXml, store)
	const numberingXml = await optionalXml(pkg, "word/numbering.xml")
	const numbering = numberingXml === undefined ? undefined : parseNumbering(numberingXml, store, styles)
	const themeXml = await optionalXml(pkg, "word/theme/theme1.xml")
	const theme = themeXml === undefined ? undefined : parseTheme(themeXml)
	const settingsXml = await optionalXml(pkg, "word/settings.xml")
	const settings = settingsXml === undefined ? readStyleSettings(document) : readStyleSettings(settingsXml)

	return { document, store, styles, numbering, theme, settings, resolver: new PropertyResolver({ styles, numbering, theme, settings }) }
}

export interface WalkedParagraph {
	readonly node: number
	readonly format: ParagraphFormat
	readonly props: ResolvedParagraphProps
	readonly runs: readonly ResolvedRunProps[]
	readonly marker: ListMarker | undefined
	readonly cell: CellFormat | undefined
	readonly text: string
}

const INLINE_CONTAINERS = new Set([
	"hyperlink",
	"ins",
	"del",
	"moveFrom",
	"moveTo",
	"smartTag",
	"customXml",
	"fldSimple",
	"sdt",
	"sdtContent",
	"dir",
	"bdo"
])
const BLOCK_CONTAINERS = new Set(["sdt", "sdtContent", "customXml", "ins", "del", "moveFrom", "moveTo"])

export function walkBody(opened: OpenedDocx, visit: (paragraph: WalkedParagraph) => void, collectText: boolean = false): void {
	const { document, store, resolver, numbering } = opened
	const r = PropsReader.for(document)
	const counters = numbering === undefined ? undefined : new NumberingCounters(numbering)

	const runsOf = (node: number, out: ResolvedRunProps[], format: ParagraphFormat, text: string[]): void => {
		for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
			const local = r.local(child)

			if (local === "r") {
				out.push(resolver.run(format, store.run(document, r.child(child, "rPr"))))

				if (collectText) {
					for (let t = document.firstChildElement(child); t >= 0; t = document.nextSiblingElement(t)) {
						if (r.local(t) === "t") text.push(document.textContent(t))
					}
				}
			} else if (INLINE_CONTAINERS.has(local)) {
				runsOf(child, out, format, text)
			}
		}
	}

	const paragraph = (node: number, cell: CellFormat | undefined): void => {
		const format = resolver.paragraph(store.paragraph(document, r.child(node, "pPr")), cell)
		const runs: ResolvedRunProps[] = []
		const text: string[] = []
		const numberingRef = format.props.numbering

		runsOf(node, runs, format, text)
		visit({
			node,
			format,
			props: format.props,
			runs,
			marker: numberingRef === undefined ? undefined : counters?.next(numberingRef.numId, numberingRef.ilvl),
			cell,
			text: text.join("")
		})
	}

	const table = (node: number): void => {
		const tableFormat = resolver.table(store.table(document, r.child(node, "tblPr")))
		const grid = r.child(node, "tblGrid")
		let gridColumnCount = 0

		for (let column = grid < 0 ? -1 : document.firstChildElement(grid); column >= 0; column = document.nextSiblingElement(column)) {
			if (r.local(column) === "gridCol") gridColumnCount++
		}

		const rows: number[] = []

		for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
			if (r.local(child) === "tr") rows.push(child)
		}

		rows.forEach((row, rowIndex) => {
			const rowProps = store.row(document, r.child(row, "trPr"))
			const exceptions = r.child(row, "tblPrEx")
			const rowTable =
				exceptions < 0
					? tableFormat
					: resolver.table(store.table(document, r.child(node, "tblPr")), store.table(document, exceptions))
			let gridColumn = rowProps.gridBefore ?? 0

			resolver.row(rowTable, rowProps, { rowIndex, rowCount: rows.length })

			for (let child = document.firstChildElement(row); child >= 0; child = document.nextSiblingElement(child)) {
				if (r.local(child) !== "tc") continue

				const cellProps = store.cell(document, r.child(child, "tcPr"))
				const gridSpan = cellProps.gridSpan ?? 1
				const cell = resolver.cell(rowTable, cellProps, { rowIndex, rowCount: rows.length, gridColumn, gridSpan, gridColumnCount })

				gridColumn += gridSpan
				blocks(child, cell)
			}
		})
	}

	const blocks = (node: number, cell: CellFormat | undefined): void => {
		for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
			const local = r.local(child)

			if (local === "p") paragraph(child, cell)
			else if (local === "tbl") table(child)
			else if (BLOCK_CONTAINERS.has(local)) blocks(child, cell)
		}
	}

	const body = r.child(document.documentElement(), "body")

	if (body >= 0) blocks(body, undefined)
}

export function walkAll(opened: OpenedDocx, collectText: boolean = false): WalkedParagraph[] {
	const out: WalkedParagraph[] = []

	walkBody(opened, paragraph => out.push(paragraph), collectText)

	return out
}
