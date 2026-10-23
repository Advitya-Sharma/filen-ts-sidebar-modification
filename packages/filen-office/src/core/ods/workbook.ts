import { Relationships } from "../opc/relationships"
import type { SheetProtectedAction, SheetRange } from "../sheet/protocol"
import { DEFAULT_FORMAT, DEFAULT_VIEW, XlsxSheet } from "../xlsx/read/worksheet"
import type {
	XlsxDefinedName,
	XlsxSection,
	XlsxSheetEntry,
	XlsxSheetModel,
	XlsxSheetProtection,
	XlsxSheetView,
	XlsxRowRun,
	XlsxStyleRuns,
	XlsxWorkbookProtection
} from "../xlsx/model/types"
import type { OdsIssues } from "./issues"
import { rangeAddress, sheetRef } from "./formula/toXlsx"
import type { OdsContent, OdsTable, OdsTableProtection } from "./read/content"
import { NULL_DATE_1899, NULL_DATE_1904 } from "./read/values"
import type { OdsCellStyles } from "./styles/cell"
import type { OdsSettings, OdsTableView } from "./settings"
import { featureParts, type OdsFeatureParts } from "./features"

// The read content as the xlsx workbook model: one sheet entry and model per table, the workbook's names, and an
// in-memory package of the parts the sheet features read (notes, link targets) in their SpreadsheetML form.

// A sheet of an OpenDocument source: the xlsx sheet model with the run index of formats and the runs of rows the
// store does not hold.
export class OdsSheet extends XlsxSheet {
	readonly styleRuns?: XlsxStyleRuns
	readonly rowRuns?: readonly XlsxRowRun[]
	readonly recalculate?: ReadonlySet<number>

	constructor(
		init: ConstructorParameters<typeof XlsxSheet>[0] & {
			readonly styleRuns: XlsxStyleRuns | undefined
			readonly rowRuns: readonly XlsxRowRun[]
			readonly recalculate: ReadonlySet<number>
		}
	) {
		super(init)

		if (init.styleRuns !== undefined) this.styleRuns = init.styleRuns
		if (init.rowRuns.length > 0) this.rowRuns = init.rowRuns
		if (init.recalculate.size > 0) this.recalculate = init.recalculate
	}
}

export interface OdsWorkbookParts {
	readonly entries: XlsxSheetEntry[]
	readonly models: XlsxSheetModel[]
	readonly definedNames: XlsxDefinedName[]
	readonly date1904: boolean
	readonly protection: XlsxWorkbookProtection | null
	readonly activeTab: number
	readonly iterate: { readonly enabled: boolean; readonly count: number; readonly delta: number }
	// Parts by name, with [Content_Types].xml.
	readonly parts: Map<string, Uint8Array>
}

const SELECTABLE: readonly SheetProtectedAction[] = ["selectLockedCells", "selectUnlockedCells"]

function sheetProtection(protection: OdsTableProtection | null): XlsxSheetProtection | null {
	if (protection === null) return null

	const allows = new Set<SheetProtectedAction>()

	for (const action of SELECTABLE) if (protection.allows.get(action) !== false) allows.add(action)
	for (const [action, allowed] of protection.allows)
		if (allowed && !SELECTABLE.includes(action as SheetProtectedAction)) allows.add(action as SheetProtectedAction)

	return {
		hash: {
			...(protection.key !== undefined && protection.key !== "" && { hashValue: protection.key }),
			...(protection.algorithm !== undefined && { algorithmName: protection.algorithm })
		},
		allows
	}
}

function view(table: OdsTableView | undefined, settings: OdsSettings, active: boolean): XlsxSheetView {
	if (table === undefined) return { ...DEFAULT_VIEW, gridLines: settings.showGrid, headings: settings.headings, tabSelected: active }

	const frozenCols = table.horizontalSplitMode === 2 ? Math.max(0, table.horizontalSplitPosition - table.positionLeft) : 0
	const frozenRows = table.verticalSplitMode === 2 ? Math.max(0, table.verticalSplitPosition - table.positionTop) : 0
	const row = Math.max(0, table.cursorRow)
	const col = Math.max(0, table.cursorCol)

	return {
		frozen:
			frozenCols > 0 || frozenRows > 0
				? {
						rows: frozenRows,
						cols: frozenCols,
						topLeftRow: frozenRows > 0 ? table.positionBottom : table.positionTop,
						topLeftCol: frozenCols > 0 ? table.positionRight : table.positionLeft
					}
				: null,
		topRow: table.positionTop,
		leftCol: table.positionLeft,
		activeRow: row,
		activeCol: col,
		selection: [{ startRow: row, startCol: col, endRow: row, endCol: col }],
		zoom: table.zoom > 0 ? table.zoom : 100,
		gridLines: table.showGrid ?? settings.showGrid,
		headings: settings.headings,
		rightToLeft: false,
		tabSelected: active,
		view: "normal"
	}
}

// "$A$1:$D$20" for a range: every part anchored (print areas and titles are absolute).
function absolute(range: SheetRange): string {
	const column = (index: number): string => {
		let text = ""

		for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + ((n - 1) % 26)) + text

		return text
	}

	return `$${column(range.startCol)}$${range.startRow + 1}:$${column(range.endCol)}$${range.endRow + 1}`
}

export function buildSheets(
	content: OdsContent,
	styles: OdsCellStyles,
	settings: OdsSettings,
	issues: OdsIssues,
	prefixes: ReadonlyMap<string, string>
): OdsWorkbookParts {
	const entries: XlsxSheetEntry[] = []
	const models: XlsxSheetModel[] = []
	const definedNames: XlsxDefinedName[] = []
	const activeName = settings.activeTable
	let activeTab = Math.max(
		0,
		content.tables.findIndex(table => table.name === activeName)
	)
	const features: OdsFeatureParts = featureParts(content, styles, issues, prefixes)

	content.tables.forEach((table: OdsTable, index) => {
		const id = index
		const part = `xl/worksheets/sheet${index + 1}.xml`
		const tableStyle = styles.sheetStyle(table.styleName)
		const built = table.built
		const sheetFeatures = features.sheets[index]
		const sections: XlsxSection[] = sheetFeatures?.sections ?? []
		const model = new OdsSheet({
			id,
			part,
			store: built.store,
			sections,
			relationships: new Relationships(sheetFeatures?.relationships ?? []),
			rootStart: 0,
			rootOpenEnd: 0,
			rootEnd: 0,
			prefix: "",
			utf16: false,
			styleRuns: built.styleRuns,
			rowRuns: built.rowRuns,
			recalculate: built.recalculate
		})

		model.cols = built.cols
		model.merges = built.merges
		model.format = { ...DEFAULT_FORMAT, defaultRowHeight: built.defaultRowHeight, defaultColPx: built.defaultColPx }
		model.view = { ...view(settings.tables.get(table.name), settings, index === activeTab), rightToLeft: tableStyle.rightToLeft }
		model.protection = sheetProtection(table.protection)
		model.tabColor = tableStyle.tabColor < 0 ? null : { argb: (0xff000000 | tableStyle.tabColor) >>> 0 }
		entries.push({
			id,
			name: table.name,
			state: tableStyle.display ? "visible" : "hidden",
			kind: "worksheet",
			sheetId: index + 1,
			relationshipId: `rId${index + 1}`,
			part,
			model: null
		})
		models.push(model)

		if (table.printRanges !== undefined && table.printRanges.trim() !== "") {
			const translated = rangeAddress(table.printRanges)

			if ("text" in translated) {
				definedNames.push({
					name: "_xlnm.Print_Area",
					sheet: id,
					formula: anchored(translated.text, table.name),
					hidden: false,
					source: null
				})
			} else {
				issues.add("name", table.name, `print range: ${translated.refused}`)
			}
		}

		const titles: string[] = []

		if (built.headerRows !== null) titles.push(`${sheetRef(table.name)}$${built.headerRows[0] + 1}:$${built.headerRows[1] + 1}`)

		if (built.headerCols !== null) {
			const range = absolute({ startRow: 0, startCol: built.headerCols[0], endRow: 0, endCol: built.headerCols[1] })

			titles.push(`${sheetRef(table.name)}${range.replace(/\$1(?=:|$)/g, "")}`)
		}

		if (titles.length > 0)
			definedNames.push({ name: "_xlnm.Print_Titles", sheet: id, formula: titles.join(","), hidden: false, source: null })
	})

	if (activeTab >= entries.length) activeTab = 0

	// A hidden active sheet is no active sheet: the first visible one shows (as XlsxDocument.activeSheet does).
	// A name is defined once in its scope: the first definition stands (a database range and a named range may share
	// a name; LibreOffice keeps them apart, a workbook cannot).
	const seen = new Set(definedNames.map(name => `${name.sheet ?? ""}\u0000${name.name.toUpperCase()}`))

	for (const name of features.definedNames) {
		const key = `${name.sheet ?? ""}\u0000${name.name.toUpperCase()}`

		if (seen.has(key)) {
			issues.add("name", undefined, `${name.name}: defined twice`)

			continue
		}

		seen.add(key)
		definedNames.push(name)
	}

	const date1904 = content.nullDate === NULL_DATE_1904

	if (content.nullDate !== NULL_DATE_1899 && !date1904) issues.add("nullDate", undefined)

	const iteration = content.calculation?.children.find(child => child.name === "table:iteration")

	return {
		entries,
		models,
		definedNames,
		date1904,
		protection: content.structureProtected
			? {
					structure: true,
					windows: false,
					hash: {
						...(content.protectionKey !== undefined && { hashValue: content.protectionKey }),
						...(content.protectionAlgorithm !== undefined && { algorithmName: content.protectionAlgorithm })
					}
				}
			: null,
		activeTab,
		iterate: {
			enabled: iteration?.attributes.get("table:status") === "enable",
			count: Number(iteration?.attributes.get("table:steps") ?? "100") || 100,
			delta: Number(iteration?.attributes.get("table:minimum-difference") ?? "0.001") || 0.001
		},
		parts: features.parts
	}
}

// A print range as Excel's absolute references on its sheet ("Sheet1!$A$1:$D$20").
function anchored(text: string, sheet: string): string {
	return text
		.split(",")
		.map(part => {
			const bang = part.lastIndexOf("!")
			const prefix = bang < 0 ? sheetRef(sheet) : part.slice(0, bang + 1)
			const area = bang < 0 ? part : part.slice(bang + 1)

			return prefix + area.replace(/\$?([A-Z]{1,3})\$?(\d+)/g, "$$$1$$$2")
		})
		.join(",")
}
