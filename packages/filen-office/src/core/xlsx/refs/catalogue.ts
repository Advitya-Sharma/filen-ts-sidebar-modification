import type { SheetRefusal, SheetRefusalReason } from "../../sheet/protocol"

// Every construct that refers to cells, rows, columns or sheets, and how a structure edit moves it. A construct
// the catalogue does not know makes the edit refuse (structureLocked, with the construct as detail).

export const REF_CATALOGUE = {
	cells: "Cell values, styles and row properties: moved with their rows and columns, deleted with them",
	formulas: "Cell formulas on every sheet; array formula and data table ranges and inputs (refused when cut)",
	sharedFormulas: "Shared formula groups: master text and range moved, children that would read differently unshared",
	merges: "Merged ranges: grow, shrink, move; dropped when one cell or none is left",
	cols: "Column ranges (width, style, hidden): grow, shrink, move",
	dimension: "The sheet dimension: moved, and kept covering the cells",
	sheetView: "Active cell, selection and pane (frozen split, top-left cell)",
	breaks: "Manual row and column page breaks",
	hyperlinks: "Hyperlink ranges (dropped with their cells) and internal locations",
	autoFilter: "Autofilter range, filter column ids and its sort state",
	sortState: "The sheet's sort state and sort conditions",
	ignoredErrors: "Ignored error ranges",
	protectedRanges: "Protected ranges (allow edit ranges)",
	conditionalFormatting: "Conditional format ranges, rule formulas and value objects (relative to the range)",
	conditionalFormattingX14: "Conditional formats in the x14 extension: ranges and formulas",
	dataValidations: "Data validation ranges and formulas (relative to the range)",
	dataValidationsX14: "Data validations in the x14 extension: ranges and formulas",
	sparklines: "Sparkline locations and data ranges, the group's date axis range",
	customSheetViews: "Custom views: top-left cell, pane, selection, breaks and autofilter",
	scenarios: "Scenario ranges and changing cells",
	dataConsolidate: "Data consolidation source ranges",
	cellWatches: "Watch window cells",
	smartTags: "Cell smart tags",
	controls: "Form control and OLE object anchors and linked cells in the sheet",
	webPublishItems: "Web publishing source ranges",
	definedNames: "Defined names (print areas and titles included); local names of a deleted sheet go",
	tables: "Table ranges, their autofilter and sort state",
	tableColumns: "Table columns added or removed by column edits, header cell text, calculated and totals formulas",
	structuredReferences: "Structured references to deleted tables or columns become #REF!",
	drawingAnchors: "Drawing object anchors (two-cell, one-cell; absolute ones stay)",
	drawingTextLinks: "Shape text linked to a cell (drawings and chart user shapes)",
	chartFormulas: "Chart series and label formulas (c:f, cx:f)",
	comments: "Note references (dropped with their cells)",
	commentShapes: "Note and form control shapes in the legacy drawing: anchor, cell, linked formulas",
	threadedComments: "Threaded comment references (dropped with their cells)",
	formControls: "Form control properties: linked cell, input range",
	pivotSources: "Pivot cache worksheet sources: locked (an edit that would move one is refused), renamed with the sheet",
	pivotTables: "Pivot table locations: locked",
	slicerSources: "Slicer sources (table columns, pivot tables): an edit that would remove one is refused",
	timelineSources: "Timeline sources (pivot tables): an edit that would remove one is refused",
	connections: "Query parameters bound to cells",
	volatileDependencies: "Real-time data topic cells",
	singleXmlCells: "XML-mapped single cells",
	externalLinks: "References into other workbooks: left as written",
	calcChain: "The calculation chain: dropped (optional, rebuilt by the reader)",
	queryTables: "Query tables: refused when a mapped table column is deleted"
} as const

export type RefConstruct = keyof typeof REF_CATALOGUE

export const REF_CONSTRUCTS = Object.keys(REF_CATALOGUE) as RefConstruct[]

// Planning stops at the first refusal: thrown inside, returned as a SheetRefusal.
export class RefRefused extends Error {
	readonly refusal: SheetRefusal

	constructor(reason: SheetRefusalReason, detail: string) {
		super(`${reason}: ${detail}`)
		this.refusal = { reason, detail }
	}
}

export function refuse(reason: SheetRefusalReason, detail: string): never {
	throw new RefRefused(reason, detail)
}
