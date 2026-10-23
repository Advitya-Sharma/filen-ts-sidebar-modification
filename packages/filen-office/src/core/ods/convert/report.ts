// What a conversion between SpreadsheetML and OpenDocument could not carry over, by kind, so the app can say so.
// Nothing is dropped unnamed: each item is converted, kept in a reduced form its kind states, or counted here.

export type SheetLossyKind =
	// Formulas one grammar has and the other cannot say (OpenFormula ⟷ Excel): written as their cached values, an
	// error as that error alone (of:=#N/A, shown as #N/A), one OpenFormula has no name for as its text.
	| "formula"
	| "conditionalFormat"
	| "dataValidation"
	// AutoFilter criteria or sort state the target cannot say.
	| "filter"
	// Named ranges or expressions the target cannot say.
	| "definedName"
	| "chart"
	// Pictures, shapes, text boxes, controls and other drawing objects.
	| "picture"
	| "note"
	// Excel's threaded comments: their first comment stays as the cell's note.
	| "threadedComment"
	// Pivot tables and DataPilot tables: their cells stay as values.
	| "pivotTable"
	// Excel tables (ListObjects): their cells stay, the table does not.
	| "table"
	| "sparkline"
	// Page setup, print options, headers and footers, page breaks.
	| "pageSetup"
	// Cell format properties the target cannot say (the rest of the format is kept).
	| "cellFormat"
	// Protection kept without its password, whose hash the target cannot hold.
	| "protectionPassword"
	// Links to other workbooks and data sources (formulas keep their cached values).
	| "externalLink"
	| "macros"
	// Document properties (title, author and the like).
	| "metadata"
	| "outline"
	// Formatted runs of text in a cell: written as plain text.
	| "richText"
	// A sheet name the target does not take (too long, a character it forbids): the sheet renamed, what names it
	// following.
	| "sheetName"
	// Anything else the source holds that the target does not: among them an error value OpenFormula has no name for
	// (#NULL!, #SPILL!, #CALC!), written as its text.
	| "other"

export interface SheetUntranslatedFormula {
	readonly sheet: string
	// "B7"; a block of repeated cells is its first and last ("B7:B9").
	readonly cell: string
	// As the source writes it (OpenFormula with its namespace prefix, or Excel's without "=").
	readonly formula: string
}

export interface SheetLossyFeature {
	readonly kind: SheetLossyKind
	readonly count: number
	// The sheets it is on, in workbook order; none for what is the workbook's.
	readonly sheets: readonly string[]
	// For "formula": each one.
	readonly formulas?: readonly SheetUntranslatedFormula[]
}

// The order kinds are listed in.
const ORDER: readonly SheetLossyKind[] = [
	"formula",
	"conditionalFormat",
	"dataValidation",
	"filter",
	"definedName",
	"chart",
	"picture",
	"note",
	"threadedComment",
	"pivotTable",
	"table",
	"sparkline",
	"pageSetup",
	"cellFormat",
	"protectionPassword",
	"externalLink",
	"macros",
	"metadata",
	"outline",
	"richText",
	"sheetName",
	"other"
]

export class SheetLossyReport {
	private readonly counts = new Map<SheetLossyKind, { count: number; sheets: Set<string> }>()
	private readonly formulas: SheetUntranslatedFormula[] = []
	private readonly sheetOrder: readonly string[]

	constructor(sheetOrder: readonly string[]) {
		this.sheetOrder = sheetOrder
	}

	note(kind: SheetLossyKind, sheet: string | undefined, count = 1): void {
		if (count <= 0) return

		let entry = this.counts.get(kind)

		if (entry === undefined) {
			entry = { count: 0, sheets: new Set() }
			this.counts.set(kind, entry)
		}

		entry.count += count

		if (sheet !== undefined) entry.sheets.add(sheet)
	}

	formula(sheet: string, cell: string, formula: string, cells = 1): void {
		this.formulas.push({ sheet, cell, formula })
		this.note("formula", sheet, cells)
	}

	list(): SheetLossyFeature[] {
		const out: SheetLossyFeature[] = []

		for (const kind of ORDER) {
			const entry = this.counts.get(kind)

			if (entry === undefined) continue

			const sheets = this.sheetOrder.filter(name => entry.sheets.has(name))

			out.push(
				kind === "formula"
					? { kind, count: entry.count, sheets, formulas: [...this.formulas] }
					: { kind, count: entry.count, sheets }
			)
		}

		return out
	}
}
