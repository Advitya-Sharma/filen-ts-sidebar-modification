// "preview" namespace, spreadsheet viewer — typing formulas: function suggestions, the syntax of
// the function being typed, and AutoSum. Spread into preview.ts; same catalog rules.
export const previewSheetFormula = {
	/** Spreadsheet formula typing — accessible name of the list of function names offered while a formula is typed */
	previewSheetFormulaSuggestions: "Functions",
	/** Spreadsheet formula typing — accessible name of the hint showing the arguments of the function being typed */
	previewSheetFormulaSignature: "Function arguments",
	/** Spreadsheet formula bar — button that adds up the numbers above or to the left of the selection (Excel's AutoSum) */
	previewSheetAutoSum: "AutoSum",
	/** Spreadsheet formula bar — icon and tooltip beside a formula whose result this app does not calculate: the cell shows the result saved in the file, which may be out of date after edits */
	previewSheetCachedResult: "Not recalculated here: this shows the result saved in the file"
}
