// Names of the engine-only functions the engine's formula text calls (refs/formulaRefs.ts writes them; the plugins in formula/functions
// define them). Kept apart from the plugins so formula text can be written without loading the engine.

// FILEN.CRITERION(value): a criteria argument of COUNTIF and its kind, an empty cell read as 0.
export const CRITERION_FUNCTION = "FILEN.CRITERION"

// FILEN.INTERSECT(range): the cell of a range in the formula's own row or column, as Excel reads a range where one
// value is wanted (implicit intersection), on any sheet.
export const INTERSECT_FUNCTION = "FILEN.INTERSECT"
