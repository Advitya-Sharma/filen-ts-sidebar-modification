// The SpreadsheetML and relationship namespaces of a workbook: Transitional, or Strict (ECMA-376 Part 1,
// Annex A), whose relationship types live under its own relationships namespace.

export interface SpreadsheetNamespaces {
	readonly main: string
	readonly relationships: string
}

export const TRANSITIONAL: SpreadsheetNamespaces = {
	main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
	relationships: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}

export const STRICT: SpreadsheetNamespaces = {
	main: "http://purl.oclc.org/ooxml/spreadsheetml/main",
	relationships: "http://purl.oclc.org/ooxml/officeDocument/relationships"
}

export function namespacesOf(strict: boolean): SpreadsheetNamespaces {
	return strict ? STRICT : TRANSITIONAL
}

// The content types of sheet parts (Transitional and Strict share them).
export const SHEET_CONTENT_TYPES: ReadonlySet<string> = new Set([
	"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
	"application/vnd.ms-excel.macrosheet+xml",
	"application/vnd.ms-excel.intlmacrosheet+xml"
])

// A relationship type ("sharedStrings", "styles", "calcChain", "worksheet") in the workbook's namespace.
export function relationshipType(strict: boolean, name: string): string {
	return `${namespacesOf(strict).relationships}/${name}`
}
