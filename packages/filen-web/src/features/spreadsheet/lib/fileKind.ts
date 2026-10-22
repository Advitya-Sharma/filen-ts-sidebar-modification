import type { SpreadsheetKind } from "@filen/office/core/sheet"
import type { SpreadsheetFileKind } from "@/features/spreadsheet/workers/spreadsheet.worker"
import { SHEET_NEXT } from "@/features/spreadsheet/next/sheetNextFlag"

// The one list of spreadsheet extensions: the preview category, the editable gate and the worker's kind
// all read it. A Map, so a name like "constructor" is not found on a prototype.
const KINDS: ReadonlyMap<string, SpreadsheetFileKind> = new Map([
	["xlsx", "xlsx"],
	["xlsm", "xlsx"],
	// Templates: edited in place, saved as the templates they are.
	["xltx", "xlsx"],
	["xltm", "xlsx"],
	["xls", "xls"],
	["csv", "csv"],
	["tsv", "tsv"],
	// The office engine's viewer also opens a legacy template like an .xls (view-only, saved as an .xlsx), and a .tab
	// file as the tab-separated text it is (edited, and saved as it was written).
	...(SHEET_NEXT
		? ([
				["xlt", "xls"],
				["tab", "tsv"]
			] as const)
		: [])
])

export const SPREADSHEET_EXTENSIONS: readonly string[] = [...KINDS.keys()]

export function spreadsheetFileKind(extension: string): SpreadsheetFileKind | null {
	return KINDS.get(extension) ?? null
}

// A legacy .xls (or .xlt) opens to be looked at only; every other kind saves in the format its extension names.
export function isEditableSpreadsheetExtension(extension: string): boolean {
	const kind = spreadsheetFileKind(extension)

	return kind !== null && kind !== "xls"
}

// The format a name promises, for telling whether a rename changed it: its kind, except that .xlsx, .xlsm
// (macros kept) and the templates open alike but are different files.
export function spreadsheetSaveFormat(extension: string): string | null {
	const kind = spreadsheetFileKind(extension)

	return kind === "xlsx" ? extension : kind
}

const ZIP = [0x50, 0x4b, 0x03, 0x04]
const OLE2 = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

function startsWith(bytes: Uint8Array, magic: readonly number[]): boolean {
	return bytes.length >= magic.length && magic.every((byte, index) => bytes[index] === byte)
}

// A file whose name does not say what it is, read from its first bytes: a zip is taken for an .xlsx, an
// OLE2 compound file for an .xls (or a password-protected workbook, which the worker tells apart), anything
// else for text.
export function sniffSpreadsheetKind(bytes: Uint8Array): SpreadsheetKind {
	if (startsWith(bytes, ZIP)) return "xlsx"
	if (startsWith(bytes, OLE2)) return "xls"

	return "csv"
}
