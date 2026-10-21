import { STYLES } from "./xlsxWriteFixtures"

// Workbooks for the conditional format and validation tests: the fixture styles with differential formats.

export { IMMEDIATE, cells, tileEntry } from "./sheetSessionFixtures"

export function STYLES_WITH(dxfs: readonly string[]): string {
	return `${STYLES}<dxfs count="${String(dxfs.length)}">${dxfs.join("")}</dxfs>`
}
