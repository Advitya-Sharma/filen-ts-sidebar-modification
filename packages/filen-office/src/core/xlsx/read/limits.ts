import { XlsxReadError } from "./errors"

// Caps on what one workbook may hold, checked as it is read, so a hostile file costs bounded memory and time.
// A platform with less memory passes lower ones.
export interface XlsxLimits {
	readonly maxSheets: number
	// Cells (stored slots) across every parsed sheet.
	readonly maxCells: number
	readonly maxSharedStrings: number
	// Inflated bytes of one worksheet part.
	readonly maxSheetPart: number
	// Inflated bytes of workbook.xml, styles.xml and the theme.
	readonly maxMetaPart: number
	readonly maxSharedStringsPart: number
	// Elements of workbook.xml or styles.xml, counted before either is parsed into a tree.
	readonly maxMetaElements: number
	// Children of the root (sections) across workbook.xml and every parsed sheet: each is kept as its own copy.
	readonly maxSections: number
	// Merged ranges across every parsed sheet.
	readonly maxMerges: number
	readonly maxDefinedNames: number
	// Records of one kind in styles.xml (fonts, fills, borders, xfs, dxfs, numFmts).
	readonly maxStyleRecords: number
	// Column ranges across every parsed sheet (one sheet holds at most one per column).
	readonly maxColRanges: number
}

export const DEFAULT_XLSX_LIMITS: XlsxLimits = {
	maxSheets: 4096,
	// The previous engine's cap.
	maxCells: 5_000_000,
	maxSharedStrings: 16_000_000,
	maxSheetPart: 512 * 1024 * 1024,
	maxMetaPart: 64 * 1024 * 1024,
	maxSharedStringsPart: 512 * 1024 * 1024,
	maxMetaElements: 2_000_000,
	maxSections: 131_072,
	maxMerges: 1_000_000,
	maxDefinedNames: 1_000_000,
	maxStyleRecords: 1_000_000,
	maxColRanges: 262_144
}

// Distinct strings one sheet's own table holds: a JS Map refuses past 2^24 entries, so the reader refuses first.
export const MAX_SHEET_STRINGS = 16_000_000

type BudgetKind = "cells" | "sections" | "merges" | "colRanges"

const BUDGET_MESSAGES: Readonly<Record<BudgetKind, string>> = {
	cells: "The workbook holds more cells than the limit",
	sections: "The workbook holds more sections than the limit",
	merges: "Too many merged ranges",
	colRanges: "Too many column ranges"
}

// What every parsed sheet of one workbook may still take, shared by sheets parsed at the same time.
export class XlsxReadBudget {
	private readonly left: Record<BudgetKind, number>

	constructor(limits: XlsxLimits, workbookSections: number) {
		this.left = {
			cells: limits.maxCells,
			sections: limits.maxSections - workbookSections,
			merges: limits.maxMerges,
			colRanges: limits.maxColRanges
		}
	}

	remaining(kind: BudgetKind): number {
		return this.left[kind]
	}

	// One sheet's reading: what it takes is given back when the reading fails.
	lease(part: string): XlsxReadLease {
		return new XlsxReadLease(this.left, part)
	}
}

export class XlsxReadLease {
	private readonly left: Record<BudgetKind, number>
	private readonly part: string
	private readonly taken: Record<BudgetKind, number> = { cells: 0, sections: 0, merges: 0, colRanges: 0 }

	constructor(left: Record<BudgetKind, number>, part: string) {
		this.left = left
		this.part = part
	}

	take(kind: BudgetKind, count: number): void {
		if (count > this.left[kind]) throw new XlsxReadError("tooLarge", BUDGET_MESSAGES[kind], this.part)

		this.left[kind] -= count
		this.taken[kind] += count
	}

	release(): void {
		for (const kind of Object.keys(this.taken) as BudgetKind[]) {
			this.left[kind] += this.taken[kind]
			this.taken[kind] = 0
		}
	}
}
