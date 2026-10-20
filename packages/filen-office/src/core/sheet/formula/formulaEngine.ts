import {
	CellError,
	DenseSparseChooseBasedOnThreshold,
	DetailedCellError,
	ErrorType,
	FunctionArgumentType,
	FunctionPlugin,
	HyperFormula,
	type ConfigParams,
	type ImplementedFunctions,
	type RawCellContent,
	type SimpleCellAddress
} from "hyperformula"
import type { AxisEdit } from "../model/edits"
import { formulaInputs, sameSheet } from "../refs/formulaRefs"
import { MAX_COLUMNS, MAX_ROWS } from "../model/model"
import type { SizeAxis } from "../sizes/sizes"

// HyperFormula over a workbook, for recalculating what an edit changes. Built only once the workbook has a
// formula to care about, and kept in step with the workbook from then on (cells, inserted and deleted rows
// and columns, added and renamed sheets). Formula text is never read back from it: references are moved in
// the workbook's own text (refs/formulaRefs.ts), which stays Excel's.
// GPL-3.0 licensed, as the app (AGPL-3.0) may combine it.
const LICENSE_KEY = "gpl-v3"

export interface RecalculatedCell {
	sheet: number
	row: number
	col: number
	// A number, string, boolean, error text ("#DIV/0!") or null. Dates come as serial numbers.
	value: number | string | boolean | null
	// The value is an error, not text that reads like one.
	error?: true
	// A result held from the file (HeldResult) whose inputs changed: it is no longer what the formula gives. The
	// value is the held one.
	stale?: true
}

// A formula's stored result, held by the engine in its place: the formula is not calculated here (a function the
// engine lacks, an array formula, syntax it cannot read). It depends on no cell there: the caller knows what the
// formula reads and tells when that changes; the engine reports it stale at each inserted or deleted row or column.
export class HeldResult {
	readonly value: RawCellContent

	constructor(value: RawCellContent) {
		this.value = value
	}
}

// A cell to give the engine. A formula also carries the result the file stored, when it has one: should
// the engine not read the formula (syntax it lacks) or not know a function in it, the cell holds that
// result instead, so what depends on it still calculates.
export interface EngineCell {
	row: number
	col: number
	content: RawCellContent | HeldResult
	fallback?: RawCellContent
}

export interface EngineSheet {
	name: string
	// A held cell (HeldResult) is listed in `formulas` too.
	rows: (RawCellContent | HeldResult)[][]
	// Its formula cells that carry a fallback or hold a result.
	formulas: EngineCell[]
}

export interface EngineName {
	name: string
	expression: string
	// Grid sheet index of a sheet-scoped name.
	scope?: number
}

// A circular reference. Excel, not iterating, shows 0 for it; iterating, the result it stored.
export const CYCLE = "#CYCLE!"

// Excel's own error texts, which the files store and the grid shows.
const ERROR_TEXT: Readonly<Partial<Record<ErrorType, string>>> = {
	[ErrorType.DIV_BY_ZERO]: "#DIV/0!",
	[ErrorType.NAME]: "#NAME?",
	[ErrorType.VALUE]: "#VALUE!",
	[ErrorType.NUM]: "#NUM!",
	[ErrorType.NA]: "#N/A",
	[ErrorType.REF]: "#REF!",
	[ErrorType.CYCLE]: CYCLE,
	[ErrorType.SPILL]: "#SPILL!"
}

const INSERT_RUN = 10_000

// The most cells the engine may copy for one removal: it copies what it removes for an undo of its own (every cell
// of the removed rows across the sheet's width, of a removed sheet), which far cells make billions.
const MAX_COPIED_CELLS = 1_048_576

// An operation the engine cannot do within bounds: the caller refuses the edit, or drops the engine and builds it
// again from the cells when next needed.
export class EngineLimitError extends Error {
	constructor() {
		super("spreadsheet: the formula engine would copy too many cells")
	}
}

// Failing formulas at build past which the engine is built again with them held.
const REBUILD_FAILING = 256

// What the engine gives back for a formula it could not read at all.
const PARSE_ERROR = "#ERROR!"

// Content standing in for a formula the engine cannot be given: text reading as its parse error.
export const UNPARSEABLE = `'${PARSE_ERROR}`

// A recalculated cell's key in a batch's changes.
function changeKey(address: SimpleCellAddress): string {
	return `${String(address.sheet)},${String(address.row)},${String(address.col)}`
}

function exported(value: unknown): RecalculatedCell["value"] {
	if (value instanceof DetailedCellError) {
		return ERROR_TEXT[value.type] ?? value.value
	}

	return typeof value === "number" || typeof value === "string" || typeof value === "boolean" ? value : null
}

interface Pending {
	address: SimpleCellAddress
	content: RawCellContent | HeldResult
	fallback: RawCellContent | undefined
}

// What the held-result functions answer from: the engine running now (calls into HyperFormula are synchronous).
type HeldAnswer = (id: number, address: SimpleCellAddress) => number | string | boolean | CellError

let running: HeldAnswer | null = null

const HELD = "FILEN.HELD"

// FILEN.HELD(id): the result held under `id`. It is recalculated at each inserted or deleted row or column, which is
// a held result's inputs changing.
class HeldResultPlugin extends FunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		[HELD]: { method: "held", parameters: [{ argumentType: FunctionArgumentType.ANY }], isDependentOnSheetStructureChange: true }
	}

	held(ast: { readonly args: readonly { readonly value?: unknown }[] }, state: { readonly formulaAddress: SimpleCellAddress }): unknown {
		const id = ast.args[0]?.value

		return running === null || typeof id !== "number" ? new CellError(ErrorType.VALUE) : running(id, state.formulaAddress)
	}
}

let registered = false

function registerHeld(): void {
	if (registered) return

	HyperFormula.registerFunctionPlugin(HeldResultPlugin, { enGB: { [HELD]: HELD } })
	registered = true
}

const ERROR_TYPES = new Map<string, ErrorType>(
	Object.entries(ERROR_TEXT).flatMap(([type, text]) => (text === CYCLE ? [] : [[text, type as ErrorType]]))
)

// What a formula naming a sheet the workbook does not have shows (Excel's reading of it).
const REF_ERROR = "#REF!"

function engineConfig(date1904: boolean): Partial<ConfigParams> {
	return {
		licenseKey: LICENSE_KEY,
		// Ordinary formulas as Excel calculates them: a range where one value is expected meets the
		// formula's row or column. Array-evaluating calls are marked with ARRAYFORMULA (refs/formulaRefs.ts).
		useArrayArithmetic: false,
		// A sheet's cells are held densely (one slot per cell of its used rectangle) only when they fill most of
		// it: by default HyperFormula always does, so one value far right and one far down need gigabytes.
		chooseAddressMappingPolicy: new DenseSparseChooseBasedOnThreshold(0.8),
		maxRows: MAX_ROWS,
		maxColumns: MAX_COLUMNS,
		// Excel's serial dates (1900-02-29 included, or counted from 1904), whitespace and empty-cell
		// arithmetic.
		leapYear1900: !date1904,
		nullDate: date1904 ? { year: 1904, month: 1, day: 1 } : { year: 1899, month: 12, day: 31 },
		ignoreWhiteSpace: "any",
		evaluateNullToZero: true,
		// Undo is the document's own; the engine's would hold copies of what each operation replaced.
		undoLimit: 0
	}
}

// The most cells one range of a formula may cover, within the sheet it names as far as that sheet reaches: the
// engine reads a range's every cell (a whole-column VLOOKUP over a sheet with a cell in its last row is twelve
// million).
const MAX_RANGE_CELLS = 2_097_152

interface SheetExtent {
	readonly name: string
	readonly height: number
	readonly width: number
}

// Why the engine is not given a formula: it names a sheet the workbook lacks ("sheet": the engine would keep a
// placeholder for it, and fails on the cell when it is written next), or a range past what the engine can read
// ("range"). Null when it is given.
function unreadable(
	content: unknown,
	sheets: readonly SheetExtent[],
	own: number,
	reach?: (sheet: string, cols: number, rows: number) => void
): "sheet" | "range" | null {
	if (typeof content !== "string" || !content.startsWith("=") || content.startsWith(`=${HELD}(`)) return null
	if (!content.includes("!") && !content.includes(":")) return null

	for (const ref of formulaInputs(content.slice(1)).refs) {
		// The engine has no 3-D references.
		if (ref.lastSheet !== null) return "sheet"

		const first = ref.sheet === null ? sheets[own] : sheets.find(sheet => sameSheet(sheet.name, ref.sheet ?? ""))

		if (first === undefined || (ref.lastSheet !== null && !sheets.some(sheet => sameSheet(sheet.name, ref.lastSheet ?? ""))))
			return "sheet"

		const area = ref.area

		if (!area.range && area.kind === "cell") continue

		const top = area.kind === "cols" ? 0 : Math.min(area.startRow, area.endRow)
		const bottom = area.kind === "cols" ? first.height - 1 : Math.min(Math.max(area.startRow, area.endRow), first.height - 1)
		const left = area.kind === "rows" ? 0 : Math.min(area.startCol, area.endCol)
		const right = area.kind === "rows" ? first.width - 1 : Math.min(Math.max(area.startCol, area.endCol), first.width - 1)

		if (Math.max(0, bottom - top + 1) * Math.max(0, right - left + 1) > MAX_RANGE_CELLS) return "range"

		// What it would read were the sheet to reach its last row or column.
		const cols = Math.abs(area.endCol - area.startCol) + 1
		const rows = Math.abs(area.endRow - area.startRow) + 1

		if (area.kind === "cols" || (area.kind === "cell" && rows * cols > MAX_RANGE_CELLS)) reach?.(first.name, cols, 0)
		if (area.kind === "rows" || (area.kind === "cell" && rows * cols > MAX_RANGE_CELLS)) reach?.(first.name, 0, rows)
	}

	return null
}

// What a formula the engine is not given shows when its file stored no result.
function unreadableShows(reason: "sheet" | "range"): RawCellContent {
	return reason === "sheet" ? REF_ERROR : UNPARSEABLE
}

// A value the engine gave, as content holding it: text kept as text, an error as its text.
function rawValue(value: unknown): RawCellContent | undefined {
	if (value instanceof DetailedCellError) return exported(value) ?? undefined
	if (typeof value === "string") return `'${value}`

	return typeof value === "number" || typeof value === "boolean" ? value : undefined
}

// A held result as the engine's functions give it: text without the apostrophe that keeps it text, an error
// text as the error.
function heldInterpreterValue(value: RawCellContent): number | string | boolean | CellError {
	if (typeof value === "number" || typeof value === "boolean") return value
	if (typeof value !== "string") return 0
	if (value.startsWith("'")) return value.slice(1)

	const type = ERROR_TYPES.get(value)

	return type === undefined ? value : new CellError(type)
}

export class FormulaEngine {
	private engine: HyperFormula
	// Grid sheet index (worksheets only, as the grid lists them) → engine sheet id.
	private readonly sheetIds: number[] = []
	// Held results by id (FILEN.HELD's first argument), as the engine's functions give them.
	private readonly heldValues: (number | string | boolean | CellError)[] = []
	private readonly held: ((sheet: number, row: number, col: number) => void) | undefined
	private readonly config: Partial<ConfigParams>
	// Per sheet (upper-case name), the widest span of columns a formula reads to the sheet's last row, and of rows it
	// reads to its last column: what a cell written past the sheet's reach makes the engine read.
	private readonly reaches = new Map<string, { cols: number; rows: number }>()
	private readonly reach = (sheet: string, cols: number, rows: number): void => {
		const key = sheet.toUpperCase()
		const known = this.reaches.get(key)

		this.reaches.set(key, { cols: Math.max(known?.cols ?? 0, cols), rows: Math.max(known?.rows ?? 0, rows) })
	}
	// The held results recalculated while an edit is recorded (their inputs changed), by changeKey.
	private recorded: Map<string, SimpleCellAddress> | null = null

	// `held`: told of each cell this engine comes to hold on its own account (a formula it fails on), by grid sheet
	// index; with it, a formula in a circle the engine finds, or naming a sheet the workbook lacks, is held too (the
	// caller tracks what held results read).
	constructor(
		sheets: readonly EngineSheet[],
		names: readonly EngineName[],
		date1904 = false,
		held?: (sheet: number, row: number, col: number) => void
	) {
		this.held = held
		this.config = engineConfig(date1904)
		registerHeld()

		// A null-prototype record: a sheet named "__proto__" is a key like any other.
		const contents = Object.create(null) as Record<string, RawCellContent[][]>

		const extents: SheetExtent[] = sheets.map(sheet => ({
			name: sheet.name,
			height: sheet.rows.length,
			width: sheet.rows.reduce((width, row) => Math.max(width, row.length), 0)
		}))
		const missing: { sheet: number; row: number; col: number }[] = []

		sheets.forEach((sheet, index) => {
			const fallbacks = new Map<number, RawCellContent>()

			for (const cell of sheet.formulas) {
				const row = sheet.rows[cell.row]

				if (row !== undefined && cell.content instanceof HeldResult) row[cell.col] = this.holding(cell.content)
				if (cell.fallback !== undefined) fallbacks.set(cell.row * MAX_COLUMNS + cell.col, cell.fallback)
			}

			if (held !== undefined) {
				sheet.rows.forEach((row, rowIndex) => {
					row.forEach((value, col) => {
						const reason = unreadable(value, extents, index, this.reach)

						if (reason === null) return

						row[col] = this.holding(new HeldResult(fallbacks.get(rowIndex * MAX_COLUMNS + col) ?? unreadableShows(reason)))
						missing.push({ sheet: index, row: rowIndex, col })
					})
				})
			}

			contents[sheet.name] = sheet.rows as RawCellContent[][]
		})

		for (const cell of missing) this.held?.(cell.sheet, cell.row, cell.col)

		// The names come with the build, so it calculates once. A name the engine cannot take (a print area it has
		// no syntax for) is left out: what uses it keeps the file's results (the caller holds them,
		// FormulaEngine.refusedNames). The engine numbers the sheets in the record's key order.
		const order = Object.keys(contents)
		const refused = FormulaEngine.refusedNames(
			sheets.map(sheet => sheet.name),
			names
		)
		const namedExpressions = names.flatMap((named, index) => {
			const sheet = named.scope === undefined ? undefined : sheets[named.scope]
			const scope = sheet === undefined ? undefined : order.indexOf(sheet.name)

			return refused.has(index) || scope === -1
				? []
				: [{ name: named.name, expression: `=${named.expression}`, ...(scope !== undefined && { scope }) }]
		})

		const build = (): HyperFormula => this.run(() => HyperFormula.buildFromSheets(contents, this.config, namedExpressions))

		this.engine = build()

		for (const sheet of sheets) {
			const id = this.engine.getSheetId(sheet.name)

			if (id === undefined) {
				throw new Error(`spreadsheet: engine has no sheet ${sheet.name}`)
			}

			this.sheetIds.push(id)
		}

		let pending = sheets.flatMap((sheet, index) => sheet.formulas.flatMap(cell => this.pending(index, cell)))

		// Many formulas the engine fails on (a circle through thousands of cells): built again with them held, which
		// costs less than replacing each one (every cell a held formula read lets go of its dependencies).
		const failing = pending.flatMap(cell => {
			const held = this.failure(cell)

			return held === null ? [] : [{ cell, held }]
		})

		if (failing.length >= REBUILD_FAILING) {
			for (const { cell, held } of failing) {
				const row = contents[order[cell.address.sheet] ?? ""]?.[cell.address.row]

				if (row !== undefined) row[cell.address.col] = this.holding(held)

				this.heldBy(cell.address)
			}

			this.engine.destroy()
			this.engine = build()

			const failed = new Set(failing.map(({ cell }) => cell))

			pending = pending.filter(cell => !failed.has(cell))
		}

		this.settle(pending, new Map())
	}

	// The names the engine cannot give a defined name, by index (it reads CustomCost1 or R2D2 as a reference, which
	// Excel does not): such a name is given to it under an alias.
	static unnamable(names: readonly string[]): Set<number> {
		const unnamable = new Set<number>()

		if (names.length === 0) return unnamable

		const engine = HyperFormula.buildEmpty({ licenseKey: LICENSE_KEY })

		try {
			names.forEach((name, index) => {
				if (!engine.isItPossibleToAddNamedExpression(name, "=1")) unnamable.add(index)
			})
		} finally {
			engine.destroy()
		}

		return unnamable
	}

	// The defined names the engine cannot take, by index: syntax it has no reading for (it would take some as
	// names whose value is a parse error). `sheets` are the grid's worksheet names, which a name's scope indexes.
	static refusedNames(sheets: readonly string[], names: readonly EngineName[]): Set<number> {
		const refused = new Set<number>()

		if (names.length === 0) return refused

		const contents = Object.create(null) as Record<string, RawCellContent[][]>

		for (const name of sheets) contents[name] = []

		const engine = HyperFormula.buildFromSheets(contents, { licenseKey: LICENSE_KEY, maxRows: MAX_ROWS, maxColumns: MAX_COLUMNS })

		try {
			names.forEach((named, index) => {
				const sheet = named.scope === undefined ? undefined : sheets[named.scope]
				const scope = sheet === undefined ? undefined : engine.getSheetId(sheet)

				const expression = `=${named.expression}`

				if (!engine.validateFormula(expression) || !engine.isItPossibleToAddNamedExpression(named.name, expression, scope))
					refused.add(index)
			})
		} finally {
			engine.destroy()
		}

		return refused
	}

	// Runs work that may calculate: the held-result functions answer from this engine meanwhile.
	private run<T>(work: () => T): T {
		const previous = running

		running = this.answer

		try {
			return work()
		} finally {
			running = previous
		}
	}

	// A held result as the formula the engine is given.
	private holding(held: HeldResult): string {
		const id = this.heldValues.length

		this.heldValues.push(heldInterpreterValue(held.value))

		return `=${HELD}(${String(id)})`
	}

	private contentOf(content: RawCellContent | HeldResult): RawCellContent {
		return content instanceof HeldResult ? this.holding(content) : content
	}

	// What FILEN.HELD gives in the cell at `address`, recorded as recalculated while an edit is.
	private readonly answer: HeldAnswer = (id, address) => {
		this.recorded?.set(changeKey(address), address)

		return this.heldValues[id] ?? new CellError(ErrorType.VALUE)
	}

	private heldBy(address: SimpleCellAddress): void {
		const sheet = this.sheetIndex(address.sheet)

		if (sheet >= 0) this.held?.(sheet, address.row, address.col)
	}

	private pending(sheet: number, cell: EngineCell): Pending[] {
		const id = this.sheetIds[sheet]
		return id === undefined || cell.fallback === undefined || cell.content instanceof HeldResult
			? []
			: [{ address: { sheet: id, row: cell.row, col: cell.col }, content: cell.content, fallback: cell.fallback }]
	}

	// What a cell the engine fails on holds instead: its stored result. Null when the engine does not fail on it.
	private failure(cell: Pending): HeldResult | null {
		const value = this.engine.getCellValue(cell.address)
		const fallback = cell.fallback

		if (!(value instanceof DetailedCellError) || fallback === undefined) {
			return null
		}

		if (value.type === ErrorType.ERROR) {
			return typeof cell.content !== "string" || !this.engine.validateFormula(cell.content) ? new HeldResult(fallback) : null
		}

		// A legacy array formula has no room to spill over the values the file keeps in its range (the
		// formula's own spill has no origin address; one it inherits has its own).
		if (value.type === ErrorType.SPILL) {
			return value.address === undefined ? new HeldResult(fallback) : null
		}

		// An unknown function or name in this very formula, not one it inherits; or a circle it is part of, which the
		// engine draws from every cell a range names (INDEX(A1:C1,1,1) in B1) where Excel's calculation may never
		// meet one. Held, the circle is gone.
		return (value.type === ErrorType.NAME || (value.type === ErrorType.CYCLE && this.held !== undefined)) &&
			value.address === this.engine.simpleCellAddressToString(cell.address, -1)
			? new HeldResult(fallback)
			: null
	}

	// Replaces formulas the engine fails on with their held results, recording what that recalculates. A second
	// round catches formulas whose own failure an inherited one hid.
	private settle(cells: Pending[], changes: Map<string, RecalculatedCell>): void {
		let pending = cells

		for (let round = 0; round < 3 && pending.length > 0; round++) {
			const failing: { cell: Pending; held: HeldResult }[] = []

			for (const cell of pending) {
				const held = this.failure(cell)

				if (held !== null) failing.push({ cell, held })
			}

			if (failing.length === 0) {
				return
			}

			this.collect(
				this.run(() =>
					this.engine.batch(() => {
						for (const { cell, held } of failing) {
							this.engine.setCellContents(cell.address, this.holding(held))
						}
					})
				),
				changes
			)

			for (const { cell } of failing) {
				const key = changeKey(cell.address)

				// The result the file stored did not change.
				changes.delete(key)
				this.recorded?.delete(key)
				this.heldBy(cell.address)
			}

			const failed = new Set(failing.map(({ cell }) => cell))

			pending = pending.filter(cell => !failed.has(cell))
		}
	}

	private sheetIndex(id: number): number {
		return this.sheetIds.indexOf(id)
	}

	// Each cell a change names, with its value read back once the operation is done: what a batch reports is not
	// always what the cell holds (a formula giving a range, =A1:A3, is reported as spilling its values over the
	// cells below, while it holds #VALUE! and they keep their own).
	private collect(exportedChanges: ReturnType<HyperFormula["batch"]>, into: Map<string, RecalculatedCell>): void {
		for (const change of exportedChanges) {
			if (!("address" in change)) {
				continue
			}

			const sheet = this.sheetIndex(change.address.sheet)

			if (sheet < 0) {
				continue
			}

			const value: unknown = this.engine.getCellValue(change.address)

			into.set(changeKey(change.address), {
				sheet,
				row: change.address.row,
				col: change.address.col,
				value: exported(value),
				...(value instanceof DetailedCellError ? { error: true } : {})
			})
		}
	}

	// Runs an edit, reporting what it recalculates: the held results whose inputs changed among it (`own`, the
	// edit's own cells, aside).
	private recording(work: () => ReturnType<HyperFormula["batch"]>, own: ReadonlySet<string> | null): Map<string, RecalculatedCell> {
		const changes = new Map<string, RecalculatedCell>()
		const recorded = new Map<string, SimpleCellAddress>()

		this.recorded = recorded

		try {
			this.collect(this.run(work), changes)
		} finally {
			this.recorded = null
		}

		this.staled(recorded, own, changes)

		return changes
	}

	private staled(
		recorded: ReadonlyMap<string, SimpleCellAddress>,
		own: ReadonlySet<string> | null,
		changes: Map<string, RecalculatedCell>
	): void {
		for (const [key, address] of recorded) {
			const sheet = this.sheetIndex(address.sheet)

			if (sheet < 0 || own?.has(key) === true || changes.has(key)) continue

			const value: unknown = this.engine.getCellValue(address)

			changes.set(key, {
				sheet,
				row: address.row,
				col: address.col,
				value: exported(value),
				...(value instanceof DetailedCellError ? { error: true } : {}),
				stale: true
			})
		}
	}

	// Sets cells and returns every cell whose value changed as a result, the set ones included (except
	// formulas that fell back to their stored result, which did not change), and the held results it made stale.
	set(sheet: number, cells: readonly EngineCell[]): RecalculatedCell[] {
		const id = this.sheetIds[sheet]

		if (id === undefined || cells.length === 0) {
			return []
		}

		if (this.held !== undefined && this.outgrows(id, cells)) return [...this.rebuildWith(sheet, cells).values()]

		const own = new Set(cells.map(cell => changeKey({ sheet: id, row: cell.row, col: cell.col })))
		// The sheets as far as the engine has seen them, read once, when a formula needs them.
		let known: SheetExtent[] | null = null
		const extents = (): SheetExtent[] =>
			(known ??= this.sheetIds.map(sheetId => {
				const { width, height } = this.engine.getSheetDimensions(sheetId)

				return { name: this.engine.getSheetName(sheetId) ?? "", height, width }
			}))
		const changes = this.recording(
			() =>
				this.engine.batch(() => {
					for (const cell of cells) {
						const address = { sheet: id, row: cell.row, col: cell.col }

						try {
							const reason = this.held === undefined ? null : unreadable(cell.content, extents(), sheet, this.reach)

							if (reason !== null) {
								this.engine.setCellContents(address, this.holding(new HeldResult(cell.fallback ?? unreadableShows(reason))))
								this.heldBy(address)
							} else {
								this.engine.setCellContents(address, this.contentOf(cell.content))
							}
						} catch {
							// A formula the parser overflows on: it holds what it stored, or reads as unparseable.
							const stored = cell.fallback

							this.engine.setCellContents(address, stored === undefined ? UNPARSEABLE : this.holding(new HeldResult(stored)))

							if (stored !== undefined) this.heldBy(address)
						}
					}
				}),
			own
		)

		this.settle(
			cells.flatMap(cell => this.pending(sheet, cell)),
			changes
		)

		// A set cell whose value is what it was is not reported by the engine, though the caller wrote it anew.
		for (const cell of cells) {
			const address = { sheet: id, row: cell.row, col: cell.col }
			const key = changeKey(address)

			if (changes.has(key)) continue

			const value: unknown = this.engine.getCellValue(address)

			changes.set(key, {
				sheet,
				row: cell.row,
				col: cell.col,
				value: exported(value),
				...(value instanceof DetailedCellError ? { error: true } : {})
			})
		}

		return [...changes.values()]
	}

	// Whether a result is the engine failing rather than the formula (a function it lacks, syntax it cannot
	// read, an array it spills where Excel does not): the file's own result is kept.
	static unsupported(value: unknown): boolean {
		return value === "#NAME?" || value === PARSE_ERROR || value === "#SPILL!" || value === CYCLE
	}

	// Whether the engine can insert or delete these rows or columns (it cannot move a spilled array onto
	// cells it would cover, nor grow a sheet past its size), or why not.
	moveRefusal(sheet: number, edit: AxisEdit): "spill" | "tooLarge" | null {
		const id = this.sheetIds[sheet]

		if (id === undefined) {
			return null
		}

		const span: [number, number] = [edit.at, edit.count]

		if (edit.type === "insert") {
			return (edit.axis === "rows" ? this.engine.isItPossibleToAddRows(id, span) : this.engine.isItPossibleToAddColumns(id, span))
				? null
				: "spill"
		}

		if (this.removedCopies(id, edit.axis, edit.count) > MAX_COPIED_CELLS) return "tooLarge"

		return (edit.axis === "rows" ? this.engine.isItPossibleToRemoveRows(id, span) : this.engine.isItPossibleToRemoveColumns(id, span))
			? null
			: "spill"
	}

	canMove(sheet: number, edit: AxisEdit): boolean {
		return this.moveRefusal(sheet, edit) === null
	}

	// The cells the engine copies to remove rows or columns of a sheet: each one across the sheet's whole width or
	// height as it has seen it.
	private removedCopies(id: number, axis: SizeAxis, count: number): number {
		const { width, height } = this.engine.getSheetDimensions(id)

		return count * (axis === "rows" ? width : height)
	}

	private ensureCopies(cells: number): void {
		if (cells > MAX_COPIED_CELLS) throw new EngineLimitError()
	}

	insert(sheet: number, axis: SizeAxis, at: number, count: number): RecalculatedCell[] {
		const id = this.sheetIds[sheet]

		if (id === undefined) return []

		// In runs: the engine spreads the new rows into one call, which a large insert overflows.
		return [
			...this.recording(
				() =>
					this.engine.batch(() => {
						for (let done = 0; done < count; done += INSERT_RUN) {
							const run = Math.min(INSERT_RUN, count - done)

							if (axis === "rows") this.engine.addRows(id, [at, run])
							else this.engine.addColumns(id, [at, run])
						}
					}),
				null
			).values()
		]
	}

	remove(sheet: number, axis: SizeAxis, at: number, count: number): RecalculatedCell[] {
		const id = this.sheetIds[sheet]

		if (id === undefined) return []

		this.ensureCopies(this.removedCopies(id, axis, count))

		return [
			...this.recording(
				() => (axis === "rows" ? this.engine.removeRows(id, [at, count]) : this.engine.removeColumns(id, [at, count])),
				null
			).values()
		]
	}

	addSheet(name: string): void {
		const added = this.run(() => this.engine.addSheet(name))
		const id = this.engine.getSheetId(added)

		if (id === undefined) {
			throw new Error(`spreadsheet: engine has no sheet ${name}`)
		}

		this.sheetIds.push(id)
	}

	// Removes the last sheet (an added one being undone).
	removeLastSheet(): RecalculatedCell[] {
		const id = this.sheetIds.at(-1)

		if (id === undefined) return []

		const { width, height } = this.engine.getSheetDimensions(id)

		this.ensureCopies(width * height)
		this.sheetIds.pop()

		return [...this.recording(() => this.engine.removeSheet(id), null).values()]
	}

	renameSheet(sheet: number, name: string): void {
		const id = this.sheetIds[sheet]

		if (id !== undefined) {
			const before = (this.engine.getSheetName(id) ?? "").toUpperCase()
			const reach = this.reaches.get(before)

			this.run(() => {
				this.engine.renameSheet(id, name)
			})

			if (reach !== undefined) {
				this.reaches.delete(before)
				this.reaches.set(name.toUpperCase(), reach)
			}
		}
	}

	// Whether writing these cells makes the sheet reach so far that a formula reading it to its last row or column
	// would read past MAX_RANGE_CELLS.
	private outgrows(id: number, cells: readonly EngineCell[]): boolean {
		const reach = this.reaches.get((this.engine.getSheetName(id) ?? "").toUpperCase())

		if (reach === undefined) return false

		const { width, height } = this.engine.getSheetDimensions(id)
		let rows = height
		let cols = width

		for (const cell of cells) {
			if (cell.content === null || cell.content === undefined || cell.content === "") continue

			rows = Math.max(rows, cell.row + 1)
			cols = Math.max(cols, cell.col + 1)
		}

		return (rows > height && reach.cols * rows > MAX_RANGE_CELLS) || (cols > width && reach.rows * cols > MAX_RANGE_CELLS)
	}

	// Builds the engine again with these cells written, holding at the results they had the formulas that read past
	// MAX_RANGE_CELLS once the sheet reaches as far, and reports every formula cell whose value that changed and the
	// cells written. Rare (a cell written at the grid's far edge) and whole-workbook: the engine would otherwise read
	// each of those ranges cell by cell, past what memory holds.
	private rebuildWith(sheet: number, cells: readonly EngineCell[]): Map<string, RecalculatedCell> {
		const old = this.engine
		const names = this.sheetIds.map(id => old.getSheetName(id) ?? "")
		const contents = Object.create(null) as Record<string, RawCellContent[][]>

		this.sheetIds.forEach((id, index) => {
			contents[names[index] ?? ""] = old.getSheetSerialized(id)
		})

		const written = contents[names[sheet] ?? ""] ?? []

		for (const cell of cells) {
			while (written.length <= cell.row) written.push([])

			const row = written[cell.row] ?? []

			while (row.length < cell.col) row.push(null)

			row[cell.col] = this.contentOf(cell.content)
		}

		const extents: SheetExtent[] = names.map(name => {
			const rows = contents[name] ?? []

			return { name, height: rows.length, width: rows.reduce((width, row) => Math.max(width, row.length), 0) }
		})
		const fallbacks = new Map(cells.map(cell => [cell.row * MAX_COLUMNS + cell.col, cell.fallback]))
		const before = new Map<string, { address: SimpleCellAddress; value: unknown }>()
		const holds: SimpleCellAddress[] = []

		this.reaches.clear()

		names.forEach((name, index) => {
			const id = this.sheetIds[index] ?? 0

			;(contents[name] ?? []).forEach((row, rowIndex) => {
				row.forEach((value, col) => {
					if (typeof value !== "string" || !value.startsWith("=")) return

					const address = { sheet: id, row: rowIndex, col }
					const own = index === sheet && fallbacks.has(rowIndex * MAX_COLUMNS + col)
					const was: unknown = own ? undefined : old.getCellValue(address)
					const reason = unreadable(value, extents, index, this.reach)

					before.set(changeKey(address), { address, value: was })

					if (reason === null) return

					const stored = own ? fallbacks.get(rowIndex * MAX_COLUMNS + col) : rawValue(was)

					row[col] = this.holding(new HeldResult(stored ?? unreadableShows(reason)))
					holds.push(address)
				})
			})
		})

		const order = Object.keys(contents)
		const named = old.getAllNamedExpressionsSerialized().map(entry => {
			const scope = entry.scope === undefined ? undefined : order.indexOf(old.getSheetName(entry.scope) ?? "")

			const { scope: _scope, ...global } = entry

			return scope === undefined || scope < 0 ? global : { ...global, scope }
		})

		old.destroy()
		this.engine = this.run(() => HyperFormula.buildFromSheets(contents, this.config, named))

		const ids = names.map(name => this.engine.getSheetId(name) ?? -1)
		const remap = (address: SimpleCellAddress): SimpleCellAddress => ({
			...address,
			sheet: ids[this.sheetIds.indexOf(address.sheet)] ?? address.sheet
		})

		for (const address of holds) this.heldBy(remap(address))

		const changes = new Map<string, RecalculatedCell>()
		const report = (address: SimpleCellAddress, was: unknown, always: boolean): void => {
			const value: unknown = this.engine.getCellValue(address)

			if (!always && exported(value) === exported(was) && value instanceof DetailedCellError === was instanceof DetailedCellError)
				return

			changes.set(changeKey(address), {
				sheet: ids.indexOf(address.sheet),
				row: address.row,
				col: address.col,
				value: exported(value),
				...(value instanceof DetailedCellError ? { error: true } : {})
			})
		}

		const previous = [...before.values()].map(entry => ({ address: remap(entry.address), value: entry.value }))

		this.sheetIds.splice(0, this.sheetIds.length, ...ids)

		for (const entry of previous) report(entry.address, entry.value, false)
		for (const cell of cells) report({ sheet: ids[sheet] ?? 0, row: cell.row, col: cell.col }, undefined, true)

		return changes
	}

	// A cell's value, and whether it is an error.
	result(sheet: number, row: number, col: number): Pick<RecalculatedCell, "value" | "error"> {
		const id = this.sheetIds[sheet]
		const value: unknown = id === undefined ? null : this.engine.getCellValue({ sheet: id, row, col })

		return value instanceof DetailedCellError ? { value: exported(value), error: true } : { value: exported(value) }
	}

	destroy(): void {
		this.engine.destroy()
	}
}
