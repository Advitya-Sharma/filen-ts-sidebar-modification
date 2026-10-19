import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, statSync } from "node:fs"
import { basename, dirname, join, relative } from "node:path"
import ts from "typescript"
import { formulaFunctions } from "./excelCached"
import { exportSheets, sheetProfile, type SheetOffice } from "./libreOffice"
import { suiteDir, suiteFiles, type SuiteName } from "./functionSuites"
import { formatA1, parseA1, readXlsxCells, type CellScalar, type XlsxCell, type XlsxSheet } from "./xlsxCells"

// The function suites (functionSuites.ts) as one case format for the formula engine's tests:
// { function, inputs, expected }. Generated at test time from the fetched files: formulajs's tests are parsed
// (never run) for expect(<fn>(<literals>)).to.<equal|eql|approximately>(<literal>); LibreOffice's .fods and
// POI's .xls workbooks are converted to xlsx by LibreOffice and read cell by cell, a case being a formula cell
// with its expected value.
//
// Sources and licences: LibreOffice sc/qa/unit/data/functions (MPL-2.0), formulajs test/ (MIT, Copyright (c) 2014
// Sutoiku, Inc.), Apache POI test-data/spreadsheet (Apache-2.0). The cases are derived from them at test time
// and never committed.

export type CaseValue =
	| CellScalar
	// An empty cell (formulajs's null).
	| { readonly type: "blank" }
	// An omitted argument (formulajs's undefined).
	| { readonly type: "missing" }
	| { readonly type: "array"; readonly rows: readonly (readonly CaseValue[])[] }

export type CaseInputs =
	// The function called with these arguments (an array argument stands for a range or array constant).
	| { readonly kind: "args"; readonly args: readonly CaseValue[] }
	// The formula in `cell` of the converted workbook (Excel syntax, as LibreOffice wrote it), with the data it
	// refers to in that workbook.
	| { readonly kind: "cell"; readonly workbook: string; readonly sheet: string; readonly cell: string; readonly formula: string }

export interface FunctionCase {
	readonly suite: SuiteName
	// Where the case comes from: "<file>:<line>" or "<file> <sheet>!<cell>".
	readonly id: string
	// Upper-case Excel name, dotted where Excel's is ("CHISQ.DIST.RT").
	readonly function: string
	readonly inputs: CaseInputs
	readonly expected: CaseValue
	// Absolute tolerance the source suite allows (chai approximately, POI's precision column).
	readonly tolerance?: number
	// The source's own description or check, when it has one.
	readonly note?: string
}

const ERRORS: Readonly<Record<string, string>> = {
	nil: "#NULL!",
	div0: "#DIV/0!",
	value: "#VALUE!",
	ref: "#REF!",
	name: "#NAME?",
	num: "#NUM!",
	na: "#N/A",
	data: "#GETTING_DATA",
	calc: "#CALC!"
}

// A literal argument or expectation; undefined when it is not a literal (a variable, a call, a Date).
function literal(node: ts.Expression): CaseValue | undefined {
	if (ts.isParenthesizedExpression(node)) return literal(node.expression)
	if (ts.isNumericLiteral(node)) return { type: "n", value: Number(node.text) }
	if (ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) return { type: "s", value: node.text }
	if (node.kind === ts.SyntaxKind.TrueKeyword) return { type: "b", value: true }
	if (node.kind === ts.SyntaxKind.FalseKeyword) return { type: "b", value: false }
	if (node.kind === ts.SyntaxKind.NullKeyword) return { type: "blank" }
	if (ts.isIdentifier(node) && node.text === "undefined") return { type: "missing" }

	if (ts.isPrefixUnaryExpression(node) && node.operator === ts.SyntaxKind.MinusToken) {
		const inner = literal(node.operand)

		return inner?.type === "n" ? { type: "n", value: -inner.value } : undefined
	}

	if (ts.isPropertyAccessExpression(node) && ts.isIdentifier(node.expression)) {
		if (node.expression.text === "error") {
			const text = ERRORS[node.name.text]

			return text === undefined ? undefined : { type: "e", value: text }
		}

		if (node.expression.text === "Math" && (node.name.text === "PI" || node.name.text === "E"))
			return { type: "n", value: node.name.text === "PI" ? Math.PI : Math.E }
	}

	if (ts.isArrayLiteralExpression(node)) {
		const items = node.elements.map(element => literal(element))

		if (items.some(item => item === undefined)) return undefined

		const values = items as CaseValue[]

		if (values.length > 0 && values.every(value => value.type === "array"))
			return { type: "array", rows: values.flatMap(value => (value.type === "array" ? value.rows : [])) }

		return values.some(value => value.type === "array") ? undefined : { type: "array", rows: [values] }
	}

	return undefined
}

// "statistical.CHISQ.DIST.RT" -> "CHISQ.DIST.RT": the property chain after the module namespace.
function calledFunction(node: ts.Expression): string | undefined {
	const parts: string[] = []
	let current = node

	while (ts.isPropertyAccessExpression(current)) {
		parts.unshift(current.name.text)
		current = current.expression
	}

	return ts.isIdentifier(current) && parts.length > 0 && /^[A-Z]/.test(parts[0] ?? "") ? parts.join(".") : undefined
}

export interface ParseTally {
	readonly cases: number
	// Left out: expect() assertions with non-literal arguments or expectations or other matchers; workbook rows
	// with a formula but no expected value.
	readonly skipped: number
}

// The literal cases of one formulajs test file.
export function formulajsCases(source: string, file: string): { cases: FunctionCase[]; tally: ParseTally } {
	const tree = ts.createSourceFile(file, source, ts.ScriptTarget.Latest, false, ts.ScriptKind.JS)
	const cases: FunctionCase[] = []
	let skipped = 0

	const visit = (node: ts.Node): void => {
		// expect(<call>).to[.deep|.be].<matcher>(<expected>[, <tolerance>])
		if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) {
			const matcher = node.expression.name.text
			let target = node.expression.expression

			while (ts.isPropertyAccessExpression(target) && ["to", "deep", "be"].includes(target.name.text)) target = target.expression

			if (ts.isCallExpression(target) && ts.isIdentifier(target.expression) && target.expression.text === "expect") {
				const call = target.arguments[0]
				const name = call !== undefined && ts.isCallExpression(call) ? calledFunction(call.expression) : undefined
				const args =
					name !== undefined && call !== undefined && ts.isCallExpression(call)
						? call.arguments.map(arg => literal(arg))
						: undefined
				const expectedNode = node.arguments[0]
				const expected = expectedNode === undefined ? undefined : literal(expectedNode)
				const toleranceNode = node.arguments[1]
				const tolerance = matcher === "approximately" && toleranceNode !== undefined ? literal(toleranceNode) : undefined
				const known =
					["equal", "eql", "approximately"].includes(matcher) && (matcher !== "approximately" || tolerance?.type === "n")

				if (
					name === undefined ||
					args === undefined ||
					args.some(arg => arg === undefined) ||
					expected === undefined ||
					expected.type === "missing" ||
					!known
				) {
					skipped++
				} else {
					const line = tree.getLineAndCharacterOfPosition(node.getStart(tree)).line + 1

					cases.push({
						suite: "formulajs",
						id: `${file}:${line}`,
						function: name.toUpperCase(),
						inputs: { kind: "args", args: args as CaseValue[] },
						expected,
						...(tolerance?.type === "n" && { tolerance: tolerance.value })
					})
				}

				return
			}
		}

		ts.forEachChild(node, visit)
	}

	visit(tree)

	return { cases, tally: { cases: cases.length, skipped } }
}

// The installed LibreOffice that converts the workbook suites: the newest (26.8 converts the .fods; its bundled
// Python is not needed).
export function conversionOffice(installs: readonly SheetOffice[]): SheetOffice | undefined {
	const key = (office: SheetOffice): number[] => office.version.split(".").map(Number)

	return [...installs].sort((a, b) => {
		const [x, y] = [key(a), key(b)]

		for (let index = 0; index < Math.max(x.length, y.length); index++)
			if ((x[index] ?? 0) !== (y[index] ?? 0)) return (y[index] ?? 0) - (x[index] ?? 0)

		return 0
	})[0]
}

const BATCH = 32
const TIMEOUT_MS = 120_000

// The suite's workbooks converted to xlsx by `office` (kept in the cache next to the suite and reused while newer
// than their source), as [source path relative to the suite, xlsx path]; failures listed by source path.
export async function convertedWorkbooks(
	suite: "libreoffice" | "poi",
	office: SheetOffice,
	limit = Number.POSITIVE_INFINITY
): Promise<{ workbooks: [string, string][]; failures: [string, string][] }> {
	const files = join(suiteDir(suite), "files")
	const out = join(suiteDir(suite), `xlsx-${office.name}`)
	const sources = suiteFiles(suite).slice(0, limit)
	const target = (source: string): string => join(out, `${relative(files, source)}.xlsx`)
	const stale = sources.filter(source => !existsSync(target(source)) || statSync(target(source)).mtimeMs < statSync(source).mtimeMs)
	const failures: [string, string][] = []
	const scratch = join(out, ".work")
	const profile = sheetProfile(scratch, `${office.name}-convert`)

	try {
		for (let start = 0; start < stale.length; start += BATCH) {
			const batch = stale.slice(start, start + BATCH)
			const outdir = join(scratch, String(start))
			const results = await exportSheets(office, profile, "xlsx", batch, outdir, TIMEOUT_MS)

			for (const result of results) {
				const output = result.outputs[0]

				if (!result.ok || output === undefined) {
					failures.push([relative(files, result.input), result.error ?? "no output"])
					continue
				}

				mkdirSync(dirname(target(result.input)), { recursive: true })
				renameSync(output, target(result.input))
			}

			rmSync(outdir, { recursive: true, force: true })
		}
	} finally {
		rmSync(scratch, { recursive: true, force: true })
	}

	return {
		workbooks: sources.filter(source => existsSync(target(source))).map(source => [relative(files, source), target(source)]),
		failures
	}
}

function cellAt(sheet: XlsxSheet): Map<string, XlsxCell> {
	return new Map(sheet.cells.map(cell => [`${cell.row},${cell.column}`, cell]))
}

function text(cell: XlsxCell | undefined): string {
	return cell?.value?.type === "s" ? cell.value.value.trim() : ""
}

// "chisq.dist.rt.fods" -> CHISQ.DIST.RT, unless the formula does not call it (a file named after a topic or a
// variant: then the formula's first function).
function functionOf(stem: string, formula: string): string {
	const named = (stem.split("_")[0] ?? "").toUpperCase()
	const called = formulaFunctions(formula)

	return called.includes(named) ? named : (called[0] ?? "(operators)")
}

// LibreOffice's function sheets: within the first three rows a header row with a column named "...expected..."; the
// formula columns run from the header "Function" or "Formula" nearest to its left (else the column before it) up
// to it, and each holds its expected value as many columns to the right. The usual sheet is
// "Function | Expected | Correct | FunctionString | Comment"; dynamic-array sheets put a spilled array under
// "Function" and its expected cells under "Expected", the formula an array formula whose ref gives the size.
// Rows with no expected value (checked through other columns, as the complex-number sheets do) count as skipped.
export async function libreOfficeCases(source: string, xlsx: string): Promise<{ cases: FunctionCase[]; tally: ParseTally }> {
	const workbook = await readXlsxCells(new Uint8Array(readFileSync(xlsx)))
	const stem = basename(source, ".fods")
	const cases: FunctionCase[] = []
	let skipped = 0

	for (const sheet of workbook.sheets) {
		const cells = cellAt(sheet)
		let header = -1
		let expectedColumn = -1

		for (let row = 0; row < 3 && header < 0; row++) {
			for (const cell of sheet.cells) {
				if (cell.row === row && cell.column > 0 && /\bexpected\b/i.test(text(cell)) && cell.formula === undefined) {
					header = row
					expectedColumn = cell.column
					break
				}
			}
		}

		if (header < 0) continue

		let firstColumn = expectedColumn - 1

		for (let column = expectedColumn - 1; column >= 0; column--) {
			if (/^(function|formula)$/i.test(text(cells.get(`${header},${column}`)))) {
				firstColumn = column
				break
			}
		}

		const width = expectedColumn - firstColumn
		const value = (row: number, column: number): CaseValue => cells.get(`${row},${column}`)?.value ?? { type: "blank" }

		for (const cell of sheet.cells) {
			const formula = cell.formula

			if (
				cell.row <= header ||
				cell.column < firstColumn ||
				cell.column >= expectedColumn ||
				formula === undefined ||
				formula.text === ""
			)
				continue

			const offset = expectedColumn + (cell.column - firstColumn)
			const range = formula.kind === "array" && formula.ref !== undefined ? formula.ref.split(":") : []
			const end = range.length === 2 ? parseA1(range[1] as string) : undefined
			let expected: CaseValue | undefined

			if (end !== undefined && (end.row > cell.row || end.column > cell.column)) {
				const rows: CaseValue[][] = []

				for (let row = cell.row; row <= end.row; row++) {
					const line: CaseValue[] = []

					for (let column = 0; column <= end.column - cell.column; column++) line.push(value(row, offset + column))

					rows.push(line)
				}

				expected = rows.flat().every(item => item.type === "blank") ? undefined : { type: "array", rows }
			} else {
				expected = cells.get(`${cell.row},${offset}`)?.value
			}

			if (expected === undefined) {
				skipped++
				continue
			}

			const check = cells.get(`${cell.row},${expectedColumn + width}`)?.formula?.text
			const ref = formatA1(cell.row, cell.column)

			cases.push({
				suite: "libreoffice",
				id: `${source} ${sheet.name}!${ref}`,
				function: functionOf(stem, formula.text),
				inputs: { kind: "cell", workbook: xlsx, sheet: sheet.name, cell: ref, formula: formula.text },
				expected,
				...(check !== undefined && { note: `check: ${check}` })
			})
		}
	}

	return { cases, tally: { cases: cases.length, skipped } }
}

// POI's BaseTestFunctionsFromSpreadsheet layout: a "Read Me" first sheet (A12: the 0-based column of each row's
// absolute precision, when numeric), then per sheet cases from row 5: column A a group marker ("<end>" stops,
// "<skip>" skips the row), B the formula, C the expected literal, D a comment.
export async function poiCases(source: string, xlsx: string): Promise<{ cases: FunctionCase[]; tally: ParseTally }> {
	const workbook = await readXlsxCells(new Uint8Array(readFileSync(xlsx)))
	const readMe = workbook.sheets[0]
	const precisionCell = readMe === undefined ? undefined : cellAt(readMe).get("11,0")?.value
	const precisionColumn = precisionCell?.type === "n" ? precisionCell.value : -1
	const cases: FunctionCase[] = []
	let skipped = 0

	for (const sheet of workbook.sheets.slice(1)) {
		const cells = cellAt(sheet)
		const last = sheet.cells.reduce((max, cell) => Math.max(max, cell.row), -1)
		let group = ""

		for (let row = 4; row <= last; row++) {
			const marker = text(cells.get(`${row},0`))

			if (marker.toLowerCase() === "<end>") break
			if (marker.toLowerCase() === "<skip>") continue
			if (marker !== "") group = marker

			const formula = cells.get(`${row},1`)?.formula
			const expected = cells.get(`${row},2`)?.value

			if (formula === undefined || formula.text === "") continue

			if (expected === undefined) {
				skipped++
				continue
			}

			const precision = precisionColumn >= 0 ? cells.get(`${row},${precisionColumn}`)?.value : undefined
			const comment = text(cells.get(`${row},3`))
			const ref = formatA1(row, 1)
			const called = formulaFunctions(formula.text)
			const named = group.toUpperCase()

			cases.push({
				suite: "poi",
				id: `${source} ${sheet.name}!${ref}`,
				function: called.includes(named) ? named : (called[0] ?? "(operators)"),
				inputs: { kind: "cell", workbook: xlsx, sheet: sheet.name, cell: ref, formula: formula.text },
				expected,
				...(precision?.type === "n" && precision.value > 0 && { tolerance: precision.value }),
				...((group !== "" || comment !== "") && { note: [group, comment].filter(Boolean).join(" - ") })
			})
		}
	}

	return { cases, tally: { cases: cases.length, skipped } }
}

// Every formulajs case of the fetched suite.
export function allFormulajsCases(): { cases: FunctionCase[]; tally: ParseTally } {
	const cases: FunctionCase[] = []
	let skipped = 0

	for (const path of suiteFiles("formulajs")) {
		const result = formulajsCases(readFileSync(path, "utf8"), `formulajs/${relative(join(suiteDir("formulajs"), "files"), path)}`)

		cases.push(...result.cases)
		skipped += result.tally.skipped
	}

	return { cases, tally: { cases: cases.length, skipped } }
}
