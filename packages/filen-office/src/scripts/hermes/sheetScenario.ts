import {
	cellKey,
	csvCellView,
	engineFormula,
	formulaTranslator,
	malformedFormula,
	parseCellInput,
	parseCsvFile,
	parseTsv,
	renameSheetInFormula,
	serializeCsv,
	shiftFormula,
	shownFormula,
	storedFormula,
	wellFormedFormula
} from "../../core/sheet/index"
import { FormulaEngine } from "../../core/sheet/formula/index"
import { engineCellFormula, engineSpillFormula } from "../../core/sheet/refs/formulaRefs"
import type { SheetCommand, SheetFileKind, SheetSessionApi, SheetTile } from "../../core/sheet/protocol"
import type { Slicing } from "../../core/slicing"
import { inThreadSheetHost } from "../../dom/sheet/host"

// The spreadsheet core's text paths (references, CSV reading and writing in every encoding, typed input) and
// a HyperFormula workbook, run alike under V8 and Hermes: each probe's value is compared between them.

function hex(bytes: Uint8Array): string {
	let out = ""

	for (const byte of bytes) out += byte.toString(16).padStart(2, "0")

	return out
}

function bytes(...values: number[]): Uint8Array {
	return Uint8Array.from(values)
}

const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text)

function utf16(text: string, littleEndian: boolean): Uint8Array {
	const out = new Uint8Array(2 + text.length * 2)

	out[0] = littleEndian ? 0xff : 0xfe
	out[1] = littleEndian ? 0xfe : 0xff

	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i)

		out[2 + i * 2] = littleEndian ? code & 0xff : code >> 8
		out[3 + i * 2] = littleEndian ? code >> 8 : code & 0xff
	}

	return out
}

const FORMULAS = [
	"SUM(A1:B2)+$C$3*Sheet2!D4",
	"'My Sheet'!A1+[1]Ext!B2+Jan:Mar!C3",
	'VLOOKUP(A:A,1:3,2,FALSE)&"A1"',
	"Table1[[#This Row],[Col]]*LOG10(A1)",
	"IF(A1>0,SUM(A$1:A1),-1",
	"A1;B1",
	"_xlfn.XLOOKUP(A1,B:B,C:C)",
	"1E+3+1E3+TRUE"
]

const CSV_FILES: readonly [string, Uint8Array, boolean][] = [
	["utf-8", utf8('id,name\r\n1,"Größe, ""x"""\r\n2,😀\r\n'), false],
	["utf-8 bom", bytes(0xef, 0xbb, 0xbf, ...utf8("a;b;c\n1;2,5;3\n")), false],
	["utf-8 double bom", bytes(0xef, 0xbb, 0xbf, 0xef, 0xbb, 0xbf, 0x61, 0x0a), false],
	["utf-16le", utf16("a\tb\r\nx€y\t😀\r\n", true), false],
	["utf-16be", utf16("a,b\n1,2", false), false],
	["utf-16le lone and odd", bytes(0xff, 0xfe, 0x61, 0x00, 0x2c, 0x00, 0x00, 0xd8, 0x0a, 0x00, 0x41), false],
	["windows-1252", bytes(0x63, 0x61, 0x66, 0xe9, 0x2c, 0x80, 0x0d, 0x0a, 0x93, 0x71, 0x94, 0x2c, 0x78, 0x0d, 0x0a), false],
	["windows-1252 undefined", bytes(0x63, 0x61, 0x81, 0x2c, 0x62), false],
	["gbk", bytes(0xc3, 0xfb, 0xb3, 0xc6, 0x2c, 0xb3, 0xc7, 0xca, 0xd0, 0x0a), false],
	["invalid utf-8", bytes(0x61, 0x2c, 0xed, 0xa0, 0x80, 0x0a), false],
	["quotes", utf8('ab"c,d"e"f,"g"\n"open\nrest'), false],
	["tsv", utf8('a\t"x\ty"\t"l1\nl2"\n'), true],
	["cr only", utf8("a,b\r1,2\r"), false],
	["empty", bytes(), false]
]

const INPUTS = ["42", "1,234.5", "12%", "-.5e3", "2024-02-29", "2023-02-29", "true", "'007", "=SUM(A1:A3)", "=SUM(A1", " text "]

function csvProbe(file: Uint8Array, tsv: boolean): string {
	const { rows, format } = parseCsvFile(file, tsv)
	const cells = rows.flatMap((row, rowIndex) =>
		row.flatMap((text, colIndex) => {
			const view = csvCellView(text)

			return view === null ? [] : [`${String(cellKey(rowIndex, colIndex))}:${view.text}${view.numeric === true ? "#" : ""}`]
		})
	)
	let saved: string

	try {
		saved = hex(serializeCsv(rows, format))
	} catch (error) {
		saved = `threw ${error instanceof Error ? error.message : String(error)}`
	}

	return JSON.stringify({ rows, format, cells, saved })
}

function engineProbe(): string {
	const engine = new FormulaEngine(
		[
			{
				name: "Data",
				rows: [
					[1, 2, "=A1+B1"],
					["=DATE(2024,2,29)", '=TEXT(A2,"yyyy-mm-dd")', "=1/0"]
				],
				formulas: []
			},
			{
				name: "Sum",
				rows: [["=SUM(Data!A1:C1)*Rate", "=NOSUCH(1)", "=VLOOKUP(2,Data!A1:C1,1,FALSE)"]],
				formulas: [{ row: 0, col: 1, content: "=NOSUCH(1)", fallback: 7 }]
			}
		],
		[{ name: "Rate", expression: "2" }]
	)

	try {
		const changed = engine
			.set(0, [{ row: 0, col: 0, content: 10 }])
			.map(cell => `${String(cell.sheet)}:${String(cell.row)},${String(cell.col)}=${String(cell.value)}`)
			.sort()

		engine.insert(0, "rows", 0, 1)

		const after = [
			engine.result(0, 1, 2),
			engine.result(0, 2, 0),
			engine.result(0, 2, 1),
			engine.result(0, 2, 2),
			engine.result(1, 0, 0)
		]

		engine.renameSheet(0, "Renamed")

		return JSON.stringify({
			changed,
			after,
			renamed: engine.result(1, 0, 0),
			fallback: engine.result(1, 0, 1),
			lookup: engine.result(1, 0, 2)
		})
	} finally {
		engine.destroy()
	}
}

// The app's own functions where the engine has none: regular expressions, LET and LAMBDA, bonds, INDIRECT, and a
// dynamic array spilling and growing.
function functionsProbe(): string {
	const typed = [
		'REGEXEXTRACT("Order 123-456","\\d+-(\\d+)",2)',
		'REGEXREPLACE("a1b22c333","\\d+","#")',
		'REGEXTEST("Ünïcode","^\\w+$")',
		"LET(x,A1,LAMBDA(n,n*x)(3))",
		"SUM(MAP(A1:A3,LAMBDA(v,v^2)))",
		"REDUCE(0,A1:A3,LAMBDA(a,b,a+b))",
		"PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)",
		"YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)",
		'INDIRECT("A"&A1)',
		'SUM(INDIRECT("R1C1:R3C1",FALSE))'
	]
	const rows: (string | number | null)[][] = [[2], [3], [4], [], []]

	typed.forEach((formula, col) => {
		const row = rows[0] as (string | number | null)[]

		row[col + 1] = engineCellFormula(storedFormula(formula)) ?? ""
	})

	const spill = engineSpillFormula("_xlfn.SEQUENCE(A1)") ?? ""

	;(rows[4] as (string | number | null)[])[0] = spill

	const engine = new FormulaEngine([{ name: "S", rows, formulas: [{ row: 4, col: 0, content: spill, spill: { rows: 1, cols: 1 } }] }], [])

	try {
		const values = typed.map((_, col) => engine.result(0, 0, col + 1))
		const grown = engine
			.set(0, [{ row: 0, col: 0, content: 3 }])
			.map(
				cell =>
					`${String(cell.row)},${String(cell.col)}=${String(cell.value)}${cell.spill === undefined ? "" : ` ${String(cell.spill.rows)}`}`
			)
			.sort()

		return JSON.stringify({ values, grown })
	} finally {
		engine.destroy()
	}
}

// LAMBDA recursion to just past its limit (1,024 calls one inside another): the same on every engine, whatever its
// stack.
function recursionProbe(): string {
	const lambdas = new Set(["SUMTO", "DOWN", "DEEP"])
	const aliases = new Map<string, string>()
	const names = [
		{ name: "SumTo", expression: engineFormula("_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=0,0,_xlpm.n+SumTo(_xlpm.n-1)))", aliases, lambdas) },
		{ name: "Down", expression: engineFormula("_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=0,0,Down(_xlpm.n-1)))", aliases, lambdas) },
		{
			name: "Deep",
			expression: engineFormula(
				"_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=0,0,ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(1+Deep(_xlpm.n-1)))))))))))))))))))",
				aliases,
				lambdas
			)
		}
	]
	const out: string[] = []

	for (const call of ["SumTo", "Down", "Deep"]) {
		const ns = [10, 500, 1023, 1024]
		const rows = [ns.map(n => engineCellFormula(`${call}(${String(n)})`, aliases, lambdas) ?? "")]
		const engine = new FormulaEngine([{ name: "S", rows, formulas: [] }], names)

		try {
			out.push(ns.map((n, col) => `${call}(${String(n)})=${String(engine.result(0, 0, col).value)}`).join(" "))
		} finally {
			engine.destroy()
		}
	}

	return out.join(" ")
}

export function sheetProbes(probe: (name: string, run: () => unknown) => void): void {
	probe("sheet.refs", () =>
		JSON.stringify(
			FORMULAS.map(formula => [
				wellFormedFormula(formula),
				shownFormula(formula),
				storedFormula(formula),
				engineFormula(formula),
				formulaTranslator(formula)(2, 1),
				shiftFormula(formula, "Sheet1", "Sheet1", { type: "insert", axis: "rows", at: 0, count: 2 }),
				shiftFormula(formula, "Sheet2", "Sheet2", { type: "delete", axis: "cols", at: 2, count: 1 }),
				renameSheetInFormula(formula, "Sheet2", "New 'one'")
			])
		)
	)

	for (const [name, file, tsv] of CSV_FILES) probe(`sheet.csv ${name}`, () => csvProbe(file, tsv))

	probe("sheet.tsv clipboard", () => JSON.stringify(parseTsv('a\t"b\nc"\td\r\n"x""y"\t\t')))
	probe("sheet.input", () =>
		JSON.stringify(
			INPUTS.map(input => {
				const parsed = parseCellInput(input)

				return [
					parsed.type === "value" && parsed.value instanceof Date ? { ...parsed, value: parsed.value.toISOString() } : parsed,
					malformedFormula(input)
				]
			})
		)
	)
	probe("sheet.hyperformula", engineProbe)
	probe("sheet.functions", functionsProbe)
	probe("sheet.recursion", recursionProbe)
}

// --- Sessions: a workbook or text file opened, edited and saved through the in-thread host ---

export interface SheetFixtureResult {
	readonly name: string
	readonly open: string
	// The save as opened: its digest, and whether it is the source byte for byte.
	readonly unedited: string
	readonly uneditedIdentical: boolean
	readonly tiles: string
	readonly verdict: string
	// What each edit returned: its patches' digest, or the refusal.
	readonly steps: readonly string[]
	readonly edited: string
	// The edited save opened again: its first tiles.
	readonly reopened: string
}

// Results do not depend on where the slices fall, so the clock may be the wall clock.
const SLICING: Slicing = {
	sliceMs: 8,
	now: () => Date.now(),
	pause: () =>
		new Promise<void>(resolve => {
			setTimeout(resolve, 0)
		})
}

// Two FNV-1a hashes and the length: equal digests are equal bytes, short enough to print.
export function bytesDigest(bytes: Uint8Array): string {
	let a = 0x811c9dc5
	let b = 0x01000193

	for (let i = 0; i < bytes.length; i++) {
		const byte = bytes[i] ?? 0

		a = Math.imul(a ^ byte, 0x01000193)
		b = Math.imul(b ^ byte, 0x0100019d) ^ (b >>> 15)
	}

	return `${String(bytes.length)}:${(a >>> 0).toString(16)}${(b >>> 0).toString(16)}`
}

function tileDigest(tile: SheetTile): string {
	const parts: string[] = [`${String(tile.tileRow)},${String(tile.tileCol)}`]

	for (let i = 0; i < tile.count; i++) {
		const text = tile.texts[i] ?? -1
		const input = tile.inputs[i] ?? -1

		parts.push(
			`${String(tile.positions[i])}=${text < 0 ? "" : (tile.strings[text] ?? "")}|${input < 0 ? "" : (tile.strings[input] ?? "")}|${String(tile.styles[i])}|${String(tile.flags[i])}`
		)
	}

	return parts.join(";")
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false

	for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false

	return true
}

export function sheetKind(name: string): SheetFileKind | null {
	const extension = name.slice(name.lastIndexOf(".") + 1).toLowerCase()

	switch (extension) {
		case "xlsx":
		case "xlsm":
		case "csv":
		case "tsv":
			return extension
		default:
			return null
	}
}

async function firstTiles(session: SheetSessionApi, sheet: number): Promise<string> {
	const { tiles } = await session.tiles(sheet, Uint32Array.of(0, 0, 0, 1, 1, 0))

	return tiles.map(tileDigest).join("\n")
}

function edits(sheet: number, text: boolean): SheetCommand[] {
	const set: SheetCommand = {
		type: "cells.set",
		sheet,
		cells: {
			rows: Uint32Array.of(0, 1, 2, 3, 4),
			cols: Uint16Array.of(0, 1, 2, 3, 0),
			inputs: ["3.14159", "2024-02-29", "=1/3", "50%", "Größe € 😀"]
		}
	}

	if (text)
		return [
			set,
			{ type: "axis.insert", sheet, axis: "rows", at: 1, count: 2 },
			{ type: "axis.delete", sheet, axis: "cols", at: 2, count: 1 }
		]

	return [
		set,
		{
			type: "format.apply",
			sheet,
			ranges: [{ startRow: 0, startCol: 0, endRow: 3, endCol: 3 }],
			patch: { numFmt: "#,##0.00;[Red]-#,##0.00", bold: true }
		},
		{ type: "axis.insert", sheet, axis: "rows", at: 1, count: 2 },
		{ type: "axis.delete", sheet, axis: "cols", at: 2, count: 1 },
		{ type: "axis.resize", sheet, axis: "cols", indices: Uint32Array.of(0), sizes: Float32Array.of(120) }
	]
}

// One file: opened, saved untouched, its first tiles, the save verdict, a run of edits with an undo and a redo,
// the edited save, and that save opened again.
export async function runSheetFixture(name: string, bytes: Uint8Array): Promise<SheetFixtureResult> {
	const kind = sheetKind(name) ?? "xlsx"
	const host = inThreadSheetHost()
	const empty = { unedited: "", uneditedIdentical: false, tiles: "", verdict: "", steps: [], edited: "", reopened: "" }

	try {
		const opened = await host.open(bytes.slice(), { kind, slicing: SLICING }, () => undefined)

		if (opened.kind !== "opened") return { name, open: opened.kind, ...empty }

		const { session, info } = opened
		const sheet = info.activeSheet
		const unedited = await session.save()
		const tiles = await firstTiles(session, sheet)
		const verdict = await session.intent()
		const steps: string[] = []
		let edited = ""
		let reopened = ""

		if (verdict.state === "writable") {
			for (const command of edits(sheet, kind === "csv" || kind === "tsv")) {
				const result = await session.command(command)

				steps.push(result.ok ? result.change.patches.map(tileDigest).join("/") : `refused ${result.refusal.reason}`)
			}

			steps.push((await session.undo()).ok ? "undo" : "undo refused")
			steps.push((await session.redo()).ok ? "redo" : "redo refused")

			const saved = await session.save()

			edited = saved.ok ? bytesDigest(saved.bytes) : `failed ${saved.reason}`

			if (saved.ok) {
				const again = await host.open(saved.bytes, { kind, slicing: SLICING }, () => undefined)

				if (again.kind === "opened") {
					reopened = await firstTiles(again.session, again.info.activeSheet)
					again.session.close()
				} else {
					reopened = again.kind
				}
			}
		}

		session.close()

		return {
			name,
			open: "opened",
			unedited: unedited.ok ? bytesDigest(unedited.bytes) : `failed ${unedited.reason}`,
			uneditedIdentical: unedited.ok && sameBytes(unedited.bytes, bytes),
			tiles,
			verdict: JSON.stringify(verdict),
			steps,
			edited,
			reopened
		}
	} finally {
		host.dispose()
	}
}
