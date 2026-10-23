import type { SheetRange } from "../../sheet/protocol"
import { cellRef, parseRangeRef } from "../../xlsx/write/text"
import { decodeUtf8 } from "../../xml/index"

// Structured references (Table1[Col], Table1[[#Headers],[A]:[C]], [@Col]; [MS-XLSX] §2.2.2) as the cell ranges they
// name now, as LibreOffice reads them: OpenFormula has no tables. Null where a reference names what cannot be said
// (a table or column that is not there, a this-row reference outside the table's rows).

export interface SheetTable {
	readonly name: string
	readonly sheet: string
	readonly ref: SheetRange
	readonly headerRows: number
	readonly totalsRows: number
	readonly columns: readonly string[]
}

const ENTITIES: Readonly<Record<string, string>> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" }

function attribute(tag: string, name: string): string | undefined {
	const value = new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1]

	return value?.replace(/&(lt|gt|amp|quot|apos);/g, (_all, entity: string) => ENTITIES[entity] ?? "")
}

// A table part (ECMA-376 Part 1 §18.5.1.2): its name formulas use (displayName), range and columns.
export function readSheetTable(bytes: Uint8Array, sheet: string): SheetTable | null {
	const text = decodeUtf8(bytes)
	const root = /<(?:\w+:)?table\b[^>]*>/.exec(text)?.[0]

	if (root === undefined) return null

	const ref = parseRangeRef(attribute(root, "ref") ?? "")
	const name = attribute(root, "displayName") ?? attribute(root, "name")

	if (ref === null || ref === undefined || name === undefined) return null

	const columns = [...text.matchAll(/<(?:\w+:)?tableColumn\b[^>]*>/g)].map(match => attribute(match[0], "name") ?? "")

	return {
		name,
		sheet,
		ref,
		headerRows: Number(attribute(root, "headerRowCount") ?? "1"),
		totalsRows: Number(attribute(root, "totalsRowCount") ?? "0"),
		columns
	}
}

function quotedSheet(sheet: string): string {
	return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) && !/^[A-Za-z]{1,3}\d+$/.test(sheet) ? sheet : `'${sheet.replaceAll("'", "''")}'`
}

function absolute(row: number, col: number): string {
	const ref = cellRef(row, col)
	const split = /^([A-Z]+)(\d+)$/.exec(ref)

	return split === null ? ref : `$${split[1] ?? ""}$${split[2] ?? ""}`
}

// Past a bracketed run (nested; "'" escapes the next character), or -1 when it never closes.
function bracketEnd(text: string, start: number): number {
	let depth = 0

	for (let index = start; index < text.length; index++) {
		const c = text[index]

		if (c === "'") {
			index++

			continue
		}

		if (c === "[") depth++
		if (c === "]" && --depth === 0) return index + 1
	}

	return -1
}

function unescaped(name: string): string {
	return name.replace(/'(.)/g, "$1").trim()
}

// The rows and columns a specifier names, null when it names what is not there.
function areaOf(table: SheetTable, specifier: string, row: number): SheetRange | null {
	const inner = specifier.slice(1, -1).trim()
	const items: string[] = []

	if (inner.startsWith("[")) {
		for (let at = 0; at < inner.length;) {
			if (inner[at] === "[") {
				const end = bracketEnd(inner, at)

				if (end < 0) return null

				items.push(inner.slice(at, end))
				at = end
			} else if (inner[at] === "," || inner[at] === ":" || inner[at] === " ") {
				if (inner[at] === ":") items.push(":")

				at++
			} else if (inner[at] === "@") {
				items.push("[#This Row]")
				at++
			} else return null
		}
	} else if (inner.startsWith("@")) {
		items.push("[#This Row]")

		if (inner.length > 1) items.push(`[${inner.slice(1)}]`)
	} else if (inner !== "") items.push(`[${inner}]`)

	const first = table.ref.startRow
	const data = { start: first + table.headerRows, end: table.ref.endRow - table.totalsRows }
	let rows: { start: number; end: number } | null = null
	let columns: number[] = []
	let range = false

	const extend = (start: number, end: number): void => {
		rows = rows === null ? { start, end } : { start: Math.min(rows.start, start), end: Math.max(rows.end, end) }
	}

	for (const item of items) {
		if (item === ":") {
			range = true

			continue
		}

		const body = item.slice(1, -1)

		switch (body.toLowerCase()) {
			case "#all":
				extend(first, table.ref.endRow)
				break
			case "#data":
				extend(data.start, data.end)
				break
			case "#headers":
				if (table.headerRows === 0) return null
				extend(first, first + table.headerRows - 1)
				break
			case "#totals":
				if (table.totalsRows === 0) return null
				extend(table.ref.endRow - table.totalsRows + 1, table.ref.endRow)
				break
			case "#this row":
				if (row < data.start || row > data.end) return null
				extend(row, row)
				break
			default: {
				const name = unescaped(body).toLowerCase()
				const index = table.columns.findIndex(column => column.toLowerCase() === name)

				if (index < 0) return null

				columns.push(table.ref.startCol + index)
			}
		}
	}

	if (range && columns.length !== 2) return null
	if (columns.length > 2 || (columns.length === 2 && !range)) return null

	const span = rows ?? data

	if (columns.length === 0) columns = [table.ref.startCol, table.ref.endCol]

	return {
		startRow: span.start,
		endRow: span.end,
		startCol: Math.min(...columns),
		endCol: Math.max(...columns)
	}
}

// `text` with every structured reference as the range it names; null when one cannot be said. `here` is the cell
// the formula is in (a this-row reference, and a reference without a table name inside its own table).
export function resolveStructured(
	text: string,
	tables: ReadonlyMap<string, SheetTable>,
	here: { readonly sheet: string; readonly row: number; readonly col: number }
): string | null {
	let out = ""
	let at = 0
	let changed = false

	while (at < text.length) {
		const c = text[at] ?? ""

		if (c === '"' || c === "'") {
			let end = at + 1

			while (end < text.length && !(text[end] === c && text[end + 1] !== c)) end += text[end] === c ? 2 : 1

			out += text.slice(at, end + 1)
			at = end + 1

			continue
		}

		const word = /^[A-Za-z_\\][\w.\\]*/.exec(text.slice(at))?.[0] ?? ""
		const open = at + word.length

		// [1]Sheet!A1 names another workbook, not a table.
		if (word === "" && /^\[\d+\]/.test(text.slice(at))) {
			const end = text.indexOf("]", at) + 1

			out += text.slice(at, end)
			at = end

			continue
		}

		if (text[open] === "[" && (word === "" ? /[\s(,;=+\-*/&^<>:%!]$|^$/.test(out) : !/[\w.!$]$/.test(out))) {
			const end = bracketEnd(text, open)
			const table =
				word === ""
					? [...tables.values()].find(
							candidate =>
								candidate.sheet === here.sheet &&
								here.row >= candidate.ref.startRow &&
								here.row <= candidate.ref.endRow &&
								here.col >= candidate.ref.startCol &&
								here.col <= candidate.ref.endCol
						)
					: tables.get(word.toUpperCase())

			if (end < 0 || table === undefined) return null

			const area = areaOf(table, text.slice(open, end), here.row)

			if (area === null) return null

			const prefix = table.sheet === here.sheet ? "" : `${quotedSheet(table.sheet)}!`
			const startCell = absolute(area.startRow, area.startCol)
			const endCell = absolute(area.endRow, area.endCol)

			out += `${prefix}${startCell}${startCell === endCell ? "" : `:${endCell}`}`
			at = end
			changed = true

			continue
		}

		out += word === "" ? c : word
		at += word === "" ? 1 : word.length
	}

	return changed ? out : null
}

// A sheet's #REF! (Sheet!#REF!, 'My sheet'!#REF!, a reference to a deleted sheet's cells) as #REF!, as LibreOffice
// reads it, outside string literals: a quoted string that spells one stays as it is.
export function sheetErrorsAsErrors(text: string): string {
	const identifier = (c: string): boolean => /[\w.]/.test(c)
	let out = ""
	let at = 0

	while (at < text.length) {
		const c = text[at] ?? ""

		if (c === '"' || c === "'") {
			let end = at + 1

			for (;;) {
				if (end >= text.length) break
				if (text[end] === c) {
					if (text[end + 1] === c) {
						end += 2

						continue
					}

					end++
					break
				}

				end++
			}

			if (c === "'" && text.startsWith("!#REF!", end)) {
				out += "#REF!"
				at = end + 6
			} else {
				out += text.slice(at, end)
				at = end
			}

			continue
		}

		if (/[A-Za-z_]/.test(c) && (at === 0 || !identifier(text[at - 1] ?? ""))) {
			let end = at + 1

			while (end < text.length && identifier(text[end] ?? "")) end++

			if (text.startsWith("!#REF!", end)) {
				out += "#REF!"
				at = end + 6
			} else {
				out += text.slice(at, end)
				at = end
			}

			continue
		}

		out += c
		at++
	}

	return out
}
