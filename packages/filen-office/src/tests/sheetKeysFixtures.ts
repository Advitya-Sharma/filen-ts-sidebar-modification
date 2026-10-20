import type { SheetEvent, SheetOpenOptions } from "../core/sheet/protocol"
import { inThreadSheetHost } from "../dom/sheet/host"
import { IMMEDIATE, buildFixture, type OpenedFixture } from "./sheetSessionFixtures"
import type { Fixture } from "./xlsxWriteFixtures"

// Sessions over small workbooks and text files, for the find, navigation and fill tests.

export async function openKeys(source: Fixture | Uint8Array | string, options: Partial<SheetOpenOptions> = {}): Promise<OpenedFixture> {
	const text = typeof source === "string"
	const bytes = text ? new TextEncoder().encode(source) : source instanceof Uint8Array ? source : buildFixture(source)
	const events: SheetEvent[] = []
	const host = inThreadSheetHost()
	const result = await host.open(bytes.slice(), { kind: text ? "csv" : "xlsx", slicing: IMMEDIATE, ...options }, event =>
		events.push(event)
	)

	if (result.kind !== "opened") throw new Error(`expected an open, got ${result.kind}`)

	return { session: result.session, info: result.info, events, bytes }
}

// A formula with the result the file keeps for it.
export interface FormulaCell {
	readonly f: string
	readonly v: string | number
}

// Rows of cells as a worksheet body: numbers, "=..." formulas without a kept result, formulas with one, else inline
// strings.
export function sheetBody(rows: readonly (readonly (string | number | FormulaCell | null)[])[]): string {
	const name = (row: number, col: number): string => {
		let letters = ""

		for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters

		return `${letters}${String(row + 1)}`
	}
	const escape = (text: string): string => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
	let xml = "<sheetData>"

	rows.forEach((values, row) => {
		xml += `<row r="${String(row + 1)}">`

		values.forEach((value, col) => {
			if (value === null) return

			const ref = name(row, col)

			if (typeof value === "number") xml += `<c r="${ref}"><v>${String(value)}</v></c>`
			else if (typeof value === "object")
				xml +=
					typeof value.v === "number"
						? `<c r="${ref}"><f>${escape(value.f)}</f><v>${String(value.v)}</v></c>`
						: `<c r="${ref}" t="str"><f>${escape(value.f)}</f><v>${escape(value.v)}</v></c>`
			else if (value.startsWith("=")) xml += `<c r="${ref}"><f>${escape(value.slice(1))}</f></c>`
			else xml += `<c r="${ref}" t="inlineStr"><is><t>${escape(value)}</t></is></c>`
		})

		xml += "</row>"
	})

	return `${xml}</sheetData>`
}
