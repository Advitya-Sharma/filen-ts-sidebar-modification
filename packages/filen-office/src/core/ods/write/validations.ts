import type { SheetRange } from "../../sheet/protocol"
import type { XlsxSheetModel, XlsxWorkbookModel } from "../../xlsx/model/types"
import { parseRangeRef } from "../../xlsx/write/text"
import { decodeUtf8, encodeUtf8, escapeAttribute, parseXml, type XmlDocument } from "../../xml/index"
import { validationRuleKey } from "../features"
import { openFormula } from "../formula/toOf"
import type { OdsTableSource } from "../source"
import { SourceAxis } from "./axis"
import { paragraphText } from "./cells"

// Validations as the writer needs them where an edit changed them: each rule's table:content-validation name (the
// source's for a rule it held, a new one's otherwise), which cell names which, the rows where that differs from the
// source, and the new rules' elements for table:content-validations (ODF 1.3 Part 3 §9.4.2).

export class ValidationRefused extends Error {}

export interface SheetValidations {
	name(row: number, col: number): string | undefined
	// Where names change along a row: its rules' first columns and the columns after their last.
	bounds(row: number): readonly number[]
	readonly rows: ReadonlySet<number>
	// New table:content-validation elements.
	readonly fresh: readonly string[]
}

interface Rule {
	readonly key: string
	readonly attributes: ReadonlyMap<string, string>
	readonly formula1: string | null
	readonly formula2: string | null
	readonly ranges: readonly SheetRange[]
}

function rules(markup: string): Rule[] {
	if (markup.trim() === "") return []

	const bytes = encodeUtf8(markup)
	const document: XmlDocument = parseXml(bytes)
	const root = document.documentElement()
	const out: Rule[] = []

	for (let node = document.firstChildElement(root); node >= 0; node = document.nextSiblingElement(node)) {
		if (document.localName(node) !== "dataValidation") continue

		const attributes = new Map<string, string>()

		for (const { name, value } of document.attributes(node)) attributes.set(name, value)

		let formula1: string | null = null
		let formula2: string | null = null

		for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
			if (document.localName(child) === "formula1") formula1 = document.textContent(child)
			if (document.localName(child) === "formula2") formula2 = document.textContent(child)
		}

		const ranges = (attributes.get("sqref") ?? "").split(/\s+/).flatMap(text => {
			const range = parseRangeRef(text)

			return range === undefined ? [] : [range]
		})

		out.push({
			key: validationRuleKey(decodeUtf8(bytes.subarray(document.start(node), document.end(node)))),
			attributes,
			formula1,
			formula2,
			ranges
		})
	}

	return out
}

const COMPARE: Readonly<Record<string, string>> = {
	equal: "=",
	notEqual: "!=",
	lessThan: "<",
	lessThanOrEqual: "<=",
	greaterThan: ">",
	greaterThanOrEqual: ">="
}

const KINDS: Readonly<Record<string, string>> = { whole: "whole-number", decimal: "decimal-number", date: "date", time: "time" }

function expression(formula: string): string {
	const written = openFormula(formula)

	if ("refused" in written) throw new ValidationRefused(`A validation's formula: ${written.refused}`)

	return written.text.replace(/^of:=/, "")
}

// The rule as table:condition (§19.618), the inverse of features.ts validationRule; undefined for any value.
function condition(rule: Rule): string | undefined {
	const type = rule.attributes.get("type") ?? "none"
	const operator = rule.attributes.get("operator") ?? "between"
	const first = rule.formula1 ?? ""
	const second = rule.formula2 ?? ""

	switch (type) {
		case "none":
			return undefined
		case "list": {
			const quoted = /^"((?:[^"]|"")*)"$/.exec(first.trim())

			if (quoted === null) return `of:cell-content-is-in-list(${expression(first)})`

			const items = (quoted[1] ?? "").replaceAll('""', '"').split(",")

			return `of:cell-content-is-in-list(${items.map(item => `"${item.replaceAll('"', '""')}"`).join(";")})`
		}
		case "custom":
			return `of:is-true-formula(${expression(first)})`
		case "textLength":
			if (operator === "between" || operator === "notBetween")
				return `of:cell-content-text-length-is-${operator === "notBetween" ? "not-" : ""}between(${expression(first)};${expression(second)})`

			return `of:cell-content-text-length()${COMPARE[operator] ?? "="}${expression(first)}`
		default: {
			const kind = KINDS[type]

			if (kind === undefined) throw new ValidationRefused(`A ${type} validation`)
			if (operator === "between" || operator === "notBetween")
				return `of:cell-content-is-${kind}() and cell-content-is-${operator === "notBetween" ? "not-" : ""}between(${expression(first)};${expression(second)})`

			return `of:cell-content-is-${kind}() and cell-content()${COMPARE[operator] ?? "="}${expression(first)}`
		}
	}
}

function message(tag: string, title: string, text: string, display: boolean, kind: string | undefined): string {
	const lines =
		text === ""
			? ""
			: text
					.split("\n")
					.map(line => `<text:p>${paragraphText(line)}</text:p>`)
					.join("")

	return `<table:${tag}${kind === undefined ? "" : ` table:message-type="${kind}"`} table:title="${escapeAttribute(title)}" table:display="${display ? "true" : "false"}">${lines}</table:${tag}>`
}

function column(index: number): string {
	let text = ""

	for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + ((n - 1) % 26)) + text

	return text
}

// A table:content-validation for a rule on `sheet`, its formulas relative to its first range's first cell.
function element(rule: Rule, name: string, sheet: string): string {
	const first = rule.ranges[0]
	const quoted = /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheet) ? sheet : `'${sheet.replaceAll("'", "''")}'`
	const base = first === undefined ? undefined : `${quoted}.${column(first.startCol)}${first.startRow + 1}`
	const test = condition(rule)
	const attributes = rule.attributes
	const list = (attributes.get("type") ?? "") === "list"

	return `<table:content-validation table:name="${escapeAttribute(name)}"${test === undefined ? "" : ` table:condition="${escapeAttribute(test)}"`} table:allow-empty-cell="${attributes.get("allowBlank") === "1" || attributes.get("allowBlank") === "true" ? "true" : "false"}"${list ? ` table:display-list="${attributes.get("showDropDown") === "1" ? "none" : "unsorted"}"` : ""}${base === undefined ? "" : ` table:base-cell-address="${escapeAttribute(base)}"`}>${message(
		"help-message",
		attributes.get("promptTitle") ?? "",
		attributes.get("prompt") ?? "",
		attributes.get("showInputMessage") === "1",
		undefined
	)}${message(
		"error-message",
		attributes.get("errorTitle") ?? "",
		attributes.get("error") ?? "",
		attributes.get("showErrorMessage") === "1",
		attributes.get("errorStyle") ?? "stop"
	)}</table:content-validation>`
}

export interface Rect {
	readonly range: SheetRange
	readonly name: string
}

function namesAt(rects: readonly Rect[], row: number, col: number): string | undefined {
	for (const rect of rects) {
		const range = rect.range

		if (row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol) return rect.name
	}

	return undefined
}

// A dataValidations section as rectangles each naming its rule (by its key): what two sheets compare by.
export function validationRects(markup: string): Rect[] {
	return rules(markup).flatMap(rule => rule.ranges.map(range => ({ range, name: rule.key })))
}

// Whether two sets of rectangles name the same rule on every cell.
export function sameRects(a: readonly Rect[], b: readonly Rect[]): boolean {
	const rows = new Set<number>()

	for (const rect of [...a, ...b])
		for (let row = rect.range.startRow; row <= rect.range.endRow && rows.size < 2_000_000; row++) rows.add(row)

	const cells = (rects: readonly Rect[], row: number): string => {
		const out: string[] = []

		for (const rect of rects) {
			if (row < rect.range.startRow || row > rect.range.endRow) continue

			for (let col = rect.range.startCol; col <= rect.range.endCol; col++) out.push(`${col}:${rect.name}`)
		}

		return out.sort().join("|")
	}

	for (const row of rows) if (cells(a, row) !== cells(b, row)) return false

	return true
}

// A row's names as column runs, for comparing two of them.
function rowKey(rects: readonly Rect[], row: number): string {
	return rects
		.filter(rect => row >= rect.range.startRow && row <= rect.range.endRow)
		.map(rect => `${rect.range.startCol}-${rect.range.endCol}:${rect.name}`)
		.sort()
		.join("|")
}

// A sheet's validations where an edit changed them (its dataValidations section written), null where it did not.
// `taken`: names in use across the workbook (new ones are added to it).
export function sheetValidations(
	workbook: XlsxWorkbookModel,
	model: XlsxSheetModel,
	sheetName: string,
	table: OdsTableSource | null,
	current: string | null,
	taken: Set<string>
): SheetValidations | null {
	const axes = model as { rowAxis?: SourceAxis; colAxis?: SourceAxis }
	const rowAxis = axes.rowAxis ?? SourceAxis.identity(0)
	const colAxis = axes.colAxis ?? SourceAxis.identity(0)

	void workbook

	if (current === null && table !== null && rowAxis.identity && colAxis.identity) return null

	const now = rules(
		current ??
			model.sections
				.filter(section => section.name === "dataValidations")
				.map(section => decodeUtf8(section.bytes ?? new Uint8Array(0)))
				.join("")
	)
	const fresh: string[] = []
	const named: Rect[] = []

	for (const rule of now) {
		let name = table?.validationNames.get(rule.key)

		if (name === undefined) {
			for (let n = 1; ; n++) {
				const candidate = `valF${String(n)}`

				if (!taken.has(candidate)) {
					name = candidate
					taken.add(candidate)

					break
				}
			}

			fresh.push(element(rule, name, sheetName))
		}

		for (const range of rule.ranges) named.push({ range, name })
	}

	// What the source held: its section's rules, by the names they came from, moved to where their cells are now.
	const before: Rect[] = []

	if (table !== null) {
		const source = rules(
			table.sections
				.filter(section => section.name === "dataValidations")
				.map(section => decodeUtf8(section.bytes ?? new Uint8Array(0)))
				.join("")
		)

		for (const rule of source) {
			const name = table.validationNames.get(rule.key)

			if (name === undefined) continue

			for (const range of rule.ranges) {
				const startCol = colAxis.target(range.startCol)
				const endCol = colAxis.target(range.endCol)

				for (let row = range.startRow; row <= range.endRow; row++) {
					const target = rowAxis.target(row)

					if (target < 0 || startCol < 0 || endCol < 0) continue

					before.push({ range: { startRow: target, endRow: target, startCol, endCol }, name })
				}
			}
		}
	}

	const rows = new Set<number>()
	const candidates = new Set<number>()

	for (const rect of [...named, ...before]) for (let row = rect.range.startRow; row <= rect.range.endRow; row++) candidates.add(row)
	for (const row of candidates) if (rowKey(named, row) !== rowKey(before, row)) rows.add(row)

	return {
		name: (row, col) => namesAt(named, row, col),
		bounds: row =>
			named
				.filter(rect => row >= rect.range.startRow && row <= rect.range.endRow)
				.flatMap(rect => [rect.range.startCol, rect.range.endCol + 1]),
		rows,
		fresh
	}
}
