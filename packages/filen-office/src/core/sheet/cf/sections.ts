import type { XlsxSheetModel, XlsxWorkbookModel } from "../../xlsx/model/types"
import { decodeUtf8 } from "../../xml/index"

// A worksheet section as it now is: the source's markup, or what an edit (the reference catalogue, a validation
// edit) writes for it. Read again only when the session's version moved; parsed again only when it changed.

export interface SectionMarkup {
	// What tells this markup from another: the source's byte arrays (compared by identity), or the written text.
	readonly key: readonly unknown[]
	markup(): readonly string[]
}

export function sectionMarkup(workbook: XlsxWorkbookModel, model: XlsxSheetModel, name: string): SectionMarkup {
	const writer = workbook.changes.sheets.get(model.id)?.sections.get(name)

	if (writer === undefined || writer === null) {
		const sections = model.sections.filter(section => section.name === name)

		return {
			key: sections.map(section => section.bytes),
			markup: () => sections.map(section => decodeUtf8(section.bytes ?? new Uint8Array(0)))
		}
	}

	const written = writer(model, workbook)
	const text = typeof written === "string" ? written : decodeUtf8(written)

	return { key: [text], markup: () => (text === "" ? [] : [text]) }
}

export function sameKey(a: readonly unknown[], b: readonly unknown[]): boolean {
	return a.length === b.length && a.every((item, i) => item === b[i])
}

// A cache of something read from a sheet's sections, per sheet: checked against the sections once per session
// version, rebuilt when they changed.
export class SectionCache<T> {
	private readonly entries = new Map<number, { version: number; key: unknown[]; value: T }>()
	private readonly names: readonly string[]
	private readonly build: (markup: readonly (readonly string[])[], model: XlsxSheetModel) => T

	constructor(names: readonly string[], build: (markup: readonly (readonly string[])[], model: XlsxSheetModel) => T) {
		this.names = names
		this.build = build
	}

	get(workbook: XlsxWorkbookModel, model: XlsxSheetModel, version: number): T {
		const cached = this.entries.get(model.id)

		if (cached !== undefined && cached.version === version) return cached.value

		// Only the keys are compared while nothing changed: the source's byte arrays by identity.
		const sections = this.names.map(name => sectionMarkup(workbook, model, name))
		const key = sections.flatMap(section => [section.key.length, ...section.key])

		if (cached !== undefined && sameKey(cached.key, key)) {
			cached.version = version

			return cached.value
		}

		const value = this.build(
			sections.map(section => section.markup()),
			model
		)

		this.entries.set(model.id, { version, key, value })

		return value
	}

	// What is held now for a sheet, without checking it.
	peek(sheet: number): T | undefined {
		return this.entries.get(sheet)?.value
	}

	forEach(visit: (value: T, sheet: number) => void): void {
		for (const [sheet, entry] of this.entries) visit(entry.value, sheet)
	}
}
