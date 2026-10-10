import type { DocxModel } from "./document"
import { ITEM_NAMES, Item, REVISION_NAMES, SdtLevel } from "./kinds"

export interface DocxModelStats {
	readonly parts: number
	readonly items: number
	readonly stories: Readonly<Record<string, number>>
	readonly kinds: Readonly<Record<string, number>>
	readonly paragraphs: number
	readonly runs: number
	readonly tables: number
	readonly nestedTables: number
	readonly maxTableDepth: number
	readonly rows: number
	readonly cells: number
	readonly sdts: { readonly block: number; readonly run: number; readonly row: number; readonly cell: number }
	readonly fields: { readonly complex: number; readonly simple: number; readonly unterminated: number; readonly unbalancedEnds: number }
	readonly fieldTypes: Readonly<Record<string, number>>
	readonly bookmarks: number
	readonly comments: number
	readonly commentAnchors: number
	readonly revisions: Readonly<Record<string, number>>
	readonly drawings: Readonly<Record<string, number>>
	readonly anchoredDrawings: number
	readonly vml: number
	readonly alternateContent: { readonly total: number; readonly choice: number; readonly fallback: number; readonly none: number }
	readonly sections: number
	readonly footnotes: number
	readonly endnotes: number
	// Opaque items by qualified element name.
	readonly opaque: Readonly<Record<string, number>>
	readonly proofErrors: number
	readonly ignoredSectionProperties: number
	readonly depthLimited: number
}

function bump(record: Record<string, number>, key: string, by: number = 1): void {
	record[key] = (record[key] ?? 0) + by
}

// Counts over the whole model, for tests and diagnostics. Field types decode instructions, so they are
// gathered only when asked for.
export function modelStats(model: DocxModel, options: { fieldTypes?: boolean } = {}): DocxModelStats {
	const stories: Record<string, number> = {}
	const kinds: Record<string, number> = {}
	const revisions: Record<string, number> = {}
	const drawings: Record<string, number> = {}
	const opaque: Record<string, number> = {}
	const fieldTypes: Record<string, number> = {}
	const sdts = { block: 0, run: 0, row: 0, cell: 0 }
	const alternates = { total: 0, choice: 0, fallback: 0, none: 0 }
	let items = 0
	let paragraphs = 0
	let runs = 0
	let tables = 0
	let nestedTables = 0
	let maxTableDepth = 0
	let rows = 0
	let cells = 0
	let complex = 0
	let simple = 0
	let bookmarks = 0

	for (const story of model.stories) bump(stories, story.kind)

	for (const part of model.parts) {
		items += part.count
		paragraphs += part.paragraphs.count
		runs += part.runs.count
		tables += part.tables.count
		rows += part.rows.count
		cells += part.cells.count
		bookmarks += part.bookmarks.count

		for (let item = 0; item < part.count; item++) {
			const kind = part.kinds[item] ?? 0

			bump(kinds, ITEM_NAMES[kind] ?? String(kind))

			if (kind === Item.Opaque) bump(opaque, part.document.name(part.node(item)))
		}

		for (let table = 0; table < part.tables.count; table++) {
			const depth = part.tables.depth[table] ?? 0

			if (depth > 0) nestedTables++

			maxTableDepth = Math.max(maxTableDepth, depth)
		}

		for (let sdt = 0; sdt < part.sdts.count; sdt++) {
			switch (part.sdtLevel(sdt)) {
				case SdtLevel.Block:
					sdts.block++

					break
				case SdtLevel.Run:
					sdts.run++

					break
				case SdtLevel.Row:
					sdts.row++

					break
				case SdtLevel.Cell:
					sdts.cell++

					break
			}
		}

		for (let field = 0; field < part.fields.count; field++) {
			if (part.fields.simple[field] === 1) simple++
			else complex++

			if (options.fieldTypes === true) bump(fieldTypes, part.fieldType(field) || "(none)")
		}

		for (let revision = 0; revision < part.revisions.count; revision++) {
			const kind = part.revisions.kind[revision] ?? 0

			bump(revisions, REVISION_NAMES[kind] ?? String(kind))
		}

		for (let alternate = 0; alternate < part.alternates.count; alternate++) {
			const selected = part.alternates.selected[alternate] ?? -1

			alternates.total++

			if (selected < 0) alternates.none++
			else if (selected === part.alternates.fallback[alternate]) alternates.fallback++
			else alternates.choice++
		}
	}

	let anchoredDrawings = 0

	for (const drawing of model.drawings) {
		bump(drawings, drawing.graphic)

		if (drawing.placement === "anchor") anchoredDrawings++
	}

	return {
		parts: model.parts.length,
		items,
		stories,
		kinds,
		paragraphs,
		runs,
		tables,
		nestedTables,
		maxTableDepth,
		rows,
		cells,
		sdts,
		fields: { complex, simple, unterminated: model.counts.unterminatedFields, unbalancedEnds: model.counts.unbalancedFieldEnds },
		fieldTypes,
		bookmarks,
		comments: model.comments.length,
		commentAnchors: model.commentAnchors.count,
		revisions,
		drawings,
		anchoredDrawings,
		vml: model.vml.length,
		alternateContent: alternates,
		sections: model.sections.length,
		footnotes: model.footnotes.size,
		endnotes: model.endnotes.size,
		opaque,
		proofErrors: model.counts.proofErrors,
		ignoredSectionProperties: model.counts.ignoredSectionProperties,
		depthLimited: model.counts.depthLimited
	}
}
