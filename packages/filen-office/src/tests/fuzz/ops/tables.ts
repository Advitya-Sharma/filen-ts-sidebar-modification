import type { Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { docxSchema, type DocxEditorCommand, type DocxEditorSession } from "../harness"
import { textPositions, textblocks } from "./base"
import type { EditOp, OpContext, Rand, ScriptEdit } from "./index"

// Table structure and properties: rows, columns, merges and splits, widths, header rows, shading and
// vertical alignment, new and deleted tables, through the table commands (also while changes are tracked).

const nodes = docxSchema.nodes

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

function select(session: DocxEditorSession, a: number, b: number = a): boolean {
	const doc = session.state.doc

	try {
		return session.dispatch(
			session.state.tr.setSelection(TextSelection.between(doc.resolve(Math.min(a, b)), doc.resolve(Math.max(a, b))))
		)
	} catch {
		return false
	}
}

// The innermost table around `pos`, with its position.
function tableAround(doc: Node, pos: number): { node: Node; pos: number } | undefined {
	const $pos = doc.resolve(pos)

	for (let d = $pos.depth; d >= 1; d--) if ($pos.node(d).type === nodes.table) return { node: $pos.node(d), pos: $pos.before(d) }

	return undefined
}

// A caret in a random table cell, or a selection from it to another cell of the same table.
function selectCells(context: OpContext): string | undefined {
	const session = context.session
	const doc = session.state.doc
	const positions = textPositions(doc, true)

	if (positions.length === 0) return undefined

	const p = pick(context.rand, positions)
	const table = tableAround(doc, p)
	let q = p

	if (table !== undefined && context.rand() < 0.4) {
		const inside = positions.filter(
			at => at > table.pos && at < table.pos + table.node.nodeSize && tableAround(doc, at)?.pos === table.pos
		)

		q = pick(context.rand, inside)
	}

	return select(session, p, q) ? `${p}-${q}` : undefined
}

function grid(session: DocxEditorSession): readonly number[] {
	return session.snapshot().context.table?.widths ?? []
}

const STRUCTURE: readonly ((rand: Rand) => DocxEditorCommand)[] = [
	() => ({ type: "table.row.insertAbove" }),
	() => ({ type: "table.row.insertBelow" }),
	() => ({ type: "table.row.delete" }),
	() => ({ type: "table.column.insertLeft" }),
	() => ({ type: "table.column.insertRight" }),
	() => ({ type: "table.column.delete" })
]

const MERGES: readonly ((rand: Rand) => DocxEditorCommand)[] = [
	() => ({ type: "table.cells.merge" }),
	rand => ({ type: "table.cell.split", rows: 1 + Math.floor(rand() * 3), cols: 1 + Math.floor(rand() * 3) })
]

const FILLS = ["FFEE00", "D9E2F3", "000000", null]

const PROPERTIES: readonly ((rand: Rand, widths: readonly number[]) => DocxEditorCommand)[] = [
	() => ({ type: "table.row.header" }),
	rand => ({ type: "table.align", align: pick(rand, ["start", "center", "end"] as const) }),
	rand => ({ type: "table.cell.shading", fill: pick(rand, FILLS) }),
	rand => ({ type: "table.cell.vAlign", align: pick(rand, ["top", "center", "bottom"] as const) }),
	(rand, widths) => ({
		type: "table.column.width",
		index: Math.floor(rand() * Math.max(1, widths.length)),
		twips: Math.round(200 + rand() * 4000)
	}),
	() => ({ type: "table.distributeColumns" })
]

// Runs one of `commands` on the cells of a random selection in a table.
function tableCommand(
	context: OpContext,
	commands: readonly ((rand: Rand, widths: readonly number[]) => DocxEditorCommand)[],
	tracked = false
): string {
	const where = selectCells(context)

	if (where === undefined) return "noop"

	const command = pick(context.rand, commands)(context.rand, grid(context.session))
	const review = context.session.context.review
	const identity = context.session.context.identity
	const author = identity.author

	if (tracked) {
		review.tracking = true
		identity.author = "Fuzz Author"
	}

	try {
		return `${command.type}@${where} -> ${context.execute(command.type, command)}`
	} finally {
		if (tracked) {
			review.tracking = false
			identity.author = author
		}
	}
}

// A new table at the start of a random paragraph (in a cell sometimes).
function newTable(context: OpContext): string {
	const blocks = textblocks(context.session.state.doc)

	if (blocks.length === 0) return "noop"

	const block = pick(context.rand, blocks)

	if (!select(context.session, block.pos + 1)) return "noop"

	const command: DocxEditorCommand = {
		type: "table.insert",
		rows: 1 + Math.floor(context.rand() * 4),
		cols: 1 + Math.floor(context.rand() * 5)
	}

	return `table.insert ${command.rows}x${command.cols}@${block.pos} -> ${context.execute(command.type, command)}`
}

// ---- scripted suite -------------------------------------------------------------------------------------

// The first table of the body (`nested`: the first table inside a cell), with its position.
function scriptTable(session: DocxEditorSession, nested = false): { node: Node; pos: number } | undefined {
	let found: { node: Node; pos: number } | undefined

	session.state.doc.descendants((node, pos, parent) => {
		if (found !== undefined || node.isInline || node.type === nodes.locked_block) return false

		if (node.type === nodes.table && (parent?.type === nodes.table_cell) === nested) {
			found = { node, pos }

			return false
		}

		return !node.isTextblock
	})

	return found
}

// The position inside the first textblock of cell `index` of row `row`.
function cellText(table: { node: Node; pos: number }, row: number, index: number): number | undefined {
	if (row >= table.node.childCount || index >= table.node.child(row).childCount) return undefined

	let pos = table.pos + 1

	for (let r = 0; r < row; r++) pos += table.node.child(r).nodeSize

	const rowNode = table.node.child(row)

	pos += 1

	for (let c = 0; c < index; c++) pos += rowNode.child(c).nodeSize

	const cell = rowNode.child(index)
	let at: number | undefined

	cell.descendants((node, offset) => {
		if (at === undefined && node.isTextblock) at = pos + 1 + offset + 1

		return at === undefined
	})

	return at
}

// Runs `command` with the selection from cell (row, index) to cell (toRow, toIndex) of the first table.
function onScriptTable(
	session: DocxEditorSession,
	command: DocxEditorCommand,
	row: number,
	index: number,
	to: readonly [number, number] = [row, index],
	nested = false
): boolean {
	const table = scriptTable(session, nested)

	if (table === undefined) return false

	const a = cellText(table, row, index)
	const b = cellText(table, to[0], to[1])

	return a !== undefined && b !== undefined && select(session, a, b) && session.execute(command)
}

function script(order: number, run: (session: DocxEditorSession) => boolean): ScriptEdit {
	return { order, run: ({ session }) => run(session) }
}

export const tableOps: readonly EditOp[] = [
	{
		name: "tableRows",
		harness: "paged",
		weight: 3,
		run: context => tableCommand(context, STRUCTURE.slice(0, 3)),
		script: script(300, session => onScriptTable(session, { type: "table.row.insertBelow" }, 0, 0))
	},
	{
		name: "tableColumns",
		harness: "paged",
		weight: 3,
		run: context => tableCommand(context, STRUCTURE.slice(3)),
		script: script(310, session => onScriptTable(session, { type: "table.column.insertRight" }, 0, 0))
	},
	{
		name: "tableMerge",
		harness: "paged",
		weight: 3,
		run: context => tableCommand(context, MERGES),
		script: script(320, session => onScriptTable(session, { type: "table.cells.merge" }, 1, 0, [1, 1]))
	},
	{
		name: "tableSplit",
		harness: "paged",
		weight: 0,
		run: context => tableCommand(context, MERGES.slice(1)),
		script: script(330, session => onScriptTable(session, { type: "table.cell.split", rows: 2, cols: 2 }, 1, 0))
	},
	{
		name: "tableProps",
		harness: "paged",
		weight: 3,
		run: context => tableCommand(context, PROPERTIES),
		script: script(
			340,
			session =>
				onScriptTable(session, { type: "table.cell.shading", fill: "D9E2F3" }, 0, 0) &&
				onScriptTable(session, { type: "table.cell.vAlign", align: "center" }, 0, 0) &&
				onScriptTable(session, { type: "table.row.header" }, 0, 0)
		)
	},
	{
		name: "tableWidths",
		harness: "paged",
		weight: 0,
		run: context => tableCommand(context, PROPERTIES.slice(4)),
		script: script(
			350,
			session =>
				onScriptTable(session, { type: "table.distributeColumns" }, 0, 0) &&
				onScriptTable(session, { type: "table.column.width", index: 0, twips: 1800 }, 0, 0)
		)
	},
	{
		name: "tableDeleteRows",
		harness: "paged",
		weight: 0,
		run: context => tableCommand(context, [STRUCTURE[2] as (rand: Rand) => DocxEditorCommand]),
		script: script(360, session => onScriptTable(session, { type: "table.row.delete" }, 2, 0))
	},
	{
		name: "tableDeleteColumns",
		harness: "paged",
		weight: 0,
		run: context => tableCommand(context, [STRUCTURE[5] as (rand: Rand) => DocxEditorCommand]),
		script: script(370, session => onScriptTable(session, { type: "table.column.delete" }, 0, 1))
	},
	{
		name: "tableNested",
		harness: "paged",
		weight: 0,
		run: context => tableCommand(context, [...STRUCTURE, ...MERGES]),
		script: script(380, session => onScriptTable(session, { type: "table.row.insertAbove" }, 0, 0, [0, 0], true))
	},
	{
		name: "tableTracked",
		harness: "paged",
		weight: 1,
		run: context => tableCommand(context, STRUCTURE, true)
	},
	{
		name: "tableInsert",
		harness: "paged",
		weight: 1,
		run: newTable,
		script: script(390, session => {
			const blocks = textblocks(session.state.doc).filter(block => session.state.doc.resolve(block.pos).depth === 0)
			const block = blocks[Math.min(blocks.length - 1, 9)]

			return block !== undefined && select(session, block.pos + 1) && session.execute({ type: "table.insert", rows: 2, cols: 3 })
		})
	},
	{
		name: "tableDelete",
		harness: "paged",
		weight: 0,
		run: context => tableCommand(context, [() => ({ type: "table.delete" })])
	},
	{
		name: "tableCommand",
		harness: "damage",
		run: context =>
			context.rand() < 0.15
				? newTable(context)
				: tableCommand(context, [...STRUCTURE, ...MERGES, ...PROPERTIES, () => ({ type: "table.delete" })], context.rand() < 0.2)
	}
]
