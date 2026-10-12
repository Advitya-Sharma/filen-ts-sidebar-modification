import type { Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { markCommand, markInternal } from "../../../core/docx/edit/invariants/index"
import { addBookmark, sessionBookmarks } from "../../../core/docx/fields/bookmarks"
import { rewritableSource } from "../../../core/docx/fields/evaluate"
import { docxSchema, type DocxEditorSession } from "../harness"
import { inlinePosition, nearPosition, textblocks } from "./base"
import type { EditOp, OpContext, Rand } from "./index"

// Fields, bookmarks and tables of contents: updates, insertions and bookmark commands under the paged editor
// (whose pages must stay those of the saved file), raw results, field characters, bookmarks and references
// for the writer and Gate 2, and two edits of the scripted suite.

const nodes = docxSchema.nodes
const NAMES = ["Alpha", "Beta", "Gamma", "Results2026", "naïve", "Kraków"]
const DATE_PICTURES = ["d MMMM yyyy", "M/d/yyyy", "dddd, d MMM yy", "HH:mm", "h:mm am/pm"]

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

function fieldsOf(doc: Node): { node: Node; pos: number }[] {
	const out: { node: Node; pos: number }[] = []

	doc.descendants((node, pos) => {
		if (node.type === nodes.field) out.push({ node, pos })

		return !node.isLeaf
	})

	return out
}

function charsOf(doc: Node): number[] {
	const out: number[] = []

	doc.descendants((node, pos) => {
		if (node.type === nodes.field_char) out.push(pos)

		return !node.isLeaf
	})

	return out
}

function select(session: DocxEditorSession, a: number, b: number = a): void {
	const doc = session.state.doc

	session.dispatch(session.state.tr.setSelection(TextSelection.between(doc.resolve(Math.min(a, b)), doc.resolve(Math.max(a, b)))))
}

function paged(name: string, weight: number, run: (context: OpContext) => string): EditOp {
	return { name, harness: "paged", weight, run: context => (context.editor === undefined ? "noop" : run(context)) }
}

// A caret in a random paragraph of the story.
function caret(context: OpContext): number | undefined {
	const session = context.session
	const at = inlinePosition(context.rand, session.state.doc)

	if (at !== undefined) select(session, at)

	return at
}

// ---- damage: raw results, characters, bookmarks and references ----

function resultDamage(context: OpContext): string {
	const session = context.session
	const fields = fieldsOf(session.state.doc)

	if (fields.length === 0) return "noop"

	const { node, pos } = pick(context.rand, fields)
	const result = pick(context.rand, ["", "1", "x\ty", "two\nlines", "Kraków & <co>"])

	return `fieldResult ${String(pos)} ${JSON.stringify(result)} -> ${context.transact("fieldResult", tr => markCommand(tr.setNodeAttribute(pos, "result", result)))}${node.attrs["src"] === -1 ? " (new)" : ""}`
}

function charDamage(context: OpContext): string {
	const at = inlinePosition(context.rand, context.session.state.doc)

	if (at === undefined) return "noop"

	const kind = pick(context.rand, ["begin", "separate", "end"] as const)
	const node = nodes.field_char.create({ kind, instruction: kind === "begin" ? ' TOC \\o "1-3" ' : "", dirty: context.rand() < 0.5 })

	return `fieldChar ${kind}@${String(at)} -> ${context.transact("fieldChar", tr => markInternal(tr.insert(at, node)))}`
}

function bookmarkDamage(context: OpContext): string {
	const rand = context.rand
	const doc = context.session.state.doc
	const a = inlinePosition(rand, doc)

	if (a === undefined) return "noop"

	const b = nearPosition(rand, doc, a) ?? a
	const id = String(pick(rand, [0, 1, 2, 77, Math.floor(rand() * 1000)]))
	const name = pick(rand, [
		...NAMES,
		"_Toc123456789",
		"x".repeat(41),
		"",
		...sessionBookmarks(context.session).list.map(bookmark => bookmark.name)
	])
	const end = rand() < 0.8

	return `bookmarkRaw ${name}#${id} ${String(a)}..${String(b)} -> ${context.transact("bookmarkRaw", tr => {
		if (end) tr.insert(Math.max(a, b), nodes.marker.create({ kind: "bookmarkEnd", id }))

		tr.insert(Math.min(a, b), nodes.marker.create({ kind: "bookmarkStart", id, name }))

		return markInternal(tr)
	})}`
}

function referenceDamage(context: OpContext): string {
	const at = inlinePosition(context.rand, context.session.state.doc)

	if (at === undefined) return "noop"

	const target = pick(context.rand, [...NAMES, ...sessionBookmarks(context.session).list.map(bookmark => bookmark.name)])
	const type = pick(context.rand, ["REF", "PAGEREF"])
	const field = nodes.field.create({ type, instruction: ` ${type} ${target} \\h `, result: target, simple: context.rand() < 0.3 })

	return `referenceRaw ${type} ${target}@${String(at)} -> ${context.transact("referenceRaw", tr => tr.insert(at, field))}`
}

// ---- scripted suite ----

// The first word of the body's first paragraph with text, bookmarked, and a reference to it at the paragraph's
// end.
function scriptedReference(session: DocxEditorSession): boolean {
	const block = textblocks(session.state.doc).find(candidate => /\p{L}{2,}/u.test(candidate.node.textContent))

	if (block === undefined) return false

	let from = -1
	let to = -1

	block.node.forEach((child, offset) => {
		if (from >= 0 || !child.isText) return

		const match = /\p{L}{2,}/u.exec(child.text ?? "")

		if (match !== null) {
			from = block.pos + 1 + offset + match.index
			to = from + match[0].length
		}
	})

	if (from < 0) return false

	select(session, from, to)

	if (!addBookmark(session, [session], "FilenScripted")) return false

	const end = session.state.doc.resolve(from).end()
	const text = session.state.doc.textBetween(from + 1, to + 1)
	const field = nodes.field.create({ type: "REF", instruction: " REF FilenScripted \\h ", result: text })

	return session.dispatch(session.state.tr.insert(end, field))
}

// A new result for the body's first source field whose result can be written anew.
function scriptedResult(session: DocxEditorSession): boolean {
	const part = session.story.part
	const field = fieldsOf(session.state.doc).find(({ node }) => node.attrs["src"] !== -1 && rewritableSource(part, node) !== undefined)

	return (
		field !== undefined &&
		session.dispatch(markCommand(session.state.tr.setNodeAttribute(field.pos, "result", `${field.node.attrs["result"] as string}*`)))
	)
}

export const fieldOps: readonly EditOp[] = [
	paged("fieldUpdateAll", 2, context => `field.updateAll -> ${context.execute("field.updateAll", { type: "field.updateAll" })}`),
	paged("fieldUpdate", 2, context => {
		const session = context.session
		const fields = fieldsOf(session.state.doc)

		const at = fields.length > 0 && context.rand() < 0.7 ? pick(context.rand, fields).pos : caret(context)

		if (at !== undefined) select(session, at)

		return `field.update@${String(at)} -> ${context.execute("field.update", { type: "field.update" })}`
	}),
	paged("tocInsert", 1, context => {
		const at = context.session === context.editor?.body.session ? caret(context) : undefined

		if (at === undefined) return "noop"

		const levels = 1 + Math.floor(context.rand() * 4)
		const hyperlinks = context.rand() < 0.7

		return `toc.insert@${String(at)} ${String(levels)} ${String(hyperlinks)} -> ${context.execute("toc.insert", { type: "toc.insert", levels, hyperlinks })}`
	}),
	paged("tocUpdate", 1, context => {
		const pagesOnly = context.rand() < 0.5

		return `toc.update ${String(pagesOnly)} -> ${context.execute("toc.update", { type: "toc.update", pagesOnly })}`
	}),
	paged("tocCut", 1, context => {
		const session = context.session
		const chars = charsOf(session.state.doc)

		if (chars.length === 0) return "noop"

		const at = pick(context.rand, chars)
		const other = nearPosition(context.rand, session.state.doc, at) ?? at + 1

		select(session, Math.min(at, other), Math.max(at + 1, other))

		return `tocCut ${String(at)} -> ${context.transact("tocCut", tr => tr.deleteSelection())}`
	}),
	paged("insertField", 2, context => {
		const at = caret(context)

		if (at === undefined) return "noop"

		const bookmarks = sessionBookmarks(context.session).list
		const target = bookmarks.length > 0 ? pick(context.rand, bookmarks).name : "Missing"
		const choice = pick(context.rand, ["page", "count", "date", "seq", "ref", "pageref", "section"] as const)

		switch (choice) {
			case "page":
				return `insert.pageNumber@${String(at)} -> ${context.execute("insert.pageNumber", { type: "insert.pageNumber" })}`
			case "count":
				return `insert.pageCount@${String(at)} -> ${context.execute("insert.pageCount", { type: "insert.pageCount" })}`
			case "date": {
				const format = pick(context.rand, DATE_PICTURES)

				return `insert.date@${String(at)} ${format} -> ${context.execute("insert.date", { type: "insert.date", format })}`
			}
			default: {
				const instruction =
					choice === "seq"
						? "SEQ Figure \\* ARABIC"
						: choice === "ref"
							? `REF ${target} \\h`
							: choice === "pageref"
								? `PAGEREF ${target} \\h`
								: "SECTIONPAGES"

				return `insert.field@${String(at)} ${instruction} -> ${context.execute("insert.field", { type: "insert.field", instruction })}`
			}
		}
	}),
	paged("bookmark", 2, context => {
		const session = context.session
		const rand = context.rand
		const bookmarks = sessionBookmarks(session).list.filter(bookmark => !bookmark.hidden)
		const action = bookmarks.length === 0 ? "add" : pick(rand, ["add", "add", "rename", "delete"] as const)

		if (action === "add") {
			const a = caret(context)

			if (a === undefined) return "noop"

			const b = nearPosition(rand, session.state.doc, a) ?? a

			select(session, a, b)

			const name = pick(rand, NAMES)

			return `bookmark.add ${String(a)}..${String(b)} ${name} -> ${context.execute("bookmark.add", { type: "bookmark.add", name })}`
		}

		const bookmark = pick(rand, bookmarks)

		if (action === "rename") {
			const to = pick(rand, NAMES)

			return `bookmark.rename ${bookmark.name} ${to} -> ${context.execute("bookmark.rename", { type: "bookmark.rename", name: bookmark.name, to })}`
		}

		return `bookmark.delete ${bookmark.name} -> ${context.execute("bookmark.delete", { type: "bookmark.delete", name: bookmark.name })}`
	}),
	{ name: "fieldResult", harness: "damage", weight: 2, run: resultDamage },
	{ name: "fieldChar", harness: "damage", weight: 1, run: charDamage },
	{ name: "bookmarkRaw", harness: "damage", weight: 1, run: bookmarkDamage },
	{ name: "referenceRaw", harness: "damage", weight: 1, run: referenceDamage },
	{
		name: "fieldReferenceScripted",
		harness: "damage",
		weight: 0,
		run: referenceDamage,
		script: { order: 700, run: target => scriptedReference(target.session) }
	},
	{
		name: "fieldResultScripted",
		harness: "damage",
		weight: 0,
		run: resultDamage,
		script: { order: 710, run: target => scriptedResult(target.session) }
	}
]
