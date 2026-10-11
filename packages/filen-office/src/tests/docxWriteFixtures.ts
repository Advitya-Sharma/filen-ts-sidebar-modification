import type { Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { DocxEditorSession, docxSchema } from "@filen/office/core/docx/edit"
import { writerProof } from "@filen/office/core/docx/write"
import type { Opened } from "./docxEditFixtures"

// Scripted edits for writer tests: each picks its own paragraph and reports what it did.

export interface EditScript {
	readonly session: DocxEditorSession
	readonly done: string[]
}

function paragraphs(doc: Node, inTables: boolean): { node: Node; pos: number }[] {
	const out: { node: Node; pos: number }[] = []

	doc.descendants((node, pos, parent) => {
		if (node.type === docxSchema.nodes.paragraph) {
			const inCell = parent?.type === docxSchema.nodes.table_cell

			if (inCell === inTables && node.textContent.length >= 8 && node.firstChild?.isText === true) out.push({ node, pos })

			return false
		}

		return !node.isInline
	})

	return out
}

function select(session: DocxEditorSession, from: number, to: number = from): void {
	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, from, to)))
}

// The edit suite: type a word, delete a range, toggle bold, apply a heading, make a list, split and join
// paragraphs, edit a table cell, insert a link. Edits that find no suitable paragraph are skipped.
export function runEditScript(opened: Opened): EditScript {
	const session = new DocxEditorSession(opened.context, { proof: writerProof(opened.context) })
	const done: string[] = []
	const pick = (index: number, inTables = false): { node: Node; pos: number } | undefined => {
		const list = paragraphs(session.state.doc, inTables)

		return list[Math.min(list.length - 1, index)]
	}
	const step = (name: string, run: () => boolean): void => {
		if (run()) done.push(name)
	}

	step("type", () => {
		const target = pick(0)

		return target !== undefined && session.dispatch(session.state.tr.insertText("Filen ", target.pos + 1))
	})
	step("delete", () => {
		const target = pick(1)

		return target !== undefined && session.dispatch(session.state.tr.delete(target.pos + 2, target.pos + 5))
	})
	step("bold", () => {
		const target = pick(2)

		if (target === undefined) return false

		select(session, target.pos + 1, target.pos + 5)

		return session.execute({ type: "format.bold" })
	})
	step("heading", () => {
		const target = pick(3)

		if (target === undefined) return false

		select(session, target.pos + 1)

		return session.execute({ type: "style.set", style: "heading2" })
	})
	step("list", () => {
		const target = pick(4)

		if (target === undefined) return false

		select(session, target.pos + 1)

		return session.execute({ type: "list.bullet" })
	})
	step("split", () => {
		const target = pick(5)

		return target !== undefined && session.dispatch(session.state.tr.split(target.pos + 4))
	})
	step("join", () => {
		const list = paragraphs(session.state.doc, false)

		for (let k = 6; k + 1 < list.length; k++) {
			const first = list[k]
			const second = list[k + 1]

			if (first === undefined || second === undefined || first.pos + first.node.nodeSize !== second.pos) continue
			if (session.dispatch(session.state.tr.join(second.pos))) return true
		}

		return false
	})
	step("cell", () => {
		const target = pick(0, true)

		return target !== undefined && session.dispatch(session.state.tr.insertText("cell ", target.pos + 1))
	})
	step("link", () => {
		const target = pick(7)

		if (target === undefined) return false

		const link = docxSchema.marks.link.create({ target: "https://filen.io/" })

		return session.dispatch(session.state.tr.addMark(target.pos + 1, target.pos + 4, link))
	})

	return { session, done }
}
