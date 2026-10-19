import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { key, openEditor, placeCaret, selectAll, spreadPosition } from "../bench/editHarness"
import { buildDocx, p } from "./docxSynthetic"

// The layout draws table, row and cell properties without their tracked-change stamps (PropsStore's looks):
// each tracked table edit writes new stamps, which would otherwise intern new properties for good. Shadows of
// overlays (their structure) intern theirs apart from the document.

describe("the property store over tracked table edits", () => {
	it("keeps the store flat over tracked row operations undone", async () => {
		const cell = (text: string): string => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`
		const rows = Array.from(
			{ length: 6 },
			(_, k) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(`a${String(k)}`)}${cell(`b${String(k)}`)}</w:tr>`
		)
		const opened = await openEditor(
			buildDocx({
				body: `${p("before")}<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.join("")}</w:tbl>${p("after")}`
			})
		)
		const editor = opened.editor
		const store = editor.session.context.store
		const counts: number[] = []

		editor.execute({ type: "review.track", on: true })

		for (let round = 0; round < 3; round++) {
			const doc = editor.session.state.doc
			let at = -1

			doc.descendants((node, pos) => {
				if (at < 0 && node.type.name === "table_cell") at = pos + 2

				return at < 0
			})
			editor.dispatch(editor.session.state.tr.setSelection(TextSelection.near(doc.resolve(at))))
			editor.execute({ type: "table.row.insertBelow" })
			await editor.whenSettled()
			editor.execute({ type: "history.undo" })
			await editor.whenSettled()

			const stats = store.stats()

			counts.push(stats.tables + stats.rows + stats.cells + stats.looks)
		}

		opened.close()
		expect(counts[2]).toBe(counts[0])
	})

	it("keeps the run bases and store flat while typing in tracked formatting", async () => {
		const body = Array.from({ length: 30 }, (_, k) => p(`Paragraph ${String(k)} with words to type into`)).join("")
		const opened = await openEditor(buildDocx({ body }))
		const editor = opened.editor
		const context = editor.session.context

		editor.execute({ type: "review.track", on: true })
		selectAll(editor)
		editor.execute({ type: "format.bold" })
		await editor.whenSettled()

		const typeAt = async (k: number): Promise<void> => {
			placeCaret(editor, spreadPosition(editor.session.state, k))
			key(editor, "type")
			await editor.whenSettled()
		}

		// The first key makes the run base of its tracked formatting.
		await typeAt(0)

		const bases = context.bases.size
		const runs = context.store.stats().runs

		for (let k = 1; k < 20; k++) await typeAt(k)

		opened.close()
		expect(context.bases.size).toBe(bases)
		expect(context.store.stats().runs).toBe(runs)
	})
})
