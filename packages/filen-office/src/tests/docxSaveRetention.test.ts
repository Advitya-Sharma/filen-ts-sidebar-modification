import { setFlagsFromString } from "node:v8"
import { runInNewContext } from "node:vm"
import { EditorState, TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { saveDocx } from "@filen/office/core/docx/write"
import { buildDocx } from "./docxSynthetic"
import { pagedEditor } from "./pagedFixtures"

setFlagsFromString("--expose-gc")

const gc = runInNewContext("gc") as () => void

function arrayBuffersMb(): number {
	gc()
	gc()

	return process.memoryUsage().arrayBuffers / 1048576
}

const cell = (text: string): string => `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`

describe("saves of a long session", () => {
	it("keep no earlier save's written story alive", async () => {
		const rows = Array.from(
			{ length: 1500 },
			(_, r) => `<w:tr>${cell(`row ${String(r)} lorem ipsum dolor sit amet`)}${cell(`b${String(r)}`)}</w:tr>`
		)
		const bytes = buildDocx({
			body: `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${rows.join("")}</w:tbl><w:p/>`
		})
		const { editor, pkg, context } = await pagedEditor(bytes)
		const session = editor.body.session
		const at = (text: string): number => {
			let found = -1

			session.state.doc.descendants((node, pos) => {
				if (found < 0 && node.isTextblock && node.textContent.startsWith(text)) found = pos + 1

				return found < 0
			})

			return found
		}
		const save = async (): Promise<void> => {
			await saveDocx({ pkg, context, sessions: editor.sessions })
		}
		// Each save writes the whole table again (a column was added) with one paragraph new to the read back.
		const edit = async (k: number): Promise<void> => {
			session.dispatch(session.state.tr.insertText(`edited ${String(k)} `, at(`row ${String(k * 97)} `)))
			await editor.whenSettled()

			// Without the undo history, which holds each version of the table.
			const state = session.state

			Reflect.set(session, "current", EditorState.create({ doc: state.doc, selection: state.selection, plugins: state.plugins }))
		}

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at("row 0 "))))
		await editor.execute({ type: "table.column.insertRight" })
		await editor.whenSettled()
		await edit(0)
		await save()
		await edit(1)
		await save()

		const before = arrayBuffersMb()

		for (let k = 2; k < 12; k++) {
			await edit(k)
			await save()
		}

		// Ten saves more: what the last of them read back, not a whole table's markup per save.
		expect(arrayBuffersMb() - before).toBeLessThan(1)
		editor.dispose()
	})
})
