import { describe, expect, it } from "vitest"
import { DOCX_EDITOR_COMMANDS, NO_REVIEW, type DocxEditorCommand } from "@filen/office/core/docx/edit"
import { commandNeedsAuthor, commandRecordsAuthor } from "@/features/preview/components/docxAuthorGate.logic"

describe("docx author gate logic", () => {
	it("asks for the author only for commands that will record one", () => {
		for (const command of DOCX_EDITOR_COMMANDS) {
			if (command.author !== true) {
				expect(commandRecordsAuthor({ type: command.type } as never, false), command.type).toBe(false)
			}
		}

		expect(commandRecordsAuthor({ type: "comment.add", text: "x" }, false)).toBe(true)
		expect(commandRecordsAuthor({ type: "comment.reply", id: "1", text: "x" }, true)).toBe(true)
		expect(commandRecordsAuthor({ type: "comment.edit", id: "1", text: "x" }, true)).toBe(true)
		expect(commandRecordsAuthor({ type: "comment.delete", id: "1" }, true)).toBe(false)
		// Tracking can always be turned off without a name.
		expect(commandRecordsAuthor({ type: "review.track", on: true }, false)).toBe(true)
		expect(commandRecordsAuthor({ type: "review.track", on: false }, true)).toBe(false)
		expect(commandRecordsAuthor({ type: "review.track" }, false)).toBe(true)
		expect(commandRecordsAuthor({ type: "review.track" }, true)).toBe(false)
	})

	it("asks before a command in the document's review state exactly when the command records the author", () => {
		const commands: DocxEditorCommand[] = [
			{ type: "comment.add", text: "x" },
			{ type: "review.track" },
			{ type: "review.track", on: false },
			{ type: "format.bold" }
		]

		for (const tracking of [false, true]) {
			const review = { ...NO_REVIEW, tracking }

			for (const command of commands) {
				expect(commandNeedsAuthor(command, review), `${command.type} ${String(tracking)}`).toBe(
					commandRecordsAuthor(command, tracking)
				)
			}
		}
	})
})

describe("docx author gate while tracking without a name", () => {
	it("asks before every write that would be tracked, never before one that is not", () => {
		const review = { ...NO_REVIEW, tracking: true, authorNeeded: true }

		expect(commandNeedsAuthor({ type: "format.bold" }, review)).toBe(true)
		expect(commandNeedsAuthor({ type: "insert.pageBreak" }, review)).toBe(true)
		// Turning tracking off, undoing, resolving changes and a comment's state never record a change.
		expect(commandNeedsAuthor({ type: "review.track", on: false }, review)).toBe(false)
		expect(commandNeedsAuthor({ type: "history.undo" }, review)).toBe(false)
		expect(commandNeedsAuthor({ type: "review.acceptAll" }, review)).toBe(false)
		expect(commandNeedsAuthor({ type: "comment.resolve", id: "1" }, review)).toBe(false)
		expect(commandNeedsAuthor({ type: "comment.delete", id: "1" }, review)).toBe(false)
		// What changes nothing asks nothing.
		expect(commandNeedsAuthor({ type: "comment.next" }, review)).toBe(false)
		// With a name, tracked writes run as they are.
		expect(commandNeedsAuthor({ type: "format.bold" }, { ...review, authorNeeded: false })).toBe(false)
	})
})
