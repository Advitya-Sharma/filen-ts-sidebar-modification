import { describe, expect, it } from "vitest"
import { MAX_NOTE_TEXT } from "../core/xlsx/comments/edit"
import { savedBytes } from "./sheetSessionFixtures"
import { NOTES, changeOf, entries, openAnnotated, settle } from "./sheetNotesFixtures"

describe("a thread's placeholder note", () => {
	it("stays within a note's text limit however long the thread grows, never splitting a character", async () => {
		const { session, sheets } = await openAnnotated(NOTES)
		const sheet = sheets[1] ?? 0

		const each = Math.floor(MAX_NOTE_TEXT / 2)

		await settle()

		for (let k = 0; k < 3; k++) changeOf(await session.command({ type: "thread.add", sheet, row: 0, col: 0, text: "😀".repeat(each) }))

		const files = entries(await savedBytes(session))
		const text = [...(files["xl/comments2.xml"] ?? "").matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(match => match[1] ?? "").join("")
		const threaded = Object.entries(files).find(([name]) => name.startsWith("xl/threadedComments/"))?.[1] ?? ""

		expect(text.length).toBeLessThanOrEqual(MAX_NOTE_TEXT)
		expect(text.length).toBeGreaterThan(MAX_NOTE_TEXT - 2)
		expect(text).not.toMatch(/[\ud800-\udbff]$/)
		expect([...threaded.matchAll(/😀/g)].length).toBe(3 * each)
		session.close()
	})
})
