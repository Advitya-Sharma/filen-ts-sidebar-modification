import { describe, expect, it } from "vitest"
import { RevisionView, TrackedLook } from "@filen/office/core/docx"
import { buildDocx } from "./docxSynthetic"
import { layOut } from "./layoutFixtures"

// Tracked changes in the Markup view are drawn in their author's colour: text styles carry the author.

describe("revision authors in text styles", () => {
	it("gives each author's insertions and deletions their own colour slot", async () => {
		const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
		const deleted = (text: string): string => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`
		const body =
			`<w:p>${run("Kept ")}<w:ins w:id="1" w:author="Ada" w:date="2026-01-01T00:00:00Z">${run("added by Ada ")}</w:ins>` +
			`<w:del w:id="2" w:author="Grace" w:date="2026-01-01T00:00:00Z">${deleted("removed by Grace ")}</w:del>` +
			`<w:ins w:id="3" w:author="Ada" w:date="2026-01-01T00:00:00Z">${run("Ada again")}</w:ins><w:ins w:id="4">${run(" anonymous")}</w:ins></w:p>`
		const layout = await layOut(buildDocx({ body }), { view: RevisionView.Markup })
		const slots = new Map<number, Set<number>>()

		layout.styleInfo.forEach(info => {
			if (info.revision === TrackedLook.NONE) return

			const set = slots.get(info.revision) ?? new Set<number>()

			set.add(info.author)
			slots.set(info.revision, set)
		})

		// Ada first (slot 1), Grace next (2), a change without an author in the last slot (8).
		expect([...(slots.get(TrackedLook.INSERTED) ?? [])].sort()).toEqual([1, 8])
		expect([...(slots.get(TrackedLook.DELETED) ?? [])]).toEqual([2])
	})
})
