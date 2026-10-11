import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import {
	STORY_FIND_LIMIT,
	findInStory,
	findPattern,
	matchHolds,
	replaceAll,
	replaceMatches,
	selectMatch
} from "@filen/office/core/docx/edit"
import { nodeText, sessionOf } from "./docxEditFixtures"
import { p } from "./docxSynthetic"

const BOLD_RUN = '<w:p><w:r><w:t xml:space="preserve">plain </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r></w:p>'

describe("findInStory", () => {
	it("finds matches across paragraphs at document positions", async () => {
		const session = await sessionOf({ body: p("One fish two fish") + p("red Fish") })
		const doc = session.state.doc
		const matches = findInStory(doc, "fish")

		expect(matches).toHaveLength(3)

		for (const match of matches) expect(doc.textBetween(match.from, match.to).toLowerCase()).toBe("fish")

		expect(findInStory(doc, "Fish", { caseSensitive: true })).toHaveLength(1)
	})

	it("treats the query literally and an empty one as nothing", async () => {
		const session = await sessionOf({ body: p("a.b a+b (x)") })
		const doc = session.state.doc

		expect(findInStory(doc, ".")).toHaveLength(1)
		expect(findInStory(doc, "a+b")).toHaveLength(1)
		expect(findInStory(doc, "(x)")).toHaveLength(1)
		expect(findInStory(doc, "")).toEqual([])
		expect(findPattern("")).toBeNull()
	})

	it("never matches across an inline atom", async () => {
		const session = await sessionOf({ body: "<w:p><w:r><w:t>ab</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>cd</w:t></w:r></w:p>" })
		const doc = session.state.doc

		expect(findInStory(doc, "bc")).toEqual([])
		expect(findInStory(doc, "cd")).toHaveLength(1)

		const [match] = findInStory(doc, "cd")

		expect(match === undefined ? "" : doc.textBetween(match.from, match.to)).toBe("cd")
	})

	it("matches across normalisation forms on whole characters", async () => {
		const session = await sessionOf({ body: p("Le CAFE\u0301 et le cafe\u0301") })
		const doc = session.state.doc
		const matches = findInStory(doc, "café")

		expect(matches.map(match => doc.textBetween(match.from, match.to))).toEqual(["CAFE\u0301", "cafe\u0301"])
		expect(matches.every(match => matchHolds(doc, match, "café"))).toBe(true)
	})

	it("honours the limit", async () => {
		const session = await sessionOf({ body: p("x ".repeat(50)) })

		expect(findInStory(session.state.doc, "x", {}, 7)).toHaveLength(7)
	})
})

describe("replaceMatches", () => {
	it("replaces every match in one undoable step, keeping the formatting where each starts", async () => {
		const session = await sessionOf({ body: BOLD_RUN + p("plain again") })
		const matches = findInStory(session.state.doc, "plain")
		const tr = replaceMatches(session.state, matches, "plain", "simple")

		expect(tr).not.toBeNull()

		if (tr === null) return

		expect(session.dispatch(tr)).toBe(true)
		expect(nodeText(session.state.doc.child(0))).toBe("simple Bold")
		expect(nodeText(session.state.doc.child(1))).toBe("simple again")
		expect(session.snapshot().dirty).toBe(true)
		expect(session.execute({ type: "history.undo" })).toBe(true)
		expect(nodeText(session.state.doc.child(0))).toBe("plain Bold")
		expect(nodeText(session.state.doc.child(1))).toBe("plain again")
	})

	it("keeps the run's formatting of the match", async () => {
		const session = await sessionOf({ body: BOLD_RUN })
		const matches = findInStory(session.state.doc, "Bold")
		const tr = replaceMatches(session.state, matches, "Bold", "Strong")

		if (tr === null) throw new Error("no replacement")

		session.dispatch(tr)

		const [match] = findInStory(session.state.doc, "Strong")

		if (match === undefined) throw new Error("not replaced")

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, match.from, match.to)))
		expect(session.snapshot().formats.bold).toBe("on")
	})

	it("deletes with an empty replacement and skips matches that no longer hold", async () => {
		const session = await sessionOf({ body: p("cat cat cat") })
		const matches = findInStory(session.state.doc, "cat")

		expect(matches).toHaveLength(3)

		// The middle one is edited away first.
		const middle = matches[1]

		if (middle === undefined) throw new Error("no match")

		session.dispatch(session.state.tr.insertText("dog", middle.from, middle.to))
		expect(matchHolds(session.state.doc, middle, "cat")).toBe(false)

		const tr = replaceMatches(session.state, matches, "cat", "")

		if (tr === null) throw new Error("nothing replaced")

		session.dispatch(tr)
		expect(nodeText(session.state.doc.child(0))).toBe(" dog ")
	})

	it("returns null when nothing holds and refuses in a read-only session", async () => {
		const session = await sessionOf({ body: p("abc") }, { readOnly: true })
		const matches = findInStory(session.state.doc, "abc")

		expect(replaceMatches(session.state, [], "abc", "x")).toBeNull()

		const tr = replaceMatches(session.state, matches, "abc", "x")

		if (tr === null) throw new Error("no transaction")

		expect(session.dispatch(tr)).toBe(false)
		expect(nodeText(session.state.doc.child(0))).toBe("abc")
	})
})

describe("replaceAll", () => {
	it("replaces every match, past what find shows, in one undoable step with the true count", async () => {
		const session = await sessionOf({ body: p("ab ".repeat(6000)) + BOLD_RUN + p("ab ".repeat(6000)) })

		expect(findInStory(session.state.doc, "ab")).toHaveLength(STORY_FIND_LIMIT)

		const result = replaceAll(session.state, "ab", "xyz")

		if (result === null) throw new Error("nothing replaced")

		expect(result.count).toBe(12_000)
		expect(result.tr.steps).toHaveLength(2)
		expect(session.dispatch(result.tr)).toBe(true)
		expect(findInStory(session.state.doc, "ab")).toHaveLength(0)
		expect(findInStory(session.state.doc, "xyz", {}, Number.POSITIVE_INFINITY)).toHaveLength(12_000)
		expect(nodeText(session.state.doc.child(1))).toBe("plain Bold")
		expect(session.execute({ type: "history.undo" })).toBe(true)
		expect(findInStory(session.state.doc, "ab", {}, Number.POSITIVE_INFINITY)).toHaveLength(12_000)
	})

	it("keeps each match's formatting, deletes with an empty replacement, and is null without a match", async () => {
		const session = await sessionOf({ body: BOLD_RUN + p("cat cat") })
		const bold = replaceAll(session.state, "Bold", "Strong")

		if (bold === null) throw new Error("no replacement")

		session.dispatch(bold.tr)

		const [match] = findInStory(session.state.doc, "Strong")

		if (match === undefined) throw new Error("not replaced")

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, match.from, match.to)))
		expect(session.snapshot().formats.bold).toBe("on")

		const cats = replaceAll(session.state, "CAT", "")

		if (cats === null) throw new Error("nothing deleted")

		expect(cats.count).toBe(2)
		session.dispatch(cats.tr)
		expect(nodeText(session.state.doc.child(1))).toBe(" ")
		expect(replaceAll(session.state, "dog", "x")).toBeNull()
		expect(replaceAll(session.state, "CAT", "x", { caseSensitive: true })).toBeNull()
	})
})

describe("selectMatch", () => {
	it("selects a match and refuses one outside the document", async () => {
		const session = await sessionOf({ body: p("find me here") })
		const [match] = findInStory(session.state.doc, "me")

		if (match === undefined) throw new Error("no match")

		const tr = selectMatch(session.state, match)

		expect(tr).not.toBeNull()

		if (tr !== null) session.dispatch(tr)

		expect(session.state.selection.from).toBe(match.from)
		expect(session.state.selection.to).toBe(match.to)
		expect(selectMatch(session.state, { from: 0, to: session.state.doc.content.size + 5 })).toBeNull()
	})
})

describe("markSaved", () => {
	it("marks the document as it was written saved, leaving later edits unsaved", async () => {
		const session = await sessionOf({ body: p("abc") })

		session.dispatch(session.state.tr.insertText("x", 1))

		const written = session.state.doc

		session.dispatch(session.state.tr.insertText("y", 1))
		session.markSaved(written)
		expect(session.snapshot().dirty).toBe(true)
		session.markSaved()
		expect(session.snapshot().dirty).toBe(false)
	})
})
