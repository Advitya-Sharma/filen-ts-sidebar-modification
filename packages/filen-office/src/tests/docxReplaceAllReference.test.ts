import { undo } from "prosemirror-history"
import { afterEach, describe, expect, it, vi } from "vitest"
import { DocxEditorSession, findInStory, replaceAll, replaceMatches } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { openSpec } from "./docxEditFixtures"
import { p, type SyntheticDocx } from "./docxSynthetic"

// Replace All against replacing each match on its own (replaceMatches), through the same session (invariants and
// tracking included): the documents are equal, the saves byte-identical, and one undo restores the original.
// The two documents are opened seconds apart (as on a loaded machine): nothing in the saves may follow the clock.

const r = (text: string, rPr = ""): string =>
	`<w:r>${rPr === "" ? "" : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`
const cell = (inner: string, tcPr = ""): string => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr>${inner}</w:tc>`
const table = (rows: string): string =>
	`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>`

const BODY =
	p("foo one foo") +
	`<w:p>${r("a foo ")}${r("bold foo", "<w:b/>")}${r(" foo end")}</w:p>` +
	`<w:p><w:bookmarkStart w:id="1" w:name="bm"/>${r("foo in bookmark")}<w:bookmarkEnd w:id="1"/>${r(" then foo")}</w:p>` +
	`<w:p><w:commentRangeStart w:id="0"/>${r("commented foo")}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>${r(" foo")}</w:p>` +
	`<w:p><w:hyperlink w:anchor="bm">${r("link foo")}</w:hyperlink>${r(" foo")}</w:p>` +
	`<w:p>${r("x foo")}<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>${r("foo y")}</w:p>` +
	`<w:p>${r("note foo")}<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>${r(" foo")}</w:p>` +
	table(
		`<w:tr>${cell(p("cell foo a"), '<w:vMerge w:val="restart"/>')}${cell(p("foo b"))}</w:tr>` +
			`<w:tr>${cell(p(""), "<w:vMerge/>")}${cell(p("foo c") + table(`<w:tr>${cell(p("nested foo"))}${cell(p("x"))}</w:tr>`) + p("after nested foo"))}</w:tr>`
	) +
	`<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${p("block control foo")}${p("second foo")}</w:sdtContent></w:sdt>` +
	`<w:sdt><w:sdtPr><w:id w:val="8"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>${p("locked control foo")}</w:sdtContent></w:sdt>` +
	`<w:p>${r("before ")}<w:sdt><w:sdtPr><w:id w:val="9"/><w:text/></w:sdtPr><w:sdtContent>${r("plain foo")}</w:sdtContent></w:sdt>${r(" after foo")}</w:p>` +
	`<w:p><w:ins w:id="20" w:author="Bob" w:date="2025-01-01T00:00:00Z">${r("inserted foo")}</w:ins>${r(" foo")}<w:del w:id="21" w:author="Bob" w:date="2025-01-01T00:00:00Z"><w:r><w:delText>deleted foo</w:delText></w:r></w:del></w:p>` +
	`<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="22" w:author="Bob" w:date="2025-01-01T00:00:00Z"><w:pPr/></w:pPrChange></w:pPr>${r("changed foo")}</w:p>` +
	p("plain tail foo") +
	table(`<w:tr>${cell(p("last table foo"))}${cell(p("y"))}</w:tr>`) +
	p("end")

const SPEC: SyntheticDocx = {
	body: BODY,
	comments: `<w:comment w:id="0" w:author="Bob" w:date="2025-01-01T00:00:00Z" w:initials="B">${p("comment text foo")}</w:comment>`,
	footnotes:
		'<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
		`<w:footnote w:id="1">${p("footnote foo")}</w:footnote>`
}

async function session(tracking: boolean): Promise<{ s: DocxEditorSession; opened: Awaited<ReturnType<typeof openSpec>> }> {
	const opened = await openSpec(SPEC)

	opened.context.identity.author = "Ann"
	opened.context.identity.now = () => Date.parse("2026-01-01T10:00:00Z")
	opened.context.review.tracking = tracking

	return { s: new DocxEditorSession(opened.context), opened }
}

describe("Replace All against one replace per match", () => {
	afterEach(() => {
		vi.useRealTimers()
	})

	for (const tracking of [false, true])
		for (const [query, replacement] of [
			["foo", "BAR"],
			["foo", ""],
			["o", "oo"],
			["foo", "foo foo"],
			[" ", "_"]
		] as const)
			it(`${query} -> ${JSON.stringify(replacement)} (tracking ${String(tracking)})`, async () => {
				vi.useFakeTimers({ toFake: ["Date"] })

				const a = await session(tracking)

				vi.setSystemTime(Date.now() + 3_000)

				const b = await session(tracking)
				const original = a.s.state.doc
				const all = replaceAll(a.s.state, query, replacement)

				if (all === null || !("tr" in all)) throw new Error(`nothing replaced: ${JSON.stringify(all)}`)

				expect(a.s.dispatch(all.tr)).toBe(true)

				const each = replaceMatches(b.s.state, findInStory(b.s.state.doc, query, {}, Number.POSITIVE_INFINITY), query, replacement)

				if (each === null) throw new Error("nothing replaced one by one")

				expect(b.s.dispatch(each)).toBe(true)

				expect(a.s.state.doc.toJSON()).toEqual(b.s.state.doc.toJSON())

				const sa = (await saveDocx({ pkg: a.opened.pkg, context: a.opened.context, sessions: [a.s] })).bytes as Uint8Array
				const sb = (await saveDocx({ pkg: b.opened.pkg, context: b.opened.context, sessions: [b.s] })).bytes as Uint8Array

				expect(Buffer.from(sa).equals(Buffer.from(sb))).toBe(true)

				expect(undo(a.s.state, tr => a.s.dispatch(tr))).toBe(true)
				expect(a.s.state.doc.eq(original)).toBe(true)
			})
})
