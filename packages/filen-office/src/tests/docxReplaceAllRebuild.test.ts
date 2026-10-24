import { Fragment, Slice } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { Mapping, StepMap } from "prosemirror-transform"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, ReviewIndex, docxSchema, replaceAll } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { RebuildStep, nodesBetweenRanges, type Replace } from "../core/docx/edit/invariants/core"
import { openSpec } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"
import { pageDigest, pagedEditor, savedDigest, settleAll } from "./pagedFixtures"
import { budget, cpuMs } from "./timing"

// Replace All rebuilds blocks in few steps, each mapping positions as its replaces do: what lies between the
// matches (a comment, a tracked change, the caret) keeps its place for every reader that maps positions.

const r = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const BODY =
	p("foo first") +
	`<w:p><w:commentRangeStart w:id="0"/>${r("commented middle")}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>` +
	`<w:p><w:ins w:id="5" w:author="Bob" w:date="2025-01-01T00:00:00Z">${r("tracked middle")}</w:ins></w:p>` +
	p("foo last")
const COMMENTS = `<w:comment w:id="0" w:author="Bob" w:date="2025-01-01T00:00:00Z" w:initials="B">${p("c")}</w:comment>`

function replacing(result: ReturnType<typeof replaceAll>): NonNullable<Extract<ReturnType<typeof replaceAll>, { tr: unknown }>> {
	if (result === null || !("tr" in result)) throw new Error(`nothing replaced: ${JSON.stringify(result)}`)

	return result
}

describe("Replace All rebuilds", () => {
	for (const tracking of [false, true])
		for (const replacement of ["bar", "longer replacement", ""])
			it(`keep the review index's entries between the matches (${JSON.stringify(replacement)}${tracking ? ", tracked" : ""})`, async () => {
				const opened = await openSpec({ body: BODY, comments: COMMENTS })

				opened.context.identity.author = "Ann"
				opened.context.review.tracking = tracking

				const session = new DocxEditorSession(opened.context)

				expect(session.reviewIndex.entries().length).toBeGreaterThan(0)
				expect(session.dispatch(replacing(replaceAll(session.state, "foo", replacement)).tr)).toBe(true)
				expect(session.reviewIndex.entries().map(entry => ({ ...entry }))).toEqual(
					new ReviewIndex(session.state.doc).entries().map(entry => ({ ...entry }))
				)

				const saved = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })

				expect(saved.bytes).toBeInstanceOf(Uint8Array)
			})

	it("save after a Replace All in the paged editor with a comment between matches", async () => {
		const opened = await pagedEditor(buildDocx({ body: BODY, comments: COMMENTS }))
		const { editor } = opened

		await settleAll(editor)
		editor.snapshot()
		expect(editor.dispatch(replacing(replaceAll(editor.session.state, "foo", "bar")).tr)).toBe(true)
		await settleAll(editor)
		expect(editor.session.reviewIndex.commentCount).toBe(1)

		const saved = (await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: editor.sessions })).bytes as Uint8Array

		await settleAll(editor)
		expect(pageDigest(editor.layout)).toEqual(await savedDigest(saved))
	})

	it("are an undo step of their own: typing right after in the span is undone alone", async () => {
		const opened = await openSpec({ body: p("foo a") + p("middle untouched") + p("foo b") })
		const session = new DocxEditorSession(opened.context)

		session.dispatch(replacing(replaceAll(session.state, "foo", "bar")).tr.setTime(1000))

		const replaced = session.state.doc

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, 12)))
		session.dispatch(session.state.tr.insertText("X").setTime(1100))
		expect(session.execute({ type: "history.undo" })).toBe(true)
		expect(session.state.doc.eq(replaced)).toBe(true)
	})

	it("walk ranges at blocks' own positions in linear time", () => {
		const blocks = Array.from({ length: 40_000 }, (_, k) => docxSchema.nodes.paragraph.create(null, docxSchema.text(`p${String(k)}`)))
		const doc = docxSchema.nodes.doc.create(null, blocks)
		const ranges: number[] = []
		let visits = 0

		doc.forEach((_, pos) => ranges.push(pos, pos + 1))

		const start = cpuMs()

		nodesBetweenRanges(doc, ranges, () => {
			visits++
		})

		expect(visits).toBe(40_000)
		expect(cpuMs() - start).toBeLessThan(budget(250))
	})
})

// A rebuild's map against ProseMirror's StepMap over the same ranges: every position, both sides, forward and
// inverted, and recovery through a mirrored Mapping (as history's remapping pairs a map with its inverse).
const inserted = (size: number): Slice => (size === 0 ? Slice.empty : new Slice(Fragment.from(docxSchema.text("x".repeat(size))), 0, 0))

function maps(spec: readonly (readonly [number, number, number])[]): { rebuild: StepMap; plain: StepMap } {
	const replaces: Replace[] = spec.map(([from, oldSize, newSize]) => ({ from, to: from + oldSize, slice: inserted(newSize) }))

	return { rebuild: new RebuildStep(0, 100, Slice.empty, replaces).getMap(), plain: new StepMap(spec.flat()) }
}

function described(map: StepMap, pos: number, assoc: number): unknown {
	const result = map.mapResult(pos, assoc)

	return {
		pos: result.pos,
		deleted: result.deleted,
		before: result.deletedBefore,
		after: result.deletedAfter,
		across: result.deletedAcross,
		map: map.map(pos, assoc)
	}
}

// StepMap's recovery value and touches() (ProseMirror's own, not in its types).
function recovery(map: StepMap, pos: number): { readonly recover: number | null; readonly touches: boolean | null } {
	const recover = Reflect.get(map.mapResult(pos, 1), "recover") as number | null
	const touches = Reflect.get(map, "touches") as (pos: number, recover: number) => boolean

	return { recover, touches: recover === null ? null : touches.call(map, pos, recover) }
}

const SPECS: (readonly [number, number, number])[][] = [
	[[5, 3, 1]],
	[
		[2, 3, 0],
		[7, 3, 0]
	],
	[
		[0, 5, 0],
		[5, 5, 0]
	],
	[
		[4, 0, 3],
		[4, 2, 0]
	],
	[
		[3, 2, 4],
		[9, 0, 2],
		[9, 1, 1],
		[15, 4, 0]
	],
	[
		[1, 1, 1],
		[2, 1, 1],
		[3, 1, 1]
	],
	[
		[10, 5, 20],
		[30, 0, 0],
		[40, 10, 2]
	]
]

describe("a rebuild's map", () => {
	it.each(SPECS.map(spec => [JSON.stringify(spec), spec] as const))("maps as a StepMap over its replaces: %s", (_name, spec) => {
		const { rebuild, plain } = maps(spec)

		for (const [a, b] of [
			[rebuild, plain],
			[rebuild.invert(), plain.invert()]
		] as const)
			for (let pos = 0; pos <= 60; pos++)
				for (const assoc of [-1, 1]) expect([pos, assoc, described(a, pos, assoc)]).toEqual([pos, assoc, described(b, pos, assoc)])
	})

	it.each(SPECS.map(spec => [JSON.stringify(spec), spec] as const))(
		"recovers as a StepMap through a mirrored Mapping: %s",
		(_name, spec) => {
			const { rebuild, plain } = maps(spec)
			const mirrored = (forward: StepMap): Mapping => {
				const mapping = new Mapping()

				mapping.appendMap(forward)
				mapping.appendMap(forward.invert(), 0)

				return mapping
			}

			for (let pos = 0; pos <= 60; pos++) {
				for (const assoc of [-1, 1])
					expect([pos, assoc, mirrored(rebuild).map(pos, assoc)]).toEqual([pos, assoc, mirrored(plain).map(pos, assoc)])

				expect([pos, recovery(rebuild, pos)]).toEqual([pos, recovery(plain, pos)])
			}
		}
	)
})
