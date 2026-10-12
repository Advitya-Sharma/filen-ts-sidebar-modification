import { TextSelection } from "prosemirror-state"
import { AddMarkStep, Step } from "prosemirror-transform"
import { describe, expect, it } from "vitest"
import { docxSchema } from "@filen/office/core/docx/edit"
import { SetMarkSetsStep, SetMarksStep } from "../core/docx/edit/marks"
import { sessionOf } from "./docxEditFixtures"
import { p } from "./docxSynthetic"

// Formatting a selection is one step whatever the runs it covers, and undo gives every run back the mark it had.

const marks = docxSchema.marks
const run = (text: string, rPr: string): string => `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`
const pr = (runs: string): string => `<w:p>${runs}</w:p>`

function boldOf(doc: import("prosemirror-model").Node): string[] {
	const out: string[] = []

	doc.descendants(node => {
		if (node.isText) out.push(`${node.text ?? ""}:${JSON.stringify(marks.bold.isInSet(node.marks)?.attrs ?? null)}`)

		return true
	})

	return out
}

describe("setting marks piece by piece", () => {
	it("bolds a selection of differently bold runs in one step, and undo restores each run's own mark", async () => {
		const session = await sessionOf({
			body:
				pr(run("plain ", "") + run("bcs ", "<w:bCs/>") + run("both ", "<w:b/><w:bCs/>") + run("b ", "<w:b/>")) +
				pr(run("next ", "") + run("bcs", "<w:bCs/>")) +
				p("untouched")
		})
		const doc = session.state.doc
		const before = boldOf(doc)
		const untouched = doc.child(2)

		session.dispatch(session.state.tr.setSelection(TextSelection.create(doc, 1, doc.child(0).nodeSize + doc.child(1).nodeSize - 1)))

		const steps: Step[] = []
		const stop = session.onApplied(transactions => {
			for (const tr of transactions) steps.push(...tr.steps)
		})

		expect(session.execute({ type: "format.bold" })).toBe(true)
		stop()
		expect(steps.filter(step => step instanceof SetMarksStep)).toHaveLength(1)
		expect(
			boldOf(session.state.doc)
				.slice(0, 6)
				.every(entry => entry.endsWith(':{"value":true,"cs":true}'))
		).toBe(true)
		// A block no piece is in keeps its node.
		expect(session.state.doc.child(2)).toBe(untouched)

		expect(session.execute({ type: "history.undo" })).toBe(true)
		expect(boldOf(session.state.doc)).toEqual(before)

		expect(session.execute({ type: "history.redo" })).toBe(true)
		expect(
			boldOf(session.state.doc)
				.slice(0, 6)
				.every(entry => entry.endsWith(':{"value":true,"cs":true}'))
		).toBe(true)
	})

	it("maps through later changes, never merges, and round-trips through JSON", async () => {
		const session = await sessionOf({ body: pr(run("one two", "<w:b/>")) })
		const doc = session.state.doc
		const bold = marks.bold.create({ value: true, cs: true })
		const step = new SetMarksStep(
			[
				{ from: 1, to: 4, mark: bold },
				{ from: 5, to: 8, mark: null }
			],
			bold
		)
		const inverted = step.invert(doc)
		const applied = step.apply(doc).doc

		expect(applied).not.toBeNull()
		expect(inverted.apply(applied as typeof doc).doc?.eq(doc)).toBe(true)
		expect(step.merge()).toBeNull()

		// Text inserted before the pieces moves them.
		const moved = session.state.tr.insertText("xx", 1)
		const mapped = step.map(moved.mapping) as SetMarksStep

		expect(mapped.pieces.map(piece => [piece.from, piece.to])).toEqual([
			[3, 6],
			[7, 10]
		])

		const json = Step.fromJSON(docxSchema, step.toJSON()) as SetMarksStep

		expect(json).toBeInstanceOf(SetMarksStep)
		expect(json.apply(doc).doc?.eq(applied as typeof doc)).toBe(true)
	})

	it("clears a dozen mark types in one step, which undo and redo restore exactly", async () => {
		const session = await sessionOf({
			body:
				pr(
					run("one ", "<w:b/><w:i/>") +
						run("two ", '<w:sz w:val="30"/><w:color w:val="FF0000"/>') +
						run("three", '<w:highlight w:val="yellow"/>')
				) + pr(run("four", '<w:u w:val="single"/><w:lang w:val="de-DE"/>'))
		})
		const before = session.state.doc

		session.dispatch(session.state.tr.setSelection(TextSelection.create(before, 1, before.content.size - 1)))

		const steps: Step[] = []
		const stop = session.onApplied(transactions => {
			for (const tr of transactions) steps.push(...tr.steps)
		})

		expect(session.execute({ type: "format.clearAll" })).toBe(true)
		stop()
		expect(steps.filter(step => step instanceof AddMarkStep)).toHaveLength(1)
		expect(steps.find(step => step instanceof SetMarkSetsStep)).toBeDefined()

		const after = session.state.doc

		expect(session.execute({ type: "history.undo" })).toBe(true)
		expect(session.state.doc.eq(before)).toBe(true)
		expect(session.execute({ type: "history.redo" })).toBe(true)
		expect(session.state.doc.eq(after)).toBe(true)
	})

	it("sets mark sets that map, never merge, invert and round-trip through JSON", async () => {
		const session = await sessionOf({ body: pr(run("one ", "<w:b/>") + run("two", "<w:i/>")) })
		const doc = session.state.doc
		const bold = marks.bold.create({ value: true, cs: null })
		const italic = marks.italic.create({ value: true, cs: null })
		const step = new SetMarkSetsStep(
			[marks.bold, marks.italic],
			[
				{ from: 1, to: 3, marks: [italic] },
				{ from: 5, to: 8, marks: [bold] }
			],
			bold
		)
		const applied = step.apply(doc).doc

		expect(applied).not.toBeNull()
		expect(marks.italic.isInSet(applied?.nodeAt(1)?.marks ?? [])).toBeDefined()
		expect(marks.bold.isInSet(applied?.nodeAt(1)?.marks ?? [])).toBeUndefined()
		expect(marks.bold.isInSet(applied?.nodeAt(5)?.marks ?? [])).toBeDefined()
		expect(
			step
				.invert(doc)
				.apply(applied as typeof doc)
				.doc?.eq(doc)
		).toBe(true)
		expect(step.merge()).toBeNull()
		expect(
			(step.map(session.state.tr.insertText("xx", 1).mapping) as SetMarkSetsStep).pieces.map(piece => [piece.from, piece.to])
		).toEqual([
			[3, 5],
			[7, 10]
		])

		const json = Step.fromJSON(docxSchema, step.toJSON()) as SetMarkSetsStep

		expect(json).toBeInstanceOf(SetMarkSetsStep)
		expect(json.apply(doc).doc?.eq(applied as typeof doc)).toBe(true)
	})
})
