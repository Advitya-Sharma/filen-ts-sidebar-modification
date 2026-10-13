import { TextSelection } from "prosemirror-state"
import { markInternal } from "../../../core/docx/edit/invariants/index"
import { bodySections, docSectionChanges, finalSectionEdit, newSectionKey, type SectionEdit } from "../../../core/docx/stories/sections"
import { docxSchema, type DocxEditorCommand, type DocxEditorSession, type PagedEditor } from "../harness"
import { textPositions, textblocks } from "./base"
import type { EditOp, OpContext, Rand } from "./index"

// Headers, footers and sections under the paged editor (whose pages must stay those of the saved file):
// section breaks inserted and deleted, page setup, headers and footers made, entered, typed into, linked,
// unlinked and removed, different first and even pages, tracked page setup accepted or rejected; section edits
// made raw for the writer and Gate 2; and page setup of the final section in the scripted suite.

const marks = docxSchema.marks
const AUTHOR = "Fuzz Author"
const WORDS = ["Header", "Filen", "Page", "Kraków", "x"]

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

// A caret somewhere in the body's text outside tables (the body entered first).
function bodyCaret(editor: PagedEditor, rand: Rand): number | undefined {
	if (editor.session !== editor.body.session) editor.exitStory()

	const session = editor.session
	const positions = textPositions(session.state.doc, false)

	if (positions.length === 0) return undefined

	const at = pick(rand, positions)

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at)))

	return at
}

function execute(context: OpContext, command: DocxEditorCommand): string {
	return context.execute(command.type, command)
}

function headerOp(name: string, weight: number, run: (context: OpContext, editor: PagedEditor) => string): EditOp {
	return {
		name,
		harness: "paged",
		weight,
		run(context) {
			const editor = context.editor

			return editor === undefined ? "noop" : run(context, editor)
		}
	}
}

function tracked(session: DocxEditorSession, run: () => string): string {
	const review = session.context.review
	const identity = session.context.identity
	const was = review.tracking
	const author = identity.author

	review.tracking = true
	identity.author ??= AUTHOR

	try {
		return run()
	} finally {
		review.tracking = was
		identity.author = author
	}
}

function pageSetup(rand: Rand): DocxEditorCommand {
	const options: DocxEditorCommand[] = [
		{ type: "page.setup", orientation: rand() < 0.5 ? "landscape" : "portrait" },
		{
			type: "page.setup",
			size: pick(rand, [
				{ width: 11906, height: 16838 },
				{ width: 12240, height: 15840 },
				{ width: 10080, height: 14400 }
			])
		},
		{
			type: "page.setup",
			margins: { top: 720 + Math.floor(rand() * 1440), left: 720 + Math.floor(rand() * 1440), header: 360 + Math.floor(rand() * 720) }
		},
		{
			type: "page.setup",
			columns: { count: 1 + Math.floor(rand() * 3), space: 360 + Math.floor(rand() * 720), separator: rand() < 0.3 }
		},
		{
			type: "page.setup",
			pageNumbers: {
				start: rand() < 0.5 ? null : 1 + Math.floor(rand() * 20),
				format: pick(rand, ["decimal", "lowerRoman", "upperLetter"] as const)
			}
		},
		{
			type: "page.setup",
			lineNumbers: rand() < 0.3 ? null : { countBy: 1 + Math.floor(rand() * 5), start: 1 + Math.floor(rand() * 3) }
		},
		// Margins that may leave no room for text: the command refuses them as a save would.
		{ type: "page.setup", margins: { left: 2000 + Math.floor(rand() * 5000), right: 2000 + Math.floor(rand() * 5000) } }
	]
	const command = pick(rand, options)

	return rand() < 0.2 && command.type === "page.setup" ? { ...command, scope: "document" } : command
}

// The body or a header or footer, entered on a random page (as a double-click does).
function enterHeader(context: OpContext, editor: PagedEditor): string {
	const rand = context.rand
	const at = bodyCaret(editor, rand)

	if (at === undefined) return "noop"

	const type = rand() < 0.6 ? "header.edit" : "footer.edit"
	const done = execute(context, { type })

	if (done === "true" && rand() < 0.8) editor.session.dispatch(editor.session.state.tr.insertText(`${pick(rand, WORDS)} `))

	return `${type}@${at} -> ${done}`
}

// In a header or footer (entered first): a toggle of the chip.
function headerToggle(context: OpContext, editor: PagedEditor): string {
	const rand = context.rand
	const kind = editor.session.story.story.kind
	const entered = kind === "header" || kind === "footer" ? "in" : enterHeader(context, editor)
	const type = pick(rand, [
		"header.differentFirst",
		"header.differentOddEven",
		"header.linkToPrevious",
		"header.remove",
		"footer.remove"
	] as const)

	return `${entered}; ${type} -> ${execute(context, { type })}`
}

// The ids of the tracked section changes (as the review index counts them at the breaks).
function resolveSections(context: OpContext, editor: PagedEditor): string {
	if (editor.session !== editor.body.session) editor.exitStory()

	const type = context.rand() < 0.5 ? "review.acceptAll" : "review.rejectAll"

	return `${type} -> ${execute(context, { type })}`
}

// ---- damage: section edits made raw ----

function sectionDamage(context: OpContext): string {
	const rand = context.rand
	const session = context.session
	const doc = session.state.doc
	const part = session.story.part

	if (session.story.story.kind !== "body") return "noop"

	const sections = bodySections(doc, part)
	const section = pick(rand, sections)
	const edit: SectionEdit = {
		section: section.key,
		base: section.edit?.base ?? section.key,
		...pick(rand, [
			{ pageSize: { width: 15840, height: 12240, landscape: true } },
			{ columns: { count: 3, space: 360, separator: true } },
			{ titlePage: rand() < 0.5 },
			{ type: pick(rand, ["continuous", "evenPage", null]) },
			{ headers: { default: Math.floor(rand() * 4) - 1 } },
			{ pageMargins: { top: 99999, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 } },
			{ change: { author: "", date: "", before: null } }
		] as Partial<SectionEdit>[])
	}

	return `sectionEdit ${section.key} ${JSON.stringify(edit)} -> ${context.transact("sectionEdit", tr => {
		if (section.pos < 0) tr.setDocAttribute("sectionChanges", { ...docSectionChanges(doc), final: edit })
		else tr.addNodeMark(section.pos, marks.section_change.create({ value: edit }))

		return markInternal(tr)
	})}`
}

// A paragraph made a new section break, with or without an edit naming it.
function breakDamage(context: OpContext): string {
	const rand = context.rand
	const session = context.session
	const blocks = textblocks(session.state.doc).filter(block => (block.node.attrs["sectPr"] as number) < 0)

	if (blocks.length === 0 || session.story.story.kind !== "body") return "noop"

	const block = pick(rand, blocks)
	const key = newSectionKey(session.context)
	const named = rand() < 0.7

	return `newBreak ${block.pos} ${String(named)} -> ${context.transact("newBreak", tr => {
		tr.setNodeAttribute(block.pos, "sectPr", key)

		if (named)
			tr.addNodeMark(
				block.pos,
				marks.section_change.create({ value: { section: key, base: session.state.doc.attrs["sectPr"] as number } })
			)

		return markInternal(tr)
	})}`
}

export const headerOps: readonly EditOp[] = [
	headerOp("sectionBreak", 2, (context, editor) => {
		const at = bodyCaret(editor, context.rand)

		if (at === undefined) return "noop"

		const kind = pick(context.rand, ["nextPage", "continuous", "evenPage", "oddPage"] as const)
		const run = (): string => execute(context, { type: "insert.sectionBreak", kind })

		return `insert.sectionBreak ${kind}@${at} -> ${context.rand() < 0.3 ? `tracked ${tracked(editor.body.session, run)}` : run()}`
	}),
	headerOp("sectionBreakDelete", 1, (context, editor) => {
		const at = bodyCaret(editor, context.rand)

		if (at === undefined) return "noop"

		const run = (): string => execute(context, { type: "section.deleteBreak" })

		return `section.deleteBreak@${at} -> ${context.rand() < 0.3 ? `tracked ${tracked(editor.body.session, run)}` : run()}`
	}),
	headerOp("pageSetup", 2, (context, editor) => {
		const at = bodyCaret(editor, context.rand)

		if (at === undefined) return "noop"

		const command = pageSetup(context.rand)
		const run = (): string => execute(context, command)

		return `${JSON.stringify(command)}@${at} -> ${context.rand() < 0.25 ? tracked(editor.body.session, run) : run()}`
	}),
	headerOp("headerEnter", 2, enterHeader),
	headerOp("headerToggle", 2, headerToggle),
	headerOp("sectionResolve", 1, resolveSections),
	{ name: "sectionDamage", harness: "damage", run: sectionDamage },
	{ name: "sectionBreakDamage", harness: "damage", run: breakDamage },
	{
		name: "pageSetupScripted",
		harness: "damage",
		weight: 0,
		run: () => "noop",
		script: {
			order: 1010,
			// The final section turns landscape with narrower margins (a new w:sectPr where the body has none).
			run(target) {
				const session = target.session
				const doc = session.state.doc

				if (finalSectionEdit(doc) !== undefined) return false

				const key = doc.attrs["sectPr"] as number
				const edit: SectionEdit = {
					section: key,
					base: key,
					pageSize: { width: 15840, height: 12240, landscape: true },
					pageMargins: { top: 1080, right: 1080, bottom: 1080, left: 1080, header: 708, footer: 708, gutter: 0 }
				}

				return session.dispatch(
					markInternal(session.state.tr.setDocAttribute("sectionChanges", { ...docSectionChanges(doc), final: edit }))
				)
			}
		}
	}
]
