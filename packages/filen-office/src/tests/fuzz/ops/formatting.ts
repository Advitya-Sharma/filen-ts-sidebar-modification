import type { Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { DOCX_HIGHLIGHT_COLORS, DOCX_THEME_COLORS } from "@filen/office/core/docx/edit"
import { docxSchema, type DocxEditorCommand, type DocxEditorSession } from "../harness"
import { inlinePosition, nearPosition, textPositions } from "./base"
import type { EditOp, OpContext, Rand, ScriptEdit } from "./index"

// Formatting and insert commands: fonts, sizes, colours, character and paragraph properties, Clear All
// Formatting, the format painter, links, breaks and symbols.

const nodes = docxSchema.nodes

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

function int(rand: Rand, min: number, max: number): number {
	return min + Math.floor(rand() * (max - min + 1))
}

const FONTS = ["Arial", "Georgia", "Courier New", "Verdana", "+body", "+headings", "Times New Roman"]
const SIZES = [2, 8, 16, 18, 21, 22, 24, 28, 36, 48, 72, 96, 144, 400]
const SYMBOLS = ["©", "—", "‑", "­", " ", "…", "\u{1f600}", "é́", "\t", "€"]
const TARGETS = ["https://filen.io/", "https://example.com/a?b=c&d=<e>", "mailto:x@example.com"]

function runCommand(rand: Rand): DocxEditorCommand {
	switch (int(rand, 0, 12)) {
		case 0:
			return { type: "font.family", family: pick(rand, FONTS) }
		case 1:
			return { type: "font.size", halfPoints: pick(rand, SIZES) }
		case 2:
			return { type: rand() < 0.5 ? "font.grow" : "font.shrink" }
		case 3: {
			const kind = int(rand, 0, 2)

			return {
				type: "color.text",
				color:
					kind === 0
						? { kind: "auto" }
						: kind === 1
							? { kind: "rgb", value: pick(rand, ["FF0000", "00B050", "000000", "0563C1"]) }
							: {
									kind: "theme",
									theme: pick(rand, DOCX_THEME_COLORS),
									...(rand() < 0.5 ? { tint: int(rand, 0, 255) } : rand() < 0.5 ? { shade: int(rand, 0, 255) } : {})
								}
			}
		}
		case 4:
			return { type: "color.highlight", name: pick(rand, DOCX_HIGHLIGHT_COLORS) }
		case 5:
			return { type: "font.spacing", twips: pick(rand, [-40, -10, 0, 15, 100]) }
		case 6:
			return { type: "format.superscript" }
		case 7:
			return { type: "format.subscript" }
		case 8:
			return { type: "format.doubleStrike" }
		case 9:
			return { type: "format.smallCaps" }
		case 10:
			return { type: "format.allCaps" }
		case 11:
			return { type: "format.clearAll" }
		default:
			return { type: "format.strike" }
	}
}

function paragraphCommand(rand: Rand): DocxEditorCommand {
	switch (int(rand, 0, 3)) {
		case 0: {
			const first = rand() < 0.5

			return {
				type: "paragraph.indent",
				...(rand() < 0.7 ? { start: pick(rand, [0, 360, 720, 1440, -360]) } : {}),
				...(rand() < 0.3 ? { end: pick(rand, [0, 720]) } : {}),
				...(first ? { firstLine: pick(rand, [0, 240, 720]) } : { hanging: pick(rand, [0, 360]) })
			}
		}
		case 1:
			return {
				type: "paragraph.spacing",
				before: pick(rand, [0, 120, 240]),
				...(rand() < 0.5 ? { after: pick(rand, [0, 160, 480]) } : {})
			}
		case 2: {
			const rule = pick(rand, ["auto", "exact", "atLeast"] as const)

			return {
				type: "paragraph.lineSpacing",
				rule,
				value: rule === "auto" ? pick(rand, [240, 259, 276, 360, 480]) : pick(rand, [200, 240, 400])
			}
		}
		default:
			return { type: "format.clearAll" }
	}
}

function insertCommand(rand: Rand, story: string): DocxEditorCommand {
	const kind = int(rand, 0, 5)

	if (kind === 5) return { type: "insert.horizontalLine" }
	if (kind === 0 && story === "body") return { type: "insert.pageBreak" }
	if (kind === 1) return { type: "insert.lineBreak" }
	if (kind === 2) return { type: "link.remove" }

	if (kind === 3) {
		const text = rand() < 0.5 ? { text: pick(rand, ["Filen", "a link", "été \u{1f600}"]) } : {}

		return rand() < 0.8 ? { type: "link.set", target: pick(rand, TARGETS), ...text } : { type: "link.set", anchor: "_Toc1", ...text }
	}

	return { type: "insert.symbol", text: pick(rand, SYMBOLS) }
}

function select(session: DocxEditorSession, a: number, b: number = a): void {
	const doc = session.state.doc

	session.dispatch(session.state.tr.setSelection(TextSelection.between(doc.resolve(Math.min(a, b)), doc.resolve(Math.max(a, b)))))
}

// A caret (a third of the time) or a range near a random text position.
function selectSome(context: OpContext): string | undefined {
	const session = context.session
	const positions = textPositions(session.state.doc)

	if (positions.length === 0) return undefined

	const a = pick(context.rand, positions)
	const b = context.rand() < 0.33 ? a : (nearPosition(context.rand, session.state.doc, a) ?? a)

	select(session, a, b)

	return `${Math.min(a, b)}-${Math.max(a, b)}`
}

function paged(name: string, run: (context: OpContext, at: string) => string, script?: ScriptEdit): EditOp {
	return {
		name,
		harness: "paged",
		run(context) {
			const at = selectSome(context)

			return at === undefined ? "noop" : run(context, at)
		},
		...(script === undefined ? {} : { script })
	}
}

function damage(name: string, command: (rand: Rand, story: string) => DocxEditorCommand): EditOp {
	return {
		name,
		harness: "damage",
		run(context) {
			const doc = context.session.state.doc
			const a = inlinePosition(context.rand, doc)

			if (a === undefined) return "noop"

			const b = context.rand() < 0.3 ? a : (nearPosition(context.rand, doc, a) ?? a)
			const label = `select ${a}-${b}`

			context.transact(label, tr =>
				tr.setSelection(TextSelection.between(tr.doc.resolve(Math.min(a, b)), tr.doc.resolve(Math.max(a, b))))
			)

			const next = command(context.rand, context.session.story.story.kind)

			return `${next.type}@${label} ${context.execute(JSON.stringify(next), next)}`
		}
	}
}

// ---- scripted suite --------------------------------------------------------------------------------------

// Body paragraphs (outside tables) that start with at least 8 characters of text.
function scriptParagraph(session: DocxEditorSession, index: number): { node: Node; pos: number } | undefined {
	const list: { node: Node; pos: number }[] = []

	session.state.doc.descendants((node, pos, parent) => {
		if (node.type === nodes.paragraph) {
			if (parent?.type !== nodes.table_cell && node.textContent.length >= 8 && node.firstChild?.isText === true)
				list.push({ node, pos })

			return false
		}

		return !node.isInline
	})

	return list[Math.min(list.length - 1, index)]
}

function scripted(session: DocxEditorSession, index: number, from: number, to: number, commands: readonly DocxEditorCommand[]): boolean {
	const target = scriptParagraph(session, index)

	if (target === undefined) return false

	select(session, target.pos + 1 + from, target.pos + 1 + to)

	return commands.map(command => session.execute(command)).some(done => done)
}

export const formattingOps: readonly EditOp[] = [
	paged(
		"formatRun",
		(context, at) => {
			const command = runCommand(context.rand)

			return `${JSON.stringify(command)}@${at} ${context.execute(command.type, command)}`
		},
		{
			order: 400,
			run: ({ session }) =>
				scripted(session, 9, 0, 5, [
					{ type: "font.family", family: "Georgia" },
					{ type: "font.size", halfPoints: 30 },
					{ type: "font.spacing", twips: 10 },
					{ type: "color.text", color: { kind: "theme", theme: "accent1", tint: 153 } },
					{ type: "color.highlight", name: "yellow" },
					{ type: "format.superscript" }
				])
		}
	),
	paged(
		"formatParagraph",
		(context, at) => {
			const command = paragraphCommand(context.rand)

			return `${JSON.stringify(command)}@${at} ${context.execute(command.type, command)}`
		},
		{
			order: 410,
			run: ({ session }) =>
				scripted(session, 10, 0, 0, [
					{ type: "paragraph.indent", start: 720, hanging: 360 },
					{ type: "paragraph.spacing", before: 240, after: 120 },
					{ type: "paragraph.lineSpacing", rule: "auto", value: 360 }
				])
		}
	),
	// Formatting at a caret, then typing with it.
	{
		name: "formatTyping",
		harness: "paged",
		run(context) {
			const session = context.session
			const positions = textPositions(session.state.doc)

			if (positions.length === 0) return "noop"

			const at = pick(context.rand, positions)

			select(session, at)

			const command = runCommand(context.rand)

			context.execute(command.type, command)
			session.dispatch(session.state.tr.insertText("Filen "))

			return `${command.type}+type@${at}`
		},
		script: {
			order: 420,
			run: ({ session }) =>
				scripted(session, 12, 0, 4, [{ type: "format.doubleStrike" }, { type: "format.smallCaps" }, { type: "font.grow" }])
		}
	},
	// The painter is the editor's (no scripted edit: the suite edits sessions); its script clears formatting.
	paged(
		"painter",
		(context, at) => {
			const sticky = context.rand() < 0.3

			context.execute("format.copy", { type: "format.copy", ...(sticky ? { sticky } : {}) })

			const to = selectSome(context)
			const done = context.execute("format.paste", { type: "format.paste" })

			if (sticky && context.editor !== undefined) context.editor.setPainter(undefined)

			return `painter ${at} -> ${to ?? "-"} ${done}`
		},
		{ order: 440, run: ({ session }) => scripted(session, 14, 0, 8, [{ type: "format.clearAll" }]) }
	),
	paged(
		"insert",
		(context, at) => {
			const command = insertCommand(context.rand, context.session.story.story.kind)

			return `${JSON.stringify(command)}@${at} ${context.execute(command.type, command)}`
		},
		{
			order: 430,
			run: ({ session }) =>
				scripted(session, 13, 3, 3, [
					{ type: "insert.symbol", text: "\u00a9\u2011" },
					{ type: "insert.lineBreak" },
					{ type: "link.set", target: "https://filen.io/", text: "Filen" }
				])
		}
	),
	damage("formatRunRaw", runCommand),
	damage("formatParagraphRaw", paragraphCommand),
	damage("insertRaw", insertCommand)
]
