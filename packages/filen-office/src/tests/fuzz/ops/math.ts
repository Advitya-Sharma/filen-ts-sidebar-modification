import type { Node } from "prosemirror-model"
import { NodeSelection, TextSelection } from "prosemirror-state"
import { ReviewKind, latexToOmml, selectEquation, storyReviewIndex, trackedInsertion } from "@filen/office/core/docx/edit"
import { docxSchema, type DocxEditorSession, type PagedEditor } from "../harness"
import { textPositions, textblocks } from "./base"
import type { EditOp, OpContext, Rand } from "./index"

// Equations: inserting them from LaTeX (inline and display), editing and deleting them, with changes tracked or
// not, and accepting or rejecting the changes inside them, under the paged editor, whose pages must stay those
// of the saved file; new math nodes (tracked or not) made raw for the writer and Gate 2; and an inline equation
// at the end of the body in the scripted suite (the schema oracle validates its OMML).

const nodes = docxSchema.nodes
const AUTHORS = ["Fuzz Author", "Second Author"]
const LATEX = [
	"x^2",
	"\\frac{a}{b}",
	"\\sqrt[3]{x+1}",
	"\\sum_{i=1}^{n} i^2",
	"\\int\\limits_0^1 f(x)\\,dx",
	"\\left( a \\middle| b \\right)",
	"\\begin{pmatrix} 1 & 0 \\\\ 0 & 1 \\end{pmatrix}",
	"\\begin{cases} x & x > 0 \\\\ -x & \\text{otherwise} \\end{cases}",
	"\\lim_{n\\to\\infty} \\frac{1}{n}",
	"\\hat{x} + \\overline{y} + \\mathbb{R}",
	"\\underbrace{a+b}_{2}",
	"\\alpha\\beta\\gamma",
	"e^{i\\pi} + 1 = 0"
]

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

function equations(doc: Node): number[] {
	const out: number[] = []

	doc.descendants((node, pos) => {
		if (node.type === nodes.math) out.push(pos)

		return !node.isInline
	})

	return out
}

// Runs `run` with changes tracked (by one of two authors) as often as not, as before after it.
function maybeTracked(session: DocxEditorSession, rand: Rand, run: () => string): string {
	if (rand() < 0.5) return run()

	const review = session.context.review
	const identity = session.context.identity
	const was = review.tracking
	const author = identity.author

	review.tracking = true
	identity.author = pick(rand, AUTHORS)

	try {
		return `tracked ${run()}`
	} finally {
		review.tracking = was
		identity.author = author
	}
}

function paged(name: string, weight: number, run: (context: OpContext, editor: PagedEditor) => string): EditOp {
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

function insert(context: OpContext, editor: PagedEditor): string {
	const rand = context.rand
	const session = editor.session
	const positions = textPositions(session.state.doc)

	if (positions.length === 0) return "noop"

	const at = pick(rand, positions)
	const latex = pick(rand, LATEX)
	const display = rand() < 0.4

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at)))

	return maybeTracked(
		session,
		rand,
		() =>
			`equation.insert${display ? " display" : ""}@${at} -> ${context.execute("equation.insert", { type: "equation.insert", latex, display })}`
	)
}

function edit(context: OpContext, editor: PagedEditor): string {
	const rand = context.rand
	const session = editor.session
	const found = equations(session.state.doc)

	if (found.length === 0) return "noop"

	const at = pick(rand, found)

	if (!selectEquation(session, at)) return "noop"

	return maybeTracked(
		session,
		rand,
		() => `equation.edit@${at} -> ${context.execute("equation.edit", { type: "equation.edit", latex: pick(rand, LATEX) })}`
	)
}

// An equation deleted (selected and deleted, as a key would).
function remove(context: OpContext, editor: PagedEditor): string {
	const rand = context.rand
	const session = editor.session
	const found = equations(session.state.doc)

	if (found.length === 0) return "noop"

	const at = pick(rand, found)

	return maybeTracked(session, rand, () => {
		session.dispatch(session.state.tr.setSelection(NodeSelection.create(session.state.doc, at)))

		return `equation.delete@${at} -> ${String(session.dispatch(session.state.tr.deleteSelection()))}`
	})
}

// A change inside an equation accepted or rejected, or all of them.
function review(context: OpContext, editor: PagedEditor): string {
	const rand = context.rand
	const session = editor.session
	const doc = session.state.doc
	const inEquations = storyReviewIndex(session)
		.entries()
		.filter(entry => entry.kind > ReviewKind.COMMENT_REFERENCE && doc.nodeAt(entry.from)?.type === nodes.math)

	if (inEquations.length === 0) return "noop"

	const accept = rand() < 0.5
	const entry = pick(rand, inEquations)
	const command =
		rand() < 0.3
			? ({ type: accept ? "review.acceptAll" : "review.rejectAll" } as const)
			: ({ type: accept ? "review.accept" : "review.reject", id: entry.id } as const)

	return `${command.type}${"id" in command ? ` ${command.id}` : ""} -> ${context.execute(command.type, command)}`
}

// A new inline equation, at the end of the last paragraph (`rand` undefined) or anywhere in a random one.
function rawEquation(session: DocxEditorSession, rand: Rand | undefined): string {
	const blocks = textblocks(session.state.doc)

	if (blocks.length === 0) return "noop"

	const block = rand === undefined ? (blocks[blocks.length - 1] as { node: Node; pos: number }) : pick(rand, blocks)
	const at = block.pos + 1 + (rand === undefined ? block.node.content.size : Math.floor(rand() * (block.node.content.size + 1)))
	const written = latexToOmml(rand === undefined ? "\\frac{a}{b} + \\sqrt{x_1^2}" : pick(rand, LATEX), false)

	if (!written.ok) return "noop"

	// Tracked as inserted by an author (half of them; the scripted one too).
	const omml =
		rand === undefined || rand() < 0.5
			? trackedInsertion(written.omml, session.context.ids.next("revision"), {
					author: AUTHORS[0] ?? "",
					date: "2026-01-01T00:00:00Z"
				})
			: written.omml

	try {
		return `mathRaw@${String(at)} -> ${String(session.dispatch(session.state.tr.insert(at, nodes.math.create({ src: -1, omml }))))}`
	} catch {
		return "noop"
	}
}

export const mathOps: readonly EditOp[] = [
	paged("equationInsert", 2, insert),
	paged("equationEdit", 1, edit),
	paged("equationDelete", 1, remove),
	paged("equationReview", 1, review),
	{ name: "mathRaw", harness: "damage", weight: 1, run: context => rawEquation(context.session, context.rand) },
	{
		name: "equationInsertScripted",
		harness: "damage",
		weight: 0,
		run: context => rawEquation(context.session, context.rand),
		script: {
			order: 730,
			run(target) {
				return rawEquation(target.session, undefined).endsWith("true")
			}
		}
	}
]
