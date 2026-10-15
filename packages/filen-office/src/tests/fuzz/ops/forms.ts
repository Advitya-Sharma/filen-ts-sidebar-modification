import type { Mark, Node } from "prosemirror-model"
import { NodeSelection, TextSelection } from "prosemirror-state"
import {
	controlOf,
	controlState,
	legacyFieldOf,
	newControlId,
	typeable,
	type ControlInfo,
	type ControlState
} from "../../../core/docx/forms/controls"
import { markCommand } from "../../../core/docx/edit/invariants/index"
import { DOCX_CONTROL_KINDS } from "../../../core/docx/edit/commands/forms"
import { docxSchema, type DocxEditorCommand, type DocxEditorSession } from "../harness"
import { BOUND_WRITES, inlinePosition } from "./base"
import type { EditOp, OpContext, Rand, ScriptEdit } from "./index"

// Content controls, legacy form fields and repeating sections through the form commands under the paged editor
// (a checkbox ticked, a list entry or date chosen, a control cleared, typed into or emptied, new controls and
// repeating section items), raw control states for the writer and Gate 2, and a checkbox and a placeholder in
// the scripted suite.

const nodes = docxSchema.nodes
const marks = docxSchema.marks

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

interface Found {
	readonly from: number
	readonly to: number
	readonly holder: Node | Mark
	readonly info: ControlInfo
	readonly block: boolean
}

// The story's controls: block ones by their content, inline ones by their span in a paragraph.
function controls(session: DocxEditorSession): Found[] {
	const part = session.story.part
	const out: Found[] = []

	session.state.doc.descendants((node, pos) => {
		if (node.type === nodes.sdt) {
			const info = controlOf(part, node)

			if (info !== undefined) out.push({ from: pos + 1, to: pos + node.nodeSize - 1, holder: node, info, block: true })
		}

		if (!node.inlineContent) return !node.isInline && node.type !== nodes.locked_block

		let at = pos + 1
		const open = new Map<Mark, { from: number; to: number }>()

		node.forEach(child => {
			for (const mark of child.marks) {
				if (mark.type !== marks.content_control) continue

				const span = open.get(mark)

				if (span !== undefined && span.to === at) span.to = at + child.nodeSize
				else open.set(mark, { from: at, to: at + child.nodeSize })
			}

			at += child.nodeSize
		})

		for (const [mark, span] of open) {
			const info = controlOf(part, mark)

			if (info !== undefined) out.push({ ...span, holder: mark, info, block: false })
		}

		return false
	})

	return out
}

function fields(session: DocxEditorSession): { pos: number; node: Node }[] {
	const out: { pos: number; node: Node }[] = []

	session.state.doc.descendants((node, pos) => {
		if (node.type === nodes.field && legacyFieldOf(session.story.part, node) !== undefined) out.push({ pos, node })

		return node.type !== nodes.locked_block
	})

	return out
}

function inside(session: DocxEditorSession, found: Found, rand: Rand, whole = false): boolean {
	const doc = session.state.doc
	const from = found.block ? found.from + 1 : found.from
	const to = found.block ? Math.max(from, found.to - 1) : found.to
	const a = whole ? from : from + Math.floor(rand() * (to - from + 1))

	try {
		return session.dispatch(session.state.tr.setSelection(TextSelection.between(doc.resolve(a), doc.resolve(whole ? to : a))))
	} catch {
		return false
	}
}

// Edits recorded as tracked changes need an author (the app asks the user; the fuzz names one).
function withAuthor(context: OpContext): DocxEditorSession {
	context.session.context.identity.author ??= "Fuzz Author"

	return context.session
}

function onControl(
	name: string,
	weight: number,
	can: (info: ControlInfo) => boolean,
	command: (info: ControlInfo, rand: Rand) => DocxEditorCommand
): EditOp {
	return {
		name,
		harness: "paged",
		weight,
		run(context) {
			const candidates = controls(withAuthor(context)).filter(found => can(found.info))

			if (candidates.length === 0) return "noop"

			const found = pick(context.rand, candidates)

			if (!inside(context.session, found, context.rand)) return "noop"

			const executed = command(found.info, context.rand)

			return `${name} ${found.from} -> ${context.execute(name, executed)}`
		}
	}
}

const ISO = (rand: Rand): string => {
	const date = new Date(Date.UTC(1990 + Math.floor(rand() * 60), Math.floor(rand() * 12), 1 + Math.floor(rand() * 28)))

	return date.toISOString().slice(0, 10)
}

const TYPED = ["x", "Ada Lovelace", "Kraków — naïve", " ", "1\t2"]

// Types into a control (over its placeholder, or into what it holds), or deletes all it holds.
function typeOrEmpty(context: OpContext): string {
	const session = withAuthor(context)
	const candidates = controls(session).filter(found => typeable(found.info) === undefined && found.info.type !== "repeatingSection")

	if (candidates.length === 0) return "noop"

	const found = pick(context.rand, candidates)
	const empty = context.rand() < 0.35

	if (!inside(session, found, context.rand, empty)) return "noop"

	const text = pick(context.rand, TYPED)

	return `formType ${found.from} ${empty ? "empty" : JSON.stringify(text)} -> ${context.transact("formType", tr =>
		empty ? tr.deleteSelection() : tr.insertText(text)
	)}`
}

function insertControl(context: OpContext): string {
	const session = withAuthor(context)
	const at = inlinePosition(context.rand, session.state.doc)

	if (at === undefined) return "noop"

	try {
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at)))
	} catch {
		return "noop"
	}

	const kind = pick(context.rand, DOCX_CONTROL_KINDS)
	const command: DocxEditorCommand =
		kind === "dropDownList" || kind === "comboBox"
			? {
					type: "control.insert",
					kind,
					items: [
						{ text: "One", value: "1" },
						{ text: "Two", value: "2" }
					]
				}
			: { type: "control.insert", kind }

	return `formInsert ${kind}@${at} -> ${context.execute("formInsert", command)}`
}

function item(context: OpContext): string {
	const candidates = controls(withAuthor(context)).filter(found => found.info.type === "repeatingSectionItem")

	if (candidates.length === 0) return "noop"

	const found = pick(context.rand, candidates)

	if (!inside(context.session, found, context.rand)) return "noop"

	const type = context.rand() < 0.6 ? "control.item.add" : "control.item.remove"

	return `${type} ${found.from} -> ${context.execute(type, { type })}`
}

function legacy(context: OpContext): string {
	const session = withAuthor(context)
	const candidates = fields(session)

	if (candidates.length === 0) return "noop"

	const field = pick(context.rand, candidates)
	const info = legacyFieldOf(session.story.part, field.node)

	if (info === undefined) return "noop"

	session.dispatch(session.state.tr.setSelection(NodeSelection.create(session.state.doc, field.pos)))

	const command: DocxEditorCommand =
		info.type === "FORMCHECKBOX"
			? { type: "control.check" }
			: info.type === "FORMDROPDOWN"
				? { type: "control.choose", value: pick(context.rand, [...info.entries, "missing"]) }
				: { type: "control.text", text: pick(context.rand, ["", "filled in", "two\nlines"]) }

	return `formLegacy ${command.type}@${field.pos} -> ${context.execute("formLegacy", command)}`
}

// ---- damage: raw control states ----

const STATES: readonly ((rand: Rand) => ControlState)[] = [
	rand => ({ checked: rand() < 0.5 }),
	rand => ({ placeholder: rand() < 0.5 }),
	() => ({ value: "1" }),
	rand => ({ date: rand() < 0.3 ? null : ISO(rand) }),
	rand => ({ made: { type: pick(rand, ["richText", "checkbox", "date", "plainText"] as const), id: newControlId(rand) } })
]

function stateDamage(context: OpContext): string {
	const rand = context.rand
	const candidates = controls(context.session)

	if (candidates.length === 0) return "noop"

	const found = pick(rand, candidates)
	const state: ControlState = { ...(controlState(found.holder) ?? {}), ...pick(rand, STATES)(rand) }

	return `formState ${found.from} ${JSON.stringify(state)} -> ${context.transact("formState", tr => {
		if (found.block) tr.setNodeAttribute(found.from - 1, "state", state)
		else {
			const mark = found.holder as Mark

			tr.removeMark(found.from, found.to, mark)
			tr.addMark(found.from, found.to, marks.content_control.create({ ...mark.attrs, state }))
		}

		return markCommand(tr)
	})}`
}

function contentDamage(context: OpContext): string {
	const rand = context.rand
	const candidates = controls(context.session)

	if (candidates.length === 0) return "noop"

	const found = pick(rand, candidates)

	return `formContent ${found.from} -> ${context.transact("formContent", tr => {
		if (rand() < 0.5) tr.delete(found.from, found.to)
		else tr.insertText(pick(rand, TYPED), found.from + (found.block ? 1 : 0))

		return rand() < 0.5 ? markCommand(tr) : tr
	})}`
}

// ---- scripted suite ----

function scripted(order: number, run: (session: DocxEditorSession) => boolean): ScriptEdit {
	return { order, run: ({ session }) => run(session) }
}

function firstControl(session: DocxEditorSession, can: (info: ControlInfo) => boolean): Found | undefined {
	return controls(session).find(found => can(found.info))
}

export const formOps: readonly EditOp[] = [
	onControl(
		"formCheck",
		3,
		info => info.checkbox !== null,
		() => ({ type: "control.check" })
	),
	onControl(
		"formChoose",
		2,
		info => (info.type === "dropDownList" || info.type === "comboBox") && info.items.length > 0,
		(info, rand) => ({ type: "control.choose", value: pick(rand, info.items).value })
	),
	// A list's entries changed: one dropped, one added (values unique, as Word keeps them).
	onControl(
		"formItems",
		1,
		info => info.type === "dropDownList" || info.type === "comboBox",
		(info, rand) => ({
			type: "control.items",
			items: [
				...info.items.filter((_, index) => index !== Math.floor(rand() * info.items.length) && info.items.length > 0),
				{ text: `Fuzz ${String(Math.floor(rand() * 1000))}`, value: `fuzz-${String(Math.floor(rand() * 1000))}` }
			].filter(
				(item, index, all) =>
					all.findIndex(other => other.value === item.value || other.text === item.text) === index && item.text !== ""
			)
		})
	),
	onControl(
		"formDate",
		2,
		info => info.date !== null,
		(_info, rand) => ({ type: "control.date", iso: ISO(rand) })
	),
	onControl(
		"formClear",
		1,
		() => true,
		() => ({ type: "control.clear" })
	),
	{ name: "formType", harness: "paged", weight: 3, run: typeOrEmpty },
	{ name: "formInsert", harness: "paged", weight: 1, run: insertControl },
	{ name: "formItem", harness: "paged", weight: 1, run: item },
	{ name: "formLegacy", harness: "paged", weight: 1, run: legacy },
	{
		name: "formCheckScript",
		harness: "paged",
		weight: 0,
		run: () => "noop",
		writes: BOUND_WRITES,
		script: scripted(750, session => {
			const found = firstControl(
				session,
				info => info.checkbox !== null && typeable(info) !== "contentLocked" && typeable(info) !== "bound"
			)

			return found !== undefined && inside(session, found, () => 0) && session.execute({ type: "control.check" })
		})
	},
	{
		name: "formTypeScript",
		harness: "paged",
		weight: 0,
		run: () => "noop",
		writes: BOUND_WRITES,
		script: scripted(760, session => {
			const found = firstControl(session, info => info.showingPlaceholder && typeable(info) === undefined && !info.binding)

			if (found === undefined || !inside(session, found, () => 0.5)) return false

			return session.dispatch(session.state.tr.insertText("Filled"))
		})
	},
	{ name: "formState", harness: "damage", weight: 2, run: stateDamage },
	{ name: "formContent", harness: "damage", weight: 1, run: contentDamage }
]
