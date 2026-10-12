import type { Mark, Node } from "prosemirror-model"
import { closeHistory } from "prosemirror-history"
import type { Transaction } from "prosemirror-state"
import { AddMarkStep, RemoveMarkStep, ReplaceAroundStep, ReplaceStep, StepMap } from "prosemirror-transform"
import { blockText, boundValue, inlineText, sharedBindings } from "../../forms/binding"
import {
	blockControlAround,
	checkboxGlyph,
	controlKey,
	controlOf,
	controlState,
	inlineControlAround,
	isoDate,
	pictureSwap,
	placeholderText,
	simpleBinding,
	typeable,
	type ControlAround,
	type ControlInfo,
	type ControlState
} from "../../forms/controls"
import { formatControlDate } from "../../forms/dates"
import type { StoryPart } from "../../model/index"
import type { DocxEditContext } from "../context"
import { protectionOf, protectionViolation } from "../protection"
import { docxSchema } from "../schema"
import type { DocxEditorSession } from "../session"
import { invariantViolation, isCommand, isInternal, markCommand, markInternal } from "./core"
import type { InvariantStage } from "./index"

// Content controls keep what Word keeps (ECMA-376 Part 1 §17.5.2): locked content (w:lock contentLocked,
// sdtContentLocked) and controls whose content only form commands change (a checkbox, a drop-down list, a
// picture, a group outside the controls inside it) or that are bound in a way the editor does not update are
// not typed into; a control locked against deletion (sdtLocked, sdtContentLocked) stays. A control showing its
// placeholder (w:showingPlcHdr) loses the placeholder text and its style once something is typed into it, and
// gets them back once emptied, as in Word. Document protection limits edits to what its mode allows
// (edit/protection.ts).
//
// These run as transaction transformers (handlers/forms.ts registers them): one before tracked changes record
// a transaction (a protection that forces tracking turns it on), one after it (protection, locks,
// placeholders). What they refuse, this stage refuses.

const nodes = docxSchema.nodes
const marks = docxSchema.marks
const refusals = new WeakMap<Transaction, string>()
// Transactions form commands made (they set control state themselves).
const formCommands = new WeakSet<Transaction>()
// Sessions whose transaction being transformed is a form command's (between the two transformers).
const commandOf = new WeakSet<DocxEditorSession>()
// Sessions holding controls the editor made (a part without source controls is not looked at otherwise).
const madeIn = new WeakSet<DocxEditorSession>()
// What placeholder text keeps of the content it stands in for.
const DROPPED_MARKS = new Set<string>(["insertion", "deletion", "move_from", "move_to", "format_change", "link"])

export const FORMS_STAGE: InvariantStage | null = {
	name: "forms",
	violation: tr => refusals.get(tr),
	// Bound values follow once everything is applied (tracked changes record deletions in a repair of their own).
	repair: (transactions, state) => {
		const session = transactions.map(tr => sessionOf.get(tr)).find(found => found !== undefined)

		if (session === undefined || sharedBindings(session.story.part).size === 0) return null

		const points: number[] = []

		transactions.forEach((tr, index) => {
			let at = touchedPoints(tr)

			for (let later = index + 1; later < transactions.length; later++) {
				const mapping = (transactions[later] as Transaction).mapping

				at = at.map(pos => mapping.map(pos, -1))
			}

			points.push(...at)
		})

		const tr = state.tr

		return mirrorBindings(tr, session, points) ? markInternal(tr) : null
	}
}

// The session a transformed transaction belongs to (for the repair after it).
const sessionOf = new WeakMap<Transaction, DocxEditorSession>()

// Marks a form command's transaction: recorded like typing when tracking is on, its state changes allowed.
// `made`: it inserts a control.
export function formCommand(tr: Transaction, session: DocxEditorSession, made = false): Transaction {
	// Each command is an undo step of its own.
	formCommands.add(closeHistory(tr))

	if (made) madeIn.add(session)

	return tr
}

// Whether a session's story may hold content controls (source ones, or ones the editor made).
export function mayHoldControls(session: DocxEditorSession): boolean {
	return session.story.part.sdts.count > 0 || madeIn.has(session)
}

export function refusedForForms(tr: Transaction): string | undefined {
	return refusals.get(tr)
}

function refuse(tr: Transaction, reason: string, session: DocxEditorSession): Transaction {
	// The invariants never filter the session's own transactions (a bulk tracked change): nothing applies instead.
	if (isInternal(tr)) return session.state.tr

	refusals.set(tr, reason)

	return tr
}

// Before tracked changes record the transaction.
export function formsBefore(tr: Transaction, session: DocxEditorSession): Transaction {
	const context = session.context

	if (protectionOf(context).forcesTracking && !context.review.tracking) context.review.tracking = true

	if (formCommands.has(tr)) commandOf.add(session)
	else commandOf.delete(session)

	return tr
}

// After tracked changes recorded it: protection, locks, placeholders.
export function formsAfter(tr: Transaction, session: DocxEditorSession): Transaction {
	const out = checked(tr, session)

	if (out.docChanged) sessionOf.set(out, session)

	return out
}

function checked(tr: Transaction, session: DocxEditorSession): Transaction {
	const command = commandOf.has(session)

	commandOf.delete(session)

	if (!tr.docChanged) return tr

	const protection = protectionViolation(tr, session, command)

	if (protection !== undefined) return refuse(tr, protection, session)
	if (command) {
		if (isInternal(tr)) return tr

		unifyStates(tr)

		return markCommand(tr)
	}

	const part = session.story.part

	if (!mayHoldControls(session)) return tr

	const lock = lockViolation(tr, part)

	if (lock !== undefined) return refuse(tr, lock, session)
	if (isInternal(tr) || isCommand(tr) || invariantViolation(tr) !== undefined) return tr

	return placeholders(tr, session, protectionOf(session.context).mode === "forms") ? markCommand(tr) : tr
}

// ---- controls bound to one node -------------------------------------------------------------------------

// Positions of the final document a transaction touched.
function touchedPoints(tr: Transaction): number[] {
	const out: number[] = []

	tr.steps.forEach((step, index) => {
		const map = tr.mapping.maps[index] ?? StepMap.empty

		for (let k = 0; k < out.length; k++) out[k] = map.map(out[k] ?? 0, -1)

		if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) out.push(step.from, step.to)
		else map.forEach((_oldStart, _oldEnd, newStart, newEnd) => out.push(newStart, newEnd))
	})

	return out
}

// The controls around `pos` (inline ones touching it, block ones holding it) bound by a path the editor
// updates to a node other controls of the part are bound to too.
function boundAt(part: StoryPart, doc: Node, pos: number, shared: ReadonlySet<string>): { around: ControlAround; key: string }[] {
	const out: { around: ControlAround; key: string }[] = []
	const add = (around: ControlAround): void => {
		const binding = around.info.binding === null || around.info.richBinding ? undefined : simpleBinding(around.info.binding)

		if (binding !== undefined && shared.has(binding.key)) out.push({ around, key: binding.key })
	}
	const $pos = doc.resolve(pos)

	if ($pos.parent.inlineContent) {
		for (const span of controlSpans($pos.parent, $pos.start())) {
			if (span.from > pos || span.to < pos) continue

			const info = controlOf(part, span.mark)

			if (info !== undefined) add({ kind: "inline", holder: span.mark, info, from: span.from, to: span.to, pos: span.from })
		}
	}

	for (let depth = $pos.depth; depth > 0; depth--) {
		const node = $pos.node(depth)

		if (node.type !== nodes.sdt) continue

		const info = controlOf(part, node)
		const at = $pos.before(depth)

		if (info !== undefined) add({ kind: "block", holder: node, info, from: at + 1, to: at + node.nodeSize - 1, pos: at })
	}

	return out
}

// One piece of a control: a control an edit split over two paragraphs is written as two, each showing the
// bound value.
function pieceKey(around: ControlAround): string {
	return `${controlKey(around.holder)}@${String(around.from)}`
}

function controlText(doc: Node, around: ControlAround): string {
	if (around.kind === "block") return blockText(around.holder as Node)

	let text = ""

	doc.nodesBetween(around.from, around.to, node => {
		if (node.isInline) text += inlineText(node)
	})

	return text
}

// Word shows a node's value in every control bound to it: after an edit of one, the others of the story that
// share its binding (forms/binding.ts sharedBindings) show its value too. Whether it changed any.
function mirrorBindings(tr: Transaction, session: DocxEditorSession, points: readonly number[]): boolean {
	const part = session.story.part
	const shared = sharedBindings(part)

	if (shared.size === 0) return false

	const doc = tr.doc
	const values = new Map<string, { value: string; source: string }>()
	const head = tr.selection.head

	for (const point of points) {
		for (const found of boundAt(part, doc, Math.min(point, doc.content.size), shared)) {
			const info = found.around.info
			const value = info.showingPlaceholder ? "" : boundValue(info, controlText(doc, found.around))
			const holds = head >= found.around.from && head <= found.around.to

			if (!values.has(found.key) || holds) values.set(found.key, { value, source: pieceKey(found.around) })
		}
	}

	if (values.size === 0) return false

	const targets: { around: ControlAround; value: string }[] = []
	const consider = (around: ControlAround): void => {
		const binding = around.info.binding === null || around.info.richBinding ? undefined : simpleBinding(around.info.binding)
		const wanted = binding === undefined ? undefined : values.get(binding.key)

		if (wanted === undefined || wanted.source === pieceKey(around)) return

		const value = around.info.showingPlaceholder ? "" : boundValue(around.info, controlText(doc, around))

		if (value !== wanted.value) targets.push({ around, value: wanted.value })
	}

	doc.descendants((node, pos) => {
		if (node.type === nodes.sdt) {
			const info = controlOf(part, node)

			if (info !== undefined) consider({ kind: "block", holder: node, info, from: pos + 1, to: pos + node.nodeSize - 1, pos })
		}

		if (!node.inlineContent) return !node.isInline && node.type !== nodes.locked_block

		for (const span of controlSpans(node, pos + 1)) {
			const info = controlOf(part, span.mark)

			if (info !== undefined) consider({ kind: "inline", holder: span.mark, info, from: span.from, to: span.to, pos: span.from })
		}

		return false
	})

	for (const { around, value } of targets.sort((a, b) => b.around.from - a.around.from))
		showValue(tr, session.context, part, around, value)

	return targets.length > 0
}

// Shows a bound node's `value` in a control as its type shows it (an empty value: the placeholder).
function showValue(tr: Transaction, context: DocxEditContext, part: StoryPart, around: ControlAround, value: string): void {
	const info = around.info
	const state = controlState(around.holder) ?? {}

	if (value === "") {
		writeControl(tr, context, around, placeholderText(part, around.holder, info), placeholderState(around.holder, info, true))

		return
	}

	if (info.checkbox !== null) {
		const checked = value === "true" || value === "1"
		const glyph = checkboxGlyph(info.checkbox, checked)

		writeControl(tr, context, around, glyph.text, { ...state, checked, placeholder: false }, glyph.font)

		return
	}

	if (info.date !== null) {
		const iso = isoDate(value)

		if (iso !== null) {
			writeControl(tr, context, around, formatControlDate(iso, info.date.format, info.date.lid, info.date.calendar), {
				...state,
				date: iso,
				placeholder: false
			})

			return
		}
	}

	if (info.type === "dropDownList" || info.type === "comboBox") {
		const item = info.items.find(entry => entry.value === value)

		writeControl(tr, context, around, item?.text ?? value, { ...state, value: item?.value ?? null, placeholder: false })

		return
	}

	writeControl(tr, context, around, value, { ...state, placeholder: false })
}

// ---- locks ----------------------------------------------------------------------------------------------

export interface ControlSpan {
	readonly from: number
	readonly to: number
	readonly mark: Mark
}

// The run-level controls of a paragraph (content from `start`) as spans, one per control (nested ones overlap).
export function controlSpans(paragraph: Node, start: number): ControlSpan[] {
	const open = new Map<string, { from: number; to: number; mark: Mark }>()
	const out: ControlSpan[] = []
	let pos = start

	paragraph.forEach(child => {
		const end = pos + child.nodeSize
		const here = new Set<string>()

		for (const mark of child.marks) {
			if (mark.type !== marks.content_control) continue

			const key = controlKey(mark)
			const span = open.get(key)

			here.add(key)

			if (span !== undefined && span.to === pos) span.to = end
			else {
				if (span !== undefined) out.push(span)

				open.set(key, { from: pos, to: end, mark })
			}
		}

		for (const [key, span] of open) {
			if (!here.has(key)) {
				out.push(span)
				open.delete(key)
			}
		}

		pos = end
	})

	for (const span of open.values()) out.push(span)

	return out
}

const REASONS = { contentLocked: "lockedContentControl", bound: "boundContentControl", notText: "formControl" } as const

// Zero-width content (range markers, comment references) may go anywhere.
function onlyMarkers(step: ReplaceStep | ReplaceAroundStep): boolean {
	if (step.from !== step.to || step instanceof ReplaceAroundStep) return false

	let only = true

	step.slice.content.forEach(node => {
		if (node.type !== nodes.marker && node.type !== nodes.comment_ref) only = false
	})

	return only
}

// Whether [from, to] lies inside a control a user may type into, nested in `outer` (the content range of one
// that may not be typed into).
function nestedTypeable(part: StoryPart, doc: Node, from: number, to: number, outer: { from: number; to: number }): boolean {
	const inner = inlineControlAround(part, doc, from, to) ?? blockControlAround(part, doc, from, to)

	return (
		inner !== undefined &&
		typeable(inner.info) === undefined &&
		inner.from >= outer.from &&
		inner.to <= outer.to &&
		(inner.from !== outer.from || inner.to !== outer.to)
	)
}

// Why a step range breaks a control's lock: `replace` changes content (`from` = `to` inserts; `inserts`: content
// comes in), `format` sets marks, `unwrap` takes a control mark away.
function rangeViolation(
	part: StoryPart,
	doc: Node,
	from: number,
	to: number,
	kind: "replace" | "format" | "unwrap",
	inserts = false
): string | undefined {
	let violation: string | undefined
	const visitParagraph = (paragraph: Node, start: number): void => {
		for (const span of controlSpans(paragraph, start)) {
			const overlaps = from === to ? from > span.from && from < span.to : from < span.to && to > span.from

			if (!overlaps) continue

			const info = controlOf(part, span.mark)

			if (info === undefined) continue

			if (kind === "unwrap") {
				if (info.lock === "sdtLocked" || info.lock === "sdtContentLocked") violation ??= REASONS.contentLocked

				continue
			}

			const whole = from < to && from <= span.from && to >= span.to
			const blocker = typeable(info)

			if (kind === "replace" && whole) {
				const exact = from === span.from && to === span.to

				// Exactly its content typed over or deleted: the placeholder comes back. Typed over, a control
				// typing does not change keeps its content; deleted with more, or deleted alone, it goes.
				if (exact && blocker === undefined) continue
				if (exact && inserts) {
					violation ??= REASONS[blocker ?? "notText"]

					continue
				}

				if (info.lock === "sdtLocked" || info.lock === "sdtContentLocked") violation ??= REASONS.contentLocked

				continue
			}

			if (kind === "format") {
				if (blocker === "contentLocked") violation ??= REASONS.contentLocked

				continue
			}

			if (blocker !== undefined && !nestedTypeable(part, doc, from, to, span)) violation ??= REASONS[blocker]
		}
	}
	const visitBlock = (node: Node, pos: number): void => {
		if (kind !== "replace") return

		const info = controlOf(part, node)

		// Deleted whole: the core keeps controls locked against deletion.
		if (info === undefined || (from <= pos && to >= pos + node.nodeSize)) return

		const blocker = typeable(info)

		if (blocker !== undefined && !nestedTypeable(part, doc, from, to, { from: pos + 1, to: pos + node.nodeSize - 1 }))
			violation ??= REASONS[blocker]
	}

	if (from === to) {
		const $pos = doc.resolve(from)

		if ($pos.parent.inlineContent) visitParagraph($pos.parent, $pos.start())

		for (let depth = $pos.depth; depth > 0 && violation === undefined; depth--) {
			const node = $pos.node(depth)

			if (node.type === nodes.sdt) visitBlock(node, $pos.before(depth))
		}

		return violation
	}

	doc.nodesBetween(from, to, (node, pos) => {
		if (violation !== undefined) return false
		if (node.type === nodes.sdt) visitBlock(node, pos)

		if (node.inlineContent) {
			visitParagraph(node, pos + 1)

			return false
		}

		return !node.isInline
	})

	return violation
}

function lockViolation(tr: Transaction, part: StoryPart): string | undefined {
	for (let k = 0; k < tr.steps.length; k++) {
		const step = tr.steps[k]
		const doc = tr.docs[k] as Node
		let violation: string | undefined

		if (step instanceof ReplaceStep || step instanceof ReplaceAroundStep) {
			if (onlyMarkers(step)) continue

			violation =
				step instanceof ReplaceAroundStep
					? (rangeViolation(part, doc, step.from, step.gapFrom, "replace") ??
						rangeViolation(part, doc, step.gapTo, step.to, "replace"))
					: pictureSwap(part, doc, step.from, step.to, step.slice.content)
						? undefined
						: rangeViolation(part, doc, step.from, step.to, "replace", step.slice.size > 0)
		} else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
			const control = step.mark.type === marks.content_control

			// A control mark added by a user is the core's to refuse (an identity).
			if (control && step instanceof AddMarkStep) continue

			violation = rangeViolation(part, doc, step.from, step.to, control ? "unwrap" : "format")
		}

		if (violation !== undefined) return violation
	}

	return undefined
}

// A form command recorded as tracked changes keeps the old content (deleted) next to the new: the control's
// content all takes the state the new content has, so it stays one control.
function unifyStates(tr: Transaction): void {
	const inserted = insertedRanges(tr)
	const seen = new Set<number>()

	for (let k = 0; k < inserted.length; k += 2) {
		const $pos = tr.doc.resolve(inserted[k] ?? 0)

		if (!$pos.parent.inlineContent || seen.has($pos.start())) continue

		seen.add($pos.start())

		const paragraph = $pos.parent
		const start = $pos.start()
		const newest = new Map<string, Mark>()

		paragraph.forEach(child => {
			if (child.marks.some(mark => mark.type === marks.deletion || mark.type === marks.move_from)) return

			for (const mark of child.marks) if (mark.type === marks.content_control) newest.set(controlKey(mark), mark)
		})

		if (newest.size === 0) continue

		let pos = start

		paragraph.forEach(child => {
			for (const mark of child.marks) {
				const wanted = mark.type === marks.content_control ? newest.get(controlKey(mark)) : undefined

				if (wanted !== undefined && !wanted.eq(mark)) {
					tr.removeMark(pos, pos + child.nodeSize, mark)
					tr.addMark(pos, pos + child.nodeSize, wanted)
				}
			}

			pos += child.nodeSize
		})
	}
}

// ---- control content --------------------------------------------------------------------------------------------

// The marks of a control's first inline content (what new content in it takes).
export function contentMarks(doc: Node, around: ControlAround): readonly Mark[] {
	if (around.kind === "inline") return doc.nodeAt(around.from)?.marks ?? [around.holder as Mark]

	let found: readonly Mark[] | undefined

	doc.nodesBetween(around.from, around.to, child => {
		if (found === undefined && child.isInline) found = child.marks

		return found === undefined
	})

	return found ?? []
}

// Replaces a control's content: an inline control's span, or the first paragraph's content of a block one
// (its other paragraphs are emptied).
export function replaceContent(tr: Transaction, around: ControlAround, content: readonly Node[]): void {
	if (around.kind === "inline") {
		tr.replaceWith(around.from, around.to, content)

		return
	}

	const ranges: [number, number][] = []

	tr.doc.nodesBetween(around.from, around.to, (child, pos) => {
		if (!child.inlineContent) return !child.isInline

		ranges.push([pos + 1, pos + child.nodeSize - 1])

		return false
	})

	const [first, ...rest] = ranges

	for (const [from, to] of rest.reverse()) if (to > from) tr.delete(from, to)
	if (first !== undefined) tr.replaceWith(first[0], first[1], content)
}

// Marks for new content of a control: its first content's, with the control mark (inline) holding `state`,
// and the placeholder style only while the placeholder shows.
export function newContentMarks(context: DocxEditContext, doc: Node, around: ControlAround, state: ControlState): readonly Mark[] {
	let set = contentMarks(doc, around)

	if (around.kind === "inline") {
		const mark = around.holder as Mark

		set = marks.content_control.create({ ...mark.attrs, state }).addToSet(mark.removeFromSet(set))
	}

	if (state.placeholder === true) return placeholderMarks(context, set)

	return set.filter(mark => !placeholderStyled(context, mark))
}

// Text as inline content (tabs as tab nodes, line feeds as line breaks).
export function textContent(text: string, set: readonly Mark[]): Node[] {
	const out: Node[] = []

	for (const [k, line] of text.split("\n").entries()) {
		if (k > 0) out.push(nodes.break.create({ type: "textWrapping" }, null, set))

		for (const [n, piece] of line.split("\t").entries()) {
			if (n > 0) out.push(nodes.tab.create(null, null, set))
			if (piece !== "") out.push(docxSchema.text(piece, set))
		}
	}

	return out
}

// The run fonts Word gives a checkbox glyph: w:rFonts naming the state's font for every script.
export function withFont(set: readonly Mark[], font: string | null): readonly Mark[] {
	if (font === null) return set

	const current = marks.font.isInSet(set)

	if (current !== undefined && (current.attrs["ascii"] === font || current.attrs["eastAsia"] === font)) return set

	return marks.font.create({ ascii: font, hAnsi: font, eastAsia: font, hint: "eastAsia" }).addToSet(marks.font.removeFromSet(set))
}

// Replaces a control's content with `text` (a checkbox glyph in its `font`) and sets its state.
export function writeControl(
	tr: Transaction,
	context: DocxEditContext,
	around: ControlAround,
	text: string,
	state: ControlState,
	font: string | null = null
): void {
	const set = withFont(newContentMarks(context, tr.doc, around, state), font)

	replaceContent(tr, around, state.placeholder === true ? placeholderContent(text, set) : textContent(text, set))

	if (around.kind === "block") tr.setNodeAttribute(around.pos, "state", state)
}

// ---- placeholders ---------------------------------------------------------------------------------------

// The character style Word gives placeholder text ("Placeholder Text"), when the document has it.
export function placeholderStyle(context: DocxEditContext): string | undefined {
	return context.styles.byName("Placeholder Text", "character")?.id ?? context.styles.byId("PlaceholderText")?.id
}

// Whether a character style mark is the placeholder style (the document's, or Word's id for it).
export function placeholderStyled(context: DocxEditContext, mark: Mark): boolean {
	if (mark.type !== marks.char_style) return false

	const id = mark.attrs["styleId"] as string

	return id === "PlaceholderText" || id === placeholderStyle(context)
}

// The state of a control whose placeholder shows (`showing`) or no longer does.
export function placeholderState(holder: Node | Mark, info: ControlInfo, showing: boolean): ControlState {
	const state = controlState(holder) ?? {}

	if (!showing) return { ...state, placeholder: false }

	return {
		...state,
		placeholder: true,
		...(info.date !== null ? { date: null } : {}),
		...(info.type === "dropDownList" || info.type === "comboBox" ? { value: null } : {})
	}
}

// Sets a control's state: a block control's attribute, an inline control's mark over its span.
export function setControlState(tr: Transaction, around: ControlAround, state: ControlState): void {
	if (around.kind === "block") {
		tr.setNodeAttribute(around.pos, "state", state)

		return
	}

	const mark = around.holder as Mark

	tr.removeMark(around.from, around.to, mark)
	tr.addMark(around.from, around.to, marks.content_control.create({ ...mark.attrs, state }))
}

// Placeholder text as inline content with `markSet` (a tab as a tab node).
export function placeholderContent(text: string, markSet: readonly Mark[]): Node[] {
	const out: Node[] = []

	text.split("\t").forEach((piece, k) => {
		if (k > 0) out.push(nodes.tab.create(null, null, markSet))
		if (piece !== "") out.push(docxSchema.text(piece, markSet))
	})

	return out
}

// The marks of placeholder text standing for content that had `markSet`: no revision, link or recorded
// formatting change, the placeholder style (no character style when the document lacks it).
export function placeholderMarks(context: DocxEditContext, markSet: readonly Mark[]): readonly Mark[] {
	const style = placeholderStyle(context)
	const kept = markSet.filter(mark => mark.type !== marks.char_style && !DROPPED_MARKS.has(mark.type.name))

	return style === undefined ? kept : marks.char_style.create({ styleId: style }).addToSet(kept)
}

// Ranges each replace step inserted, in the final document.
function insertedRanges(tr: Transaction): number[] {
	const out: number[] = []

	tr.steps.forEach((step, index) => {
		const map = tr.mapping.maps[index] ?? StepMap.empty

		if (map !== StepMap.empty) {
			for (let k = 0; k < out.length; k += 2) {
				out[k] = map.map(out[k] ?? 0, -1)
				out[k + 1] = map.map(out[k + 1] ?? 0, 1)
			}
		}

		if (!(step instanceof ReplaceStep) || step.slice.size === 0) return

		map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
			if (newEnd > newStart) out.push(newStart, newEnd)
		})
	})

	return out
}

interface Emptied {
	readonly mark: Mark
	readonly markSet: readonly Mark[]
	pos: number
}

// Applies the placeholder rules to `tr`; whether it added steps.
function placeholders(tr: Transaction, session: DocxEditorSession, forms: boolean): boolean {
	const part = session.story.part
	const context = session.context
	const before = tr.steps.length
	const emptied: Emptied[] = []
	const inserted = insertedRanges(tr)
	// A selection of exactly a control's content typed over: what is typed goes into the control (the editor's
	// marks would leave it outside, as the control ends where the selection does).
	const selection = session.state.selection
	const covered =
		inserted.length > 0 && !selection.empty && tr.before === session.state.doc
			? inlineControlAround(part, tr.before, selection.from, selection.to)
			: undefined
	const retyped =
		covered !== undefined && covered.from === selection.from && covered.to === selection.to && typeable(covered.info) === undefined
	const coveredKey = retyped ? controlKey(covered.holder) : undefined

	// Inline controls a step deleted exactly (all their content): they come back with their placeholder.
	tr.steps.forEach((step, index) => {
		const map = tr.mapping.maps[index] ?? StepMap.empty

		for (const entry of emptied) entry.pos = map.map(entry.pos, -1)

		if (!(step instanceof ReplaceStep) || step.to <= step.from) return

		const doc = tr.docs[index] as Node
		const $from = doc.resolve(step.from)

		if (!$from.parent.inlineContent || step.to > $from.end()) return

		for (const span of controlSpans($from.parent, $from.start())) {
			if (span.from !== step.from || span.to !== step.to) continue

			const info = controlOf(part, span.mark)

			if (info === undefined || typeable(info) !== undefined || controlKey(span.mark) === coveredKey) continue

			// The marks every node of the content shares: the control's own and those of controls around it (not
			// those of controls inside it, which go with their content).
			let shared: readonly Mark[] | undefined

			doc.nodesBetween(span.from, span.to, node => {
				if (node.isInline) shared = shared === undefined ? node.marks : shared.filter(mark => mark.isInSet(node.marks))
			})

			emptied.push({ mark: span.mark, markSet: shared ?? [span.mark], pos: map.map(step.from, -1) })
		}
	})

	if (retyped && covered !== undefined) {
		for (let k = 0; k < inserted.length; k += 2) {
			const from = inserted[k] ?? 0
			const to = inserted[k + 1] ?? 0

			if (to <= from || !tr.doc.resolve(from).parent.inlineContent) continue

			let inside = false

			tr.doc.nodesBetween(from, to, node => {
				if (node.isInline && node.marks.some(mark => mark.type === marks.content_control && controlKey(mark) === coveredKey))
					inside = true
			})

			if (!inside) tr.addMark(from, to, covered.holder as Mark)
		}
	}

	// Under forms protection, typing at the edge of an inline control goes into the control.
	if (forms) {
		for (let k = 0; k < inserted.length; k += 2) {
			const from = inserted[k] ?? 0
			const to = inserted[k + 1] ?? 0
			const doc = tr.doc

			if (from === to || !doc.resolve(from).parent.inlineContent || inlineControlAround(part, doc, from, to) !== undefined) continue

			const edge = inlineControlAround(part, doc, from, from) ?? inlineControlAround(part, doc, to, to)

			if (edge !== undefined && typeable(edge.info) === undefined) tr.addMark(from, to, edge.holder as Mark)
		}
	}

	// Typed into a control showing its placeholder: the placeholder text and its style go.
	const seen = new Set<string>()

	for (let k = inserted.length - 2; k >= 0; k -= 2) {
		const doc = tr.doc
		const mapping = tr.mapping.slice(before)
		const from = mapping.map(inserted[k] ?? 0, 1)
		const to = mapping.map(inserted[k + 1] ?? 0, -1)

		if (to <= from) continue

		const around = inlineControlAround(part, doc, from, to) ?? blockControlAround(part, doc, from, to)

		if (around === undefined || !around.info.showingPlaceholder || typeable(around.info) !== undefined) continue

		const key = controlKey(around.holder)

		if (seen.has(key)) continue

		seen.add(key)
		replacePlaceholder(tr, session, around, from, to)
	}

	// Emptied block controls (every paragraph empty): the placeholder again.
	const points = inserted.length > 0 ? inserted.filter((_pos, k) => k % 2 === 0) : [tr.selection.from]

	for (const point of points) {
		const doc = tr.doc
		const at = Math.min(tr.mapping.slice(before).map(point, -1), doc.content.size)
		const around = blockControlAround(part, doc, at, at)

		if (around === undefined || around.info.showingPlaceholder || typeable(around.info) !== undefined) continue
		if (emptyContent(around.holder as Node)) restoreBlock(tr, context, part, around)
	}

	for (const entry of emptied.reverse()) {
		const doc = tr.doc
		const pos = Math.min(tr.mapping.slice(before).map(entry.pos, -1), doc.content.size)
		const $pos = doc.resolve(pos)
		const key = controlKey(entry.mark)
		const info = controlOf(part, entry.mark)

		if (info === undefined || !$pos.parent.inlineContent) continue
		if (controlSpans($pos.parent, $pos.start()).some(span => controlKey(span.mark) === key)) continue

		const mark = marks.content_control.create({ ...entry.mark.attrs, state: placeholderState(entry.mark, info, true) })
		const markSet = placeholderMarks(context, mark.addToSet(entry.mark.removeFromSet(entry.markSet)))

		tr.insert(pos, placeholderContent(placeholderText(part, entry.mark, info), markSet))
	}

	return tr.steps.length > before
}

function emptyContent(node: Node): boolean {
	let empty = true

	node.descendants(child => {
		if (child.isInline) empty = false

		return empty && !child.isInline
	})

	return empty
}

function replacePlaceholder(tr: Transaction, session: DocxEditorSession, around: ControlAround, from: number, to: number): void {
	const doc = tr.doc
	const holder = around.holder
	const info = around.info
	// What else the control holds goes (from the end).
	const pieces: [number, number][] = []

	if (around.kind === "inline") {
		if (to < around.to) pieces.push([to, around.to])
		if (around.from < from) pieces.push([around.from, from])
	} else {
		doc.nodesBetween(around.from, around.to, (node, pos) => {
			if (!node.inlineContent) return !node.isInline

			const start = pos + 1
			const end = pos + node.nodeSize - 1

			if (start < from) pieces.push([start, Math.min(end, from)])
			if (end > to) pieces.push([Math.max(start, to), end])

			return false
		})
		pieces.sort((a, b) => b[0] - a[0])
	}

	let start = from
	let end = to

	for (const [a, b] of pieces) {
		if (b <= a) continue

		tr.delete(a, b)

		if (b <= start) {
			start -= b - a
			end -= b - a
		}
	}

	tr.doc.nodesBetween(start, end, (node, pos) => {
		for (const mark of node.marks) {
			if (placeholderStyled(session.context, mark)) tr.removeMark(Math.max(start, pos), Math.min(end, pos + node.nodeSize), mark)
		}
	})

	const mapped = mappedAround(tr, around, start, end)

	if (mapped === undefined) return

	setControlState(tr, mapped, placeholderState(holder, info, false))

	if (info.temporary) unwrap(tr, mappedAround(tr, mapped, start, end) ?? mapped)
}

// The control `around` stood for, as it stands now around [from, to].
function mappedAround(tr: Transaction, around: ControlAround, from: number, to: number): ControlAround | undefined {
	const doc = tr.doc
	const $pos = doc.resolve(from)
	const key = controlKey(around.holder)

	if (around.kind === "block") {
		for (let depth = $pos.depth; depth > 0; depth--) {
			const node = $pos.node(depth)

			if (node.type === nodes.sdt && controlKey(node) === key)
				return { ...around, holder: node, pos: $pos.before(depth), from: $pos.before(depth) + 1, to: $pos.after(depth) - 1 }
		}

		return undefined
	}

	if (!$pos.parent.inlineContent) return undefined

	const span = controlSpans($pos.parent, $pos.start()).find(item => controlKey(item.mark) === key && item.from <= from && item.to >= to)

	return span === undefined ? undefined : { ...around, holder: span.mark, from: span.from, to: span.to, pos: span.from }
}

// A temporary control (w:temporary) goes once edited; its content stays.
function unwrap(tr: Transaction, around: ControlAround): void {
	if (around.kind === "inline") {
		tr.removeMark(around.from, around.to, around.holder as Mark)

		return
	}

	const node = tr.doc.nodeAt(around.pos)

	if (node !== null && node.type === nodes.sdt) tr.replaceWith(around.pos, around.pos + node.nodeSize, node.content)
}

function restoreBlock(tr: Transaction, context: DocxEditContext, part: StoryPart, around: ControlAround): void {
	const node = around.holder as Node
	let at = -1

	tr.doc.nodesBetween(around.from, around.to, (child, pos) => {
		if (at >= 0) return false
		if (child.inlineContent) at = pos + 1

		return at < 0 && !child.isInline
	})

	if (at < 0) return

	// The formatting of what was there: the first inline content the transaction took.
	const original = tr.before
	const back = Math.min(tr.mapping.invert().map(at, 1), original.content.size)
	let markSet: readonly Mark[] = []

	original.nodesBetween(back, Math.min(original.content.size, back + 4096), child => {
		// Controls inside the block go with their content.
		if (markSet.length === 0 && child.isInline) markSet = child.marks.filter(mark => mark.type !== marks.content_control)

		return markSet.length === 0 && !child.isInline
	})

	tr.insert(at, placeholderContent(placeholderText(part, node, around.info), placeholderMarks(context, markSet)))
	tr.setNodeAttribute(around.pos, "state", placeholderState(node, around.info, true))
}
