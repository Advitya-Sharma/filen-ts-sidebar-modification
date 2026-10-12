import type { Node } from "prosemirror-model"
import type { EditorState, Transaction } from "prosemirror-state"
import {
	AddMarkStep,
	AddNodeMarkStep,
	AttrStep,
	RemoveMarkStep,
	RemoveNodeMarkStep,
	ReplaceAroundStep,
	ReplaceStep,
	type Step
} from "prosemirror-transform"
import { Ns } from "../model/namespaces"
import type { StoryPart } from "../model/index"
import { controlAround, pictureSwap, typeable } from "../forms/controls"
import { SetAttrsStep } from "./attrs"
import type { DocxEditorCommandType } from "./commands"
import type { DocxEditContext } from "./context"
import { docxSchema, NO_SOURCE, type MarkerKind, type MarkerList } from "./schema"
import type { DocxEditorSession } from "./session"

// Document protection that allows editing within limits (ECMA-376 Part 1 §17.15.1.29 w:documentProtection
// with w:enforcement on): "forms" lets only content controls and legacy form fields change, plus the ranges an
// editing permission opens to everyone (§17.13.7.2 w:permStart w:edGrp "everyone" or "editors"); "comments"
// lets only comments change, plus those ranges; "trackedChanges" records every edit as a tracked change: the
// switch cannot be turned off and changes cannot be accepted or rejected (Word's behaviour). "readOnly" keeps
// the document view-only (gate 0 reason "protected"). The password hash is kept as it is: editing within the
// limits needs no password.

export type DocxProtectionMode = "none" | "forms" | "comments" | "trackedChanges"

export interface DocxProtection {
	readonly mode: DocxProtectionMode
	// Tracking cannot be switched off (trackedChanges).
	readonly forcesTracking: boolean
	// The protection carries a password hash (kept as written).
	readonly hashed: boolean
}

export const NO_PROTECTION: DocxProtection = { mode: "none", forcesTracking: false, hashed: false }

const LIMITED_MODES = new Set<string>(["forms", "comments", "trackedChanges"])
const protections = new WeakMap<DocxEditContext, DocxProtection>()

// The mode an enforced w:documentProtection w:edit opens the document in ("none" for the view-only and
// unprotected ones).
export function protectionMode(edit: string | undefined, enforced: boolean): DocxProtectionMode {
	return enforced && edit !== undefined && LIMITED_MODES.has(edit) ? (edit as DocxProtectionMode) : "none"
}

export function protectionOf(context: DocxEditContext): DocxProtection {
	let found = protections.get(context)

	if (found === undefined) {
		const protection = context.model.package.facts.documentProtection
		const mode = protection === undefined ? "none" : protectionMode(protection.edit, protection.enforced)

		found =
			mode === "none"
				? NO_PROTECTION
				: Object.freeze({ mode, forcesTracking: mode === "trackedChanges", hashed: protection?.hashed ?? false })
		protections.set(context, found)
	}

	return found
}

// Commands each mode allows; any other write command is refused (undo and redo stay: they only take back what
// the mode allowed).
const FORM_FILLING = new Set<DocxEditorCommandType>([
	"control.check",
	"control.choose",
	"control.date",
	"control.clear",
	"control.text",
	"control.item.add",
	"control.item.remove"
])
const COMMENTING = new Set<DocxEditorCommandType>([
	"comment.add",
	"comment.reply",
	"comment.edit",
	"comment.resolve",
	"comment.reopen",
	"comment.delete",
	"comment.next",
	"comment.prev"
])
const NOT_WHILE_TRACKED = new Set<DocxEditorCommandType>([
	"review.track",
	"review.accept",
	"review.reject",
	"review.acceptAll",
	"review.rejectAll"
])

// Whether write command `type` may run on `session` (the story being edited) under the document's protection.
// Other story commands than the mode's own run only where the selection lies in a range open to everyone.
export function commandAllowed(session: DocxEditorSession, type: DocxEditorCommandType): boolean {
	const protection = protectionOf(session.context)

	switch (protection.mode) {
		case "none":
			return true
		case "trackedChanges":
			return !NOT_WHILE_TRACKED.has(type)
		case "forms":
			if (FORM_FILLING.has(type)) return true
			break
		case "comments":
			if (COMMENTING.has(type) || session.story.story.kind === "comment") return true
			break
	}

	if (type === "history.undo" || type === "history.redo") return true

	const selection = session.state.selection

	return permissionRanges(session).covers(session.state.doc, selection.from, selection.to)
}

// ---- editing permission ranges --------------------------------------------------------------------------

const OPEN_GROUPS = new Set(["everyone", "editors"])

// Ranges of a story's document an editing permission opens to everyone (pairs of positions), kept through
// every transaction the session applies: the ranges' markers always survive, so their positions move with the
// edits.
export class PermissionRanges {
	private ranges: number[]
	private doc: Node

	constructor(session: DocxEditorSession) {
		this.doc = session.state.doc
		this.ranges = openRanges(session.story.part, this.doc)
		session.onApplied((transactions, state) => {
			for (const tr of transactions) {
				if (!tr.docChanged || this.ranges.length === 0) continue

				for (let k = 0; k < this.ranges.length; k += 2) {
					this.ranges[k] = tr.mapping.map(this.ranges[k] ?? 0, -1)
					this.ranges[k + 1] = tr.mapping.map(this.ranges[k + 1] ?? 0, 1)
				}
			}

			this.doc = state.doc
		})
	}

	// Whether [from, to] lies in an open range of `ranges` (by default the session document's).
	covers(doc: Node, from: number, to: number, ranges: readonly number[] = this.of(doc)): boolean {
		for (let k = 0; k < ranges.length; k += 2) {
			if ((ranges[k] ?? 0) <= from && to <= (ranges[k + 1] ?? -1)) return true
		}

		return false
	}

	// The ranges of `doc` when it is the session's document (none for any other).
	of(doc: Node): readonly number[] {
		return doc === this.doc ? this.ranges : []
	}
}

const permissions = new WeakMap<DocxEditorSession, PermissionRanges>()

export function permissionRanges(session: DocxEditorSession): PermissionRanges {
	let found = permissions.get(session)

	if (found === undefined) {
		found = new PermissionRanges(session)
		permissions.set(session, found)
	}

	return found
}

// Whether w:permStart item `item` opens its range to everyone (a range for one named user, w:ed, does not:
// the editor cannot tell who edits).
function openStart(part: StoryPart, item: number): boolean {
	if (item === NO_SOURCE || item >= part.count) return false

	const group = part.reader.attr(part.node(item), Ns.W, "edGrp")

	return group !== undefined && OPEN_GROUPS.has(group)
}

// The open ranges of `doc`, by one walk over its markers (once per session, under protection only).
function openRanges(part: StoryPart, doc: Node): number[] {
	const starts = new Map<string, number>()
	const pairs: [number, number][] = []
	const visit = (list: MarkerList, pos: number): void => {
		for (const marker of list) {
			if (marker.id === null) continue

			if (marker.kind === "permissionStart") {
				if (openStart(part, marker.src)) starts.set(marker.id, pos)
			} else if (marker.kind === "permissionEnd") {
				const start = starts.get(marker.id)

				if (start !== undefined) {
					pairs.push([start, pos])
					starts.delete(marker.id)
				}
			}
		}
	}

	const walk = (node: Node, pos: number): void => {
		if (node.type === docxSchema.nodes.marker) {
			const kind = node.attrs["kind"] as MarkerKind

			if (kind === "permissionStart" || kind === "permissionEnd")
				visit([{ src: node.attrs["src"] as number, kind, id: node.attrs["id"] as string | null }], pos + 1)

			return
		}

		if (!node.isBlock) return

		const before = node.attrs["before"] as MarkerList | undefined
		const end = node.attrs["end"] as MarkerList | undefined

		if (before !== undefined && before.length > 0) visit(before, pos)
		if (!node.isAtom) node.forEach((child, offset) => walk(child, pos + 1 + offset))
		if (end !== undefined && end.length > 0 && !node.isTextblock) visit(end, pos + node.nodeSize - 1)
	}

	doc.forEach((child, offset) => walk(child, offset))

	const end = doc.attrs["end"] as MarkerList | undefined

	if (end !== undefined && end.length > 0) visit(end, doc.content.size)

	return pairs.sort((a, b) => a[0] - b[0]).flat()
}

// ---- transactions ---------------------------------------------------------------------------------------

// What a step touches, as [from, to] ranges of the document it applies to; undefined for steps no protection
// lets through (document attributes).
function stepRanges(step: Step, doc: Node): number[] | undefined {
	if (step instanceof ReplaceStep || step instanceof ReplaceAroundStep) return [step.from, step.to]
	if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) return [step.from, step.to]

	if (step instanceof AttrStep || step instanceof AddNodeMarkStep || step instanceof RemoveNodeMarkStep) {
		const node = doc.nodeAt(step.pos)

		return node === null ? undefined : [step.pos, step.pos + node.nodeSize]
	}

	if (step instanceof SetAttrsStep) {
		const out: number[] = []

		step.eachNode(doc, (node, pos) => out.push(pos, pos + node.nodeSize))

		return out
	}

	return undefined
}

// A form command setting the state of a block control, or the state and result of a legacy form field.
function controlStateStep(step: Step, doc: Node): boolean {
	if (!(step instanceof AttrStep)) return false

	const node = doc.nodeAt(step.pos)

	if (node?.type === docxSchema.nodes.sdt) return step.attr === "state"

	return node?.type === docxSchema.nodes.field && (step.attr === "state" || step.attr === "result")
}

// Why `tr` (a transaction about to be applied to `state`, the session's) edits outside what the document's
// protection allows, or undefined. `form`: a form command's, which may change any control's content.
export function protectionViolation(
	tr: Transaction,
	session: DocxEditorSession,
	form = false,
	state: EditorState = session.state
): string | undefined {
	const mode = protectionOf(session.context).mode

	if (mode === "none" || mode === "trackedChanges" || !tr.docChanged) return undefined
	if (mode === "comments" && session.story.story.kind === "comment") return undefined

	const index = permissionRanges(session)
	let ranges: readonly number[] = state.doc === tr.before ? index.of(tr.before) : []
	const part = session.story.part

	for (let k = 0; k < tr.steps.length; k++) {
		const step = tr.steps[k] as Step
		const doc = tr.docs[k] as Node
		const touched = stepRanges(step, doc)

		if (touched === undefined) return "protected"

		for (let r = 0; r < touched.length; r += 2) {
			const from = touched[r] ?? 0
			const to = touched[r + 1] ?? 0

			if (index.covers(doc, from, to, ranges)) continue
			if (
				mode === "forms" &&
				(controlStateStep(step, doc) || formFillable(part, doc, from, to, form) || swapsPicture(part, doc, step))
			)
				continue

			return "protected"
		}

		const map = tr.mapping.maps[k]

		if (map !== undefined && ranges.length > 0) ranges = ranges.map((pos, n) => map.map(pos, n % 2 === 0 ? -1 : 1))
	}

	return undefined
}

// A picture control's picture replaced (Word lets a form's user change it).
function swapsPicture(part: StoryPart, doc: Node, step: Step): boolean {
	return step instanceof ReplaceStep && pictureSwap(part, doc, step.from, step.to, step.slice.content)
}

// Whether [from, to] lies inside a content control a user may type into (`form`: one a form command changes,
// whose content is not locked).
function formFillable(part: StoryPart, doc: Node, from: number, to: number, form: boolean): boolean {
	const around = controlAround(part, doc, from, to)
	const blocker = around === undefined ? "none" : typeable(around.info)

	return blocker === undefined || (form && blocker === "notText")
}
