import type { EditorState, Transaction } from "prosemirror-state"
import type { DocxEditContext } from "../edit/context"
import { markInternal } from "../edit/invariants/core"
import { exemptFromProtection } from "../edit/protection"
import { boundAt, controlText, showValue, touchedPoints } from "../edit/invariants/forms"
import type { DocxEditorSession } from "../edit/session"
import { liveSessions, storySessionOf } from "./sessions"
import type { StoryPart } from "../model/index"
import { boundValue, noteBoundEdit, partHasBindings, unmappedKeys } from "./binding"
import { sourceControl, type ControlAround } from "./controls"
import { bindingKey, boundAround, boundIndex, followIndex, holdsBindings } from "./positions"

// Controls in different stories bound to one node (a document title in the body and the header, a date in a
// footer): Word shows the node's value in all of them, so an edit of one shows in the others at once. After a
// story's session applied an edit that changed what a bound control holds, the sessions of the other stories
// holding controls bound to its node (made by the document's host when not open yet) get the value too, in an
// undo-less change of their own: the edit stays one undo step of the story it was made in, and undoing it there
// changes the value again, which reaches the others the same way. Within one story the forms stage keeps controls
// alike (invariants/forms.ts mirrorBindings).

// The source stories holding controls bound to each node controls of more than one story are bound to, by
// binding key.
type StoryBindings = ReadonlyMap<string, readonly number[]>

const bindingsOf = new WeakMap<DocxEditContext, StoryBindings>()

// Each source control counts for the innermost story holding it (a text box's, not its holder's). Read once per
// document, from the parts holding bindings.
function storyBindings(context: DocxEditContext): StoryBindings {
	let found = bindingsOf.get(context)

	if (found !== undefined) return found

	const model = context.model
	const stories = new Map<number, { story: number; from: number; to: number }[]>()

	for (const story of model.stories) {
		const part = model.parts[story.part]

		if (part === undefined || !partHasBindings(part)) continue

		const list = stories.get(story.part) ?? []

		list.push({ story: story.index, from: story.item < 0 ? 0 : story.item, to: story.item < 0 ? part.count : part.end(story.item) })
		stories.set(story.part, list)
	}

	const byKey = new Map<string, Set<number>>()

	for (const [index, list] of stories) {
		const part = model.parts[index] as StoryPart

		for (let row = 0; row < part.sdts.count; row++) {
			const item = part.sdts.item[row] ?? -1
			const key = bindingKey(sourceControl(part, item))

			if (key === undefined) continue

			let holder: { story: number; from: number } | undefined

			for (const entry of list)
				if (item >= entry.from && item < entry.to && (holder === undefined || entry.from >= holder.from)) holder = entry
			if (holder === undefined) continue

			const set = byKey.get(key) ?? new Set()

			set.add(holder.story)
			byKey.set(key, set)
		}
	}

	const shared = new Map<string, number[]>()

	for (const [key, set] of byKey) if (set.size > 1) shared.set(key, [...set])

	found = shared
	bindingsOf.set(context, found)

	return found
}

// Changes this module made (a session hearing one does not pass it on).
const mirrors = new WeakSet<Transaction>()
const watched = new WeakSet<DocxEditorSession>()
// Documents whose stories are being brought in line (no change made meanwhile is passed on again).
const syncing = new WeakSet<DocxEditContext>()

// Called for each transaction a session transforms: a session whose story holds controls bound to a node other
// stories' controls are bound to passes the values its edits leave on to them from now on.
export function watchBoundStories(session: DocxEditorSession): void {
	if (watched.has(session) || !holdsBindings(session)) return

	watched.add(session)

	const part = session.story.part

	// The positions of the story's bound controls follow every change (forms/positions.ts), and its values go to
	// the other stories holding controls of their nodes: source ones, and copies pasted or made since.
	session.onApplied((transactions, state) => {
		followIndex(part, transactions)
		const unmapped = unmappedKeys(session.context)

		passOn(session, new Set([...boundIndex(part, state.doc).keys()].filter(key => !unmapped.has(key))), transactions, state)
	})
}

// The stories other than `own` holding controls bound to `key`: from the source, and the open sessions' own.
function storiesOf(context: DocxEditContext, key: string, own: number): number[] {
	const out = new Set(storyBindings(context).get(key) ?? [])

	for (const session of liveSessions(context)) {
		const part = session.story.part

		if (holdsBindings(session) && boundIndex(part, session.state.doc).has(key)) out.add(session.story.story.index)
	}

	out.delete(own)

	return [...out]
}

// The values the story's controls of the keys `transactions` touched hold now: each key's first control in the
// story, its first piece (what a save stores in the node, forms/binding.ts boundControls; the forms stage keeps
// the story's controls of a key alike).
function touchedValues(
	session: DocxEditorSession,
	keys: ReadonlySet<string>,
	transactions: readonly Transaction[],
	state: EditorState
): Map<string, string> {
	const part = session.story.part
	const doc = state.doc
	const touched = new Set<string>()

	transactions.forEach((tr, index) => {
		let at = touchedPoints(tr)

		for (let later = index + 1; later < transactions.length; later++) {
			const mapping = (transactions[later] as Transaction).mapping

			at = at.map(pos => mapping.map(pos, -1))
		}

		for (const point of at) for (const found of boundAt(part, doc, Math.min(point, doc.content.size), keys)) touched.add(found.key)
	})

	const out = new Map<string, string>()

	if (touched.size === 0) return out

	const index = boundIndex(part, doc)

	for (const key of touched) {
		const first = boundAround(part, doc, index, key)[0]

		if (first !== undefined) out.set(key, first.info.showingPlaceholder ? "" : boundValue(first.info, controlText(doc, first)))
	}

	return out
}

function passOn(session: DocxEditorSession, keys: ReadonlySet<string>, transactions: readonly Transaction[], state: EditorState): void {
	const context = session.context
	const first = transactions[0]

	if (first === undefined || mirrors.has(first) || syncing.has(context) || !transactions.some(tr => tr.docChanged)) return

	const values = touchedValues(session, keys, transactions, state)

	if (values.size === 0) return

	for (const [key, value] of values) noteBoundEdit(context, key, value)

	const targets = new Map<number, Map<string, string>>()
	const own = session.story.story.index

	for (const [key, value] of values) {
		for (const story of storiesOf(context, key, own)) {
			const wanted = targets.get(story) ?? new Map()

			wanted.set(key, value)
			targets.set(story, wanted)
		}
	}

	syncing.add(context)

	try {
		for (const [story, wanted] of targets) {
			const target = storySessionOf(context, story)

			if (target !== undefined && target !== session && !target.readOnly) mirror(target, wanted)
		}
	} finally {
		syncing.delete(context)
	}
}

// Shows `values` (by binding key) in the controls of `target` bound to them: one change undo leaves alone.
function mirror(target: DocxEditorSession, values: ReadonlyMap<string, string>): void {
	const part = target.story.part
	const state = target.state
	const index = boundIndex(part, state.doc)
	const changes: { around: ControlAround; value: string }[] = []

	for (const [key, value] of values) {
		for (const around of boundAround(part, state.doc, index, key)) {
			const shown = around.info.showingPlaceholder ? "" : boundValue(around.info, controlText(state.doc, around))

			if (shown !== value) changes.push({ around, value })
		}
	}

	if (changes.length === 0) return

	// A control one written holds goes with that one's content (one span under two control marks included), as in
	// the session's own invariant; the rest change from the end.
	changes.sort((a, b) => a.around.from - b.around.from || b.around.to - a.around.to)

	const written: ControlAround[] = []
	const kept: typeof changes = []

	for (const change of changes) {
		if (written.some(other => other.from <= change.around.from && other.to >= change.around.to)) continue

		written.push(change.around)
		kept.push(change)
	}

	const tr = state.tr

	for (const { around, value } of kept.reverse()) showValue(tr, target.context, part, around, value)

	if (!tr.docChanged) return

	mirrors.add(exemptFromProtection(markInternal(tr.setMeta("addToHistory", false))))
	target.dispatch(tr)
}
