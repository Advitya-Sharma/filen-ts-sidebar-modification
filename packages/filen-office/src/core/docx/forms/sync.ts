import type { Mark, Node } from "prosemirror-model"
import type { EditorState, Transaction } from "prosemirror-state"
import type { DocxEditContext } from "../edit/context"
import { markInternal } from "../edit/invariants/core"
import { boundAt, controlSpans, controlText, showValue, touchedPoints } from "../edit/invariants/forms"
import { docxSchema } from "../edit/schema"
import type { DocxEditorSession } from "../edit/session"
import { storySessionOf } from "./sessions"
import type { StoryPart } from "../model/index"
import { boundValue, partHasBindings } from "./binding"
import { controlOf, simpleBinding, sourceControl, type ControlAround, type ControlInfo } from "./controls"

// Controls in different stories bound to one node (a document title in the body and the header, a date in a
// footer): Word shows the node's value in all of them, so an edit of one shows in the others at once. After a
// story's session applied an edit that changed what a bound control holds, the sessions of the other stories
// holding controls bound to its node (made by the document's host when not open yet) get the value too, in an
// undo-less change of their own: the edit stays one undo step of the story it was made in, and undoing it there
// changes the value again, which reaches the others the same way. Within one story the forms stage keeps controls
// alike (invariants/forms.ts mirrorBindings).

const nodes = docxSchema.nodes

// Bound controls of the document's stories by binding key, those bound to a node controls of other stories are
// bound to too.
interface StoryBindings {
	readonly byStory: ReadonlyMap<number, ReadonlySet<string>>
	readonly byKey: ReadonlyMap<string, readonly number[]>
}

const NO_BINDINGS: StoryBindings = { byStory: new Map(), byKey: new Map() }
const bindingsOf = new WeakMap<DocxEditContext, StoryBindings>()

function keyOf(info: ControlInfo | undefined): string | undefined {
	return info?.binding === null || info === undefined || info.richBinding ? undefined : simpleBinding(info.binding)?.key
}

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
			const key = keyOf(sourceControl(part, item))

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
	const byStory = new Map<number, Set<string>>()

	for (const [key, set] of byKey) {
		if (set.size < 2) continue

		shared.set(key, [...set])

		for (const story of set) {
			const keys = byStory.get(story) ?? new Set()

			keys.add(key)
			byStory.set(story, keys)
		}
	}

	found = shared.size === 0 ? NO_BINDINGS : { byStory, byKey: shared }
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
	if (watched.has(session)) return

	watched.add(session)

	const keys = storyBindings(session.context).byStory.get(session.story.story.index)

	if (keys === undefined) return

	session.onApplied((transactions, state) => passOn(session, keys, transactions, state))
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

	// In document order, stopping once every key has its value.
	doc.descendants((node, pos) => {
		if (out.size === touched.size) return false

		if (node.type === nodes.sdt) {
			const info = controlOf(part, node)
			const key = keyOf(info)

			if (info !== undefined && key !== undefined && touched.has(key) && !out.has(key))
				out.set(
					key,
					info.showingPlaceholder
						? ""
						: boundValue(
								info,
								controlText(doc, { kind: "block", holder: node, info, from: pos + 1, to: pos + node.nodeSize - 1, pos })
							)
				)
		}

		if (!node.inlineContent) return !node.isInline && node.type !== nodes.locked_block

		for (const span of controlSpans(node, pos + 1)) {
			const info = controlOf(part, span.mark)
			const key = keyOf(info)

			if (info !== undefined && key !== undefined && touched.has(key) && !out.has(key))
				out.set(
					key,
					info.showingPlaceholder
						? ""
						: boundValue(
								info,
								controlText(doc, { kind: "inline", holder: span.mark, info, from: span.from, to: span.to, pos: span.from })
							)
				)
		}

		return false
	})

	return out
}

function passOn(session: DocxEditorSession, keys: ReadonlySet<string>, transactions: readonly Transaction[], state: EditorState): void {
	const context = session.context
	const first = transactions[0]

	if (first === undefined || mirrors.has(first) || syncing.has(context) || !transactions.some(tr => tr.docChanged)) return

	const values = touchedValues(session, keys, transactions, state)

	if (values.size === 0) return

	const byKey = storyBindings(context).byKey
	const targets = new Map<number, Map<string, string>>()
	const own = session.story.story.index

	for (const [key, value] of values) {
		for (const story of byKey.get(key) ?? []) {
			if (story === own) continue

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

// Where a story's controls of each binding key are: block controls by their position, inline ones by their
// paragraph's. Kept per document; a document this module changed gets its predecessor's mapped.
interface Anchors {
	readonly blocks: Map<string, number[]>
	readonly paragraphs: Map<string, number[]>
}

const anchorsOf = new WeakMap<Node, Anchors>()

function push(map: Map<string, number[]>, key: string, pos: number): void {
	const list = map.get(key)

	if (list === undefined) map.set(key, [pos])
	else if (list.at(-1) !== pos) list.push(pos)
}

function anchors(part: StoryPart, doc: Node, keys: ReadonlySet<string>): Anchors {
	const known = anchorsOf.get(doc)

	if (known !== undefined) return known

	const out: Anchors = { blocks: new Map(), paragraphs: new Map() }

	doc.descendants((node, pos) => {
		if (node.type === nodes.sdt) {
			const key = keyOf(controlOf(part, node))

			if (key !== undefined && keys.has(key)) push(out.blocks, key, pos)
		}

		if (!node.inlineContent) return !node.isInline && node.type !== nodes.locked_block

		for (const span of controlSpans(node, pos + 1)) {
			const key = keyOf(controlOf(part, span.mark))

			if (key !== undefined && keys.has(key)) push(out.paragraphs, key, pos)
		}

		return false
	})

	anchorsOf.set(doc, out)

	return out
}

function mapped(map: Map<string, number[]>, tr: Transaction): Map<string, number[]> {
	return new Map([...map].map(([key, list]) => [key, list.map(pos => tr.mapping.map(pos, 1))]))
}

// The controls of `key` in `doc`, from its anchors.
function controlsOf(part: StoryPart, doc: Node, found: Anchors, key: string): ControlAround[] {
	const out: ControlAround[] = []

	for (const pos of found.blocks.get(key) ?? []) {
		const node = doc.nodeAt(pos)
		const info = node?.type === nodes.sdt ? controlOf(part, node) : undefined

		if (node !== null && info !== undefined && keyOf(info) === key)
			out.push({ kind: "block", holder: node, info, from: pos + 1, to: pos + node.nodeSize - 1, pos })
	}

	for (const pos of found.paragraphs.get(key) ?? []) {
		const node = doc.nodeAt(pos)

		if (node === null || !node.inlineContent) continue

		for (const span of controlSpans(node, pos + 1)) {
			const info = controlOf(part, span.mark as Mark)

			if (info !== undefined && keyOf(info) === key)
				out.push({ kind: "inline", holder: span.mark, info, from: span.from, to: span.to, pos: span.from })
		}
	}

	return out
}

// Shows `values` (by binding key) in the controls of `target` bound to them: one change undo leaves alone.
function mirror(target: DocxEditorSession, values: ReadonlyMap<string, string>): void {
	const part = target.story.part
	const keys = storyBindings(target.context).byStory.get(target.story.story.index)

	if (keys === undefined) return

	const state = target.state
	const found = anchors(part, state.doc, keys)
	const changes: { around: ControlAround; value: string }[] = []

	for (const [key, value] of values) {
		for (const around of controlsOf(part, state.doc, found, key)) {
			const shown = around.info.showingPlaceholder ? "" : boundValue(around.info, controlText(state.doc, around))

			if (shown !== value) changes.push({ around, value })
		}
	}

	if (changes.length === 0) return

	const tr = state.tr

	for (const { around, value } of changes.sort((a, b) => b.around.from - a.around.from))
		showValue(tr, target.context, part, around, value)

	if (!tr.docChanged) return

	mirrors.add(markInternal(tr.setMeta("addToHistory", false)))

	if (target.dispatch(tr) && target.state.doc === tr.doc) {
		anchorsOf.set(tr.doc, { blocks: mapped(found.blocks, tr), paragraphs: mapped(found.paragraphs, tr) })
	}
}
