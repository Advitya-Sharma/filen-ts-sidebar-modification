import { Item, RevisionView, TRANSPARENT } from "./kinds"
import type { StoryPart } from "./part"

const HIDDEN_BY_VIEW: readonly Uint8Array[] = [RevisionView.Markup, RevisionView.Final, RevisionView.Original].map(view => {
	const hidden = new Uint8Array(128)

	if (view === RevisionView.Final) {
		hidden[Item.Deletion] = 1
		hidden[Item.MoveFrom] = 1
	} else if (view === RevisionView.Original) {
		hidden[Item.Insertion] = 1
		hidden[Item.MoveTo] = 1
	}

	return hidden
})

// Allocation-free pre-order walk over the children of one container. Transparent containers (runs and
// wrappers: sdt, customXml, smartTag, hyperlink, fldSimple, ins/del/move, bidi, AlternateContent) are
// returned and then entered; every other item is returned with its subtree skipped, so walking a story
// yields its blocks (descend into a table or paragraph with a second walker), and walking a paragraph
// yields runs, wrappers and inline leaves. Tracked changes hidden by the view (revision wrappers, inserted
// or deleted rows and cells) are skipped whole; paragraphs whose mark is hidden are still returned, see
// StoryPart.paragraphMarkHidden.
export class ItemWalker {
	private kinds: Uint8Array = new Uint8Array(0)
	private ends: Int32Array = new Int32Array(0)
	private cursor = 0
	private limit = 0
	private part: StoryPart | null = null
	private view: RevisionView
	private hidden: Uint8Array

	constructor(view: RevisionView = RevisionView.Markup) {
		this.view = view
		this.hidden = HIDDEN_BY_VIEW[view] ?? new Uint8Array(128)
	}

	setView(view: RevisionView): void {
		this.view = view
		this.hidden = HIDDEN_BY_VIEW[view] ?? this.hidden
	}

	// Walk the children of `container`.
	reset(part: StoryPart, container: number): this {
		return this.resetRange(part, container + 1, part.end(container))
	}

	// Walk items [start, end), which must be a run of whole sibling subtrees (a section's blocks, say).
	resetRange(part: StoryPart, start: number, end: number): this {
		this.part = part
		this.kinds = part.kinds
		this.ends = part.ends
		this.cursor = start
		this.limit = end

		return this
	}

	next(): number {
		const kinds = this.kinds
		const ends = this.ends
		const hidden = this.hidden

		while (this.cursor < this.limit) {
			const item = this.cursor
			const kind = kinds[item] ?? 0

			if (
				hidden[kind] === 1 ||
				(this.view !== RevisionView.Markup &&
					(kind === Item.Row || kind === Item.Cell) &&
					this.part?.hidden(item, this.view) === true)
			) {
				this.cursor = ends[item] ?? item + 1

				continue
			}

			this.cursor = TRANSPARENT[kind] === 1 ? item + 1 : (ends[item] ?? item + 1)

			return item
		}

		return -1
	}

	// Skip the rest of a transparent container that was returned (its children are not walked).
	skip(item: number): void {
		this.cursor = Math.max(this.cursor, this.ends[item] ?? this.cursor)
	}
}
