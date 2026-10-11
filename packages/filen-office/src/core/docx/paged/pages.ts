import { Region, SOURCE_INTS, SourceInt, type PageDisplayList } from "../layout/index"
import type { StoryEdits, TopEntry } from "./edits"

// Which pages show which top-level blocks of a story, from the pages' sources (paragraphs laid out on each
// page): where layout restarts after an edit and where a position of the editing document is drawn. Kept
// per page as pages are laid out; entries replaced by an edit keep their pages until those are laid out again.
export class PageIndex {
	private readonly story: StoryEdits
	private readonly region: Region
	private readonly byEntry = new Map<TopEntry, number[]>()
	private readonly byPage: (readonly TopEntry[] | undefined)[] = []

	constructor(story: StoryEdits, region: Region = Region.BODY) {
		this.story = story
		this.region = region
	}

	// Records what page `page` (now at its index) shows, replacing what was recorded for that index.
	index(page: PageDisplayList): void {
		this.forget(page.index)

		const story = this.story.story.story.index
		const seen: TopEntry[] = []
		let last: TopEntry | undefined

		for (let source = 0; source < page.sourceCount; source++) {
			const at = source * SOURCE_INTS

			if (page.sourceInts[at + SourceInt.REGION] !== this.region || page.sourceInts[at + SourceInt.STORY] !== story) continue

			const entry = this.story.entryOf(page.sourceParts[source] ?? -1, page.sourceInts[at + SourceInt.ITEM] ?? -1)

			if (entry === undefined || entry === last || seen.includes(entry)) continue

			last = entry
			seen.push(entry)

			const pages = this.byEntry.get(entry)

			if (pages === undefined) this.byEntry.set(entry, [page.index])
			else insertSorted(pages, page.index)
		}

		this.byPage[page.index] = seen
	}

	// Forgets pages from `count` on (the layout has fewer pages now).
	truncate(count: number): void {
		for (let index = count; index < this.byPage.length; index++) this.forget(index)

		this.byPage.length = Math.min(this.byPage.length, count)
	}

	private forget(index: number): void {
		const entries = this.byPage[index]

		if (entries === undefined) return

		for (const entry of entries) {
			const pages = this.byEntry.get(entry)

			if (pages === undefined) continue

			const at = pages.indexOf(index)

			if (at >= 0) pages.splice(at, 1)
			if (pages.length === 0) this.byEntry.delete(entry)
		}

		this.byPage[index] = undefined
	}

	pagesOf(entry: TopEntry): readonly number[] {
		return this.byEntry.get(entry) ?? []
	}

	entriesOn(page: number): readonly TopEntry[] {
		return this.byPage[page] ?? []
	}

	// The first page showing any of `entries`, or -1.
	firstPage(entries: Iterable<TopEntry>): number {
		let first = -1

		for (const entry of entries) {
			const page = this.byEntry.get(entry)?.[0]

			if (page !== undefined && (first < 0 || page < first)) first = page
		}

		return first
	}

	// The page layout restarts at for edits to `entries`, or -1 when no page shows them: the first page showing
	// one, or the page before when one opens it (it may now fit, in part, at the end of that page).
	restartPage(entries: readonly TopEntry[]): number {
		const first = this.firstPage(entries)
		const opening = first > 0 ? this.byPage[first]?.[0] : undefined

		return opening !== undefined && entries.includes(opening) ? first - 1 : first
	}
}

function insertSorted(list: number[], value: number): void {
	let at = list.length

	while (at > 0 && (list[at - 1] ?? 0) > value) at--

	if (list[at - 1] !== value) list.splice(at, 0, value)
}
