import type { Node } from "prosemirror-model"
import type { EditorState, Transaction } from "prosemirror-state"
import type { Mapping } from "prosemirror-transform"
import type { DocxEditorSession } from "../edit/session"
import { docxSchema } from "../edit/schema"
import type { SpellingChecker, SpellingIssue, SpellingMark, SpellingSuggestOptions } from "./checker"
import { ParagraphScanner, inWordStyle, spellingQuery, type RelativeIssue } from "./scan"

// Spelling of one editing session, headless: the issues of the document (misspelled and repeated words)
// as ranges for a view to mark, kept in step with every edit, and the actions on them (suggestions,
// replace, ignore all, add to dictionary). Edits map the existing marks at once; the edited paragraphs
// are checked again after a short pause, and each paragraph's result is kept until its node changes, so a
// pass over an edited document only scans what changed. The word being typed is not marked until the
// caret leaves it. Nothing here touches the DOM.

export interface DocxSpellingOptions {
	// Language for text whose runs and styles name none (the app's editing language).
	readonly defaultLanguage?: string
	// Word's proofing options, on by default.
	readonly ignoreUppercase?: boolean
	readonly ignoreWordsWithNumbers?: boolean
	readonly ignoreAddresses?: boolean
	readonly flagRepeated?: boolean
	// Pause after an edit before the document is checked again (ms).
	readonly delayMs?: number
	// Time one pass may spend scanning paragraphs before it yields (ms).
	readonly budgetMs?: number
	readonly now?: () => number
	readonly timers?: SpellingTimers
}

export interface SpellingTimers {
	set(callback: () => void, ms: number): unknown
	clear(handle: unknown): void
}

const DELAY_MS = 300
const BUDGET_MS = 8
// Marks kept per document (in document order). A text in the wrong language would otherwise be marked
// word by word, and every keystroke maps and redraws every mark.
export const MAX_SPELLING_ISSUES = 10_000
const NO_MARKS: readonly SpellingMark[] = Object.freeze([])

// An issue as kept here: its place moves with edits.
type MovableIssue = Omit<SpellingIssue, "from" | "to"> & { from: number; to: number }

function defaultNow(): () => number {
	const performance: unknown = Reflect.get(globalThis, "performance")
	const now: unknown = typeof performance === "object" && performance !== null ? Reflect.get(performance, "now") : undefined

	return typeof now === "function" ? () => (now as () => number).call(performance) : Date.now
}

function defaultTimers(): SpellingTimers {
	const set: unknown = Reflect.get(globalThis, "setTimeout")
	const clear: unknown = Reflect.get(globalThis, "clearTimeout")

	if (typeof set !== "function" || typeof clear !== "function") throw new Error("No timers: pass DocxSpellingOptions.timers")

	return {
		set: (callback, ms) => (set as (callback: () => void, ms: number) => unknown)(callback, ms),
		clear: handle => (clear as (handle: unknown) => void)(handle)
	}
}

// The range [from, to] of the document `mapping` produced that its steps changed, or undefined.
function changedRange(mapping: Mapping): { from: number; to: number } | undefined {
	let from = Number.POSITIVE_INFINITY
	let to = Number.NEGATIVE_INFINITY
	const maps = mapping.maps

	maps.forEach((map, index) => {
		map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
			let start = newStart
			let end = newEnd

			for (let i = index + 1; i < maps.length; i++) {
				start = maps[i]?.map(start, -1) ?? start
				end = maps[i]?.map(end, 1) ?? end
			}

			from = Math.min(from, start)
			to = Math.max(to, end)
		})
	})

	return from <= to ? { from, to } : undefined
}

export class DocxSpelling {
	// The document asks for spelling errors not to be shown (w:hideSpellingErrors): nothing is checked.
	readonly hidden: boolean
	private readonly session: DocxEditorSession
	private readonly checker: SpellingChecker
	private readonly scanner: ParagraphScanner
	private readonly delay: number
	private readonly budget: number
	private readonly now: () => number
	private readonly timers: SpellingTimers
	private readonly listeners = new Set<(marks: readonly SpellingMark[]) => void>()
	private readonly unsubscribe: (() => void)[] = []
	// Complete results by paragraph node (nodes are immutable: a changed paragraph is a new node).
	private results = new WeakMap<Node, readonly RelativeIssue[]>()
	// Results of paragraphs with words still unchecked: shown as they are and not scanned again until a check
	// answers, an edit or a notice from the checker.
	private pending = new WeakMap<Node, readonly RelativeIssue[]>()
	// Words asked for since the last edit or notice, by dictionary: never asked twice, so a checker that
	// answers without a verdict cannot start a loop.
	private requested = new Map<string, Set<string>>()
	// Dictionaries whose check failed (not loadable): not asked again until the next edit or notice.
	private readonly failed = new Set<string>()
	// In document order. Moved in place by edits (an issue keeps its identity while it stays), so a keystroke
	// allocates no issue; marks() is a new array whenever they changed.
	private issues: MovableIssue[] = []
	private readonly ignored = new Set<string>()
	private readonly ignoredRepeats = new Set<string>()
	// The last edit and the issue it is typing (not marked while the caret is in it).
	private lastEdit: { from: number; to: number } | undefined = undefined
	private suppressed: SpellingIssue | undefined = undefined
	private marksValue: readonly SpellingMark[] = NO_MARKS
	private timer: unknown = undefined
	private timerDue = Number.POSITIVE_INFINITY
	private disposed = false

	constructor(session: DocxEditorSession, checker: SpellingChecker, options: DocxSpellingOptions = {}) {
		this.session = session
		this.checker = checker
		this.hidden = session.context.model.settings.hideSpellingErrors
		this.delay = options.delayMs ?? DELAY_MS
		this.budget = options.budgetMs ?? BUDGET_MS
		this.now = options.now ?? defaultNow()
		this.timers = options.timers ?? defaultTimers()
		this.scanner = new ParagraphScanner(session.context, session.story.part, checker, {
			defaultLanguage: options.defaultLanguage,
			ignoreUppercase: options.ignoreUppercase ?? true,
			ignoreWordsWithNumbers: options.ignoreWordsWithNumbers ?? true,
			ignoreAddresses: options.ignoreAddresses ?? true,
			flagRepeated: options.flagRepeated ?? true
		})

		if (this.hidden) return

		this.unsubscribe.push(
			session.onApplied((transactions, state) => this.onApplied(transactions, state)),
			checker.subscribe(() => this.invalidate())
		)
		this.schedule(0)
	}

	// Ranges to mark now, in document order (empty while the document hides spelling errors).
	marks(): readonly SpellingMark[] {
		return this.marksValue
	}

	// Hears every change of marks() (synchronously after the edit that moved them).
	subscribe(listener: (marks: readonly SpellingMark[]) => void): () => void {
		this.listeners.add(listener)

		return () => {
			this.listeners.delete(listener)
		}
	}

	dispose(): void {
		this.disposed = true
		this.cancel()

		for (const stop of this.unsubscribe) stop()

		this.unsubscribe.length = 0
		this.listeners.clear()
	}

	// The issue at a document position (inside the word or at either end), the one being typed included.
	spellingAt(pos: number): SpellingIssue | undefined {
		const issues = this.issues
		let low = 0
		let high = issues.length

		while (low < high) {
			const middle = (low + high) >> 1

			if ((issues[middle]?.to ?? 0) < pos) low = middle + 1
			else high = middle
		}

		const issue = issues[low]

		return issue !== undefined && issue.from <= pos && pos <= issue.to && this.current(issue) ? issue : undefined
	}

	// Corrections for the misspelled word at `pos`, best first: the fast pass, and the full one when that
	// finds nothing or `more` asks for it. A repeated word has none (its fix is deleting it).
	async suggestions(pos: number, options: SpellingSuggestOptions = {}): Promise<readonly string[]> {
		const issue = this.spellingAt(pos)

		if (issue === undefined || issue.kind !== "spelling") return []

		const query = spellingQuery(issue.word)
		const fast = options.more === true ? [] : await this.checker.suggest(issue.dictionary, query, { more: false })

		return inWordStyle(issue.word, fast.length > 0 ? fast : await this.checker.suggest(issue.dictionary, query, { more: true }))
	}

	// Replaces the word at `pos` (keeping its formatting); "" for a repeated word deletes it with the space
	// before it. False when there is no issue there or the session refused the edit (read-only).
	replaceWord(pos: number, replacement: string): boolean {
		const issue = this.spellingAt(pos)

		if (issue === undefined) return false

		const state = this.session.state
		let tr: Transaction

		if (replacement === "") {
			const $from = state.doc.resolve(issue.from)
			const before = $from.parent.textBetween(Math.max(0, $from.parentOffset - 32), $from.parentOffset, undefined, "￼")
			const spaces = before.length - before.trimEnd().length

			tr = state.tr.delete(issue.from - spaces, issue.to)
		} else {
			tr = state.tr.insertText(replacement, issue.from, issue.to)
		}

		return this.session.dispatch(tr.scrollIntoView())
	}

	// Stops marking this word (a repeated word: stops flagging its repetitions) for the rest of the session.
	ignoreAll(pos: number): boolean {
		const issue = this.spellingAt(pos)

		if (issue === undefined) return false

		if (issue.kind === "repeated") this.ignoredRepeats.add(issue.word.toLowerCase())
		else this.ignored.add(spellingQuery(issue.word))

		this.invalidate()

		return true
	}

	// Adds the word at `pos` to the user's dictionary; the checker's change notice re-checks the document.
	async addToDictionary(pos: number): Promise<boolean> {
		const issue = this.spellingAt(pos)

		if (issue?.kind !== "spelling") return false

		await this.checker.addToDictionary(spellingQuery(issue.word))

		return true
	}

	// ---- internals --------------------------------------------------------------------------------------

	// Whether a mapped issue still covers its word (an edit since the last pass may have changed it).
	private current(issue: SpellingIssue): boolean {
		const doc = this.session.state.doc

		if (issue.to > doc.content.size) return false

		return doc.textBetween(issue.from, issue.to, undefined, "￼").replace(/[­￼]/g, "") === issue.word
	}

	private onApplied(transactions: readonly Transaction[], state: EditorState): void {
		let changed = false

		for (const tr of transactions) {
			if (!tr.docChanged) continue

			changed = true
			this.mapIssues(this.issues, tr.mapping)
			this.lastEdit = changedRange(tr.mapping) ?? this.lastEdit
		}

		if (changed) {
			this.retry()
			this.schedule(this.delay)
		}

		this.emit(state, changed)
	}

	// Moves the issues through an edit, in place; those it deleted go.
	private mapIssues(issues: MovableIssue[], mapping: Mapping): void {
		let kept = 0
		const count = issues.length

		// Indexed: the array is compacted as it is read.
		for (let k = 0; k < count; k++) {
			const issue = issues[k]

			if (issue === undefined) continue

			const from = mapping.map(issue.from, 1)
			const to = mapping.map(issue.to, -1)

			if (to <= from) continue

			issue.from = from
			issue.to = to
			issues[kept++] = issue
		}

		issues.length = kept
	}

	private invalidate(): void {
		this.results = new WeakMap()
		this.retry()
		this.schedule(0)
	}

	// Unchecked words and failed dictionaries get asked for again.
	private retry(): void {
		this.pending = new WeakMap()
		this.requested = new Map()
		this.failed.clear()
	}

	private schedule(ms: number): void {
		if (this.disposed || this.hidden) return

		const due = this.now() + ms

		// An earlier pass already due covers this request, unless this one asks for a later one (an edit).
		if (this.timer !== undefined && due <= this.timerDue && ms === 0) return

		this.cancel()
		this.timerDue = due
		this.timer = this.timers.set(() => {
			this.timer = undefined
			this.timerDue = Number.POSITIVE_INFINITY
			this.pass()
		}, ms)
	}

	private cancel(): void {
		if (this.timer !== undefined) this.timers.clear(this.timer)

		this.timer = undefined
		this.timerDue = Number.POSITIVE_INFINITY
	}

	// One pass over the document: cached and pending paragraphs cost a lookup, new ones are scanned within the
	// time budget (the rest keep their mapped issues and the next pass continues), words never asked for go to
	// the checker, whose answer starts another pass.
	private pass(): void {
		if (this.disposed) return

		const state = this.session.state
		const deadline = this.now() + this.budget
		const previous = this.issues
		const next: MovableIssue[] = []
		// The previous issues are in document order: one equal to a new one (same place, word and verdict) is
		// kept, so a pass over a document that barely changed allocates little.
		let reused = 0
		const issueAt = (issue: RelativeIssue, start: number): MovableIssue => {
			const from = issue.from + start
			const to = issue.to + start

			while (reused < previous.length && (previous[reused]?.from ?? 0) < from) reused++

			const old = previous[reused]

			if (
				old !== undefined &&
				old.from === from &&
				old.to === to &&
				old.kind === issue.kind &&
				old.word === issue.word &&
				old.dictionary === issue.dictionary &&
				old.language === issue.language
			) {
				reused++

				return old
			}

			return { ...issue, from, to }
		}
		const unknown = new Map<string, Set<string>>()
		let carried = 0
		let scanned = 0
		let more = false

		const carry = (from: number, to: number): void => {
			while (carried < previous.length && (previous[carried]?.from ?? 0) < from) carried++

			while (carried < previous.length && (previous[carried]?.to ?? 0) <= to) {
				const issue = previous[carried++]

				if (issue !== undefined && next.length < MAX_SPELLING_ISSUES) next.push(issue)
			}
		}

		state.doc.descendants((node, pos) => {
			if (node.type === docxSchema.nodes.paragraph) {
				let issues = this.results.get(node) ?? this.pending.get(node)

				if (issues === undefined) {
					// At least one paragraph per pass, so a pass always makes progress.
					if (scanned > 0 && this.now() > deadline) {
						more = true
						carry(pos, pos + node.nodeSize)

						return false
					}

					const scan = this.scanner.scan(node, this.ignored, this.ignoredRepeats, unknown)

					scanned++

					if (scan.complete) this.results.set(node, scan.issues)
					else this.pending.set(node, scan.issues)

					issues = scan.issues
				}

				const start = pos + 1

				for (const issue of issues) {
					if (next.length >= MAX_SPELLING_ISSUES) break

					next.push(issueAt(issue, start))
				}

				return false
			}

			return node.type !== docxSchema.nodes.locked_block
		})

		this.issues = next

		for (const [dictionary, words] of unknown) {
			if (this.failed.has(dictionary)) continue

			let asked = this.requested.get(dictionary)

			if (asked === undefined) {
				asked = new Set()
				this.requested.set(dictionary, asked)
			}

			const fresh: string[] = []

			for (const word of words) {
				if (!asked.has(word)) {
					asked.add(word)
					fresh.push(word)
				}
			}

			if (fresh.length === 0) continue

			const requested = this.requested

			this.checker.check(dictionary, fresh).then(
				() => {
					// Stale: an edit or notice since asked again already.
					if (requested !== this.requested) return

					this.pending = new WeakMap()
					this.schedule(0)
				},
				() => {
					if (requested === this.requested) this.failed.add(dictionary)
				}
			)
		}

		if (more) this.schedule(0)

		this.emit(state, true)
	}

	// Recomputes the issue being typed and publishes marks when they changed.
	private emit(state: EditorState, issuesChanged: boolean): void {
		const selection = state.selection
		let suppressed: SpellingIssue | undefined

		if (selection.empty && this.lastEdit !== undefined) {
			const caret = selection.head
			const edit = this.lastEdit

			// An insertion inside or at the end of the word, or a deletion within it (an edit that only touches
			// the word's start, like typing before it, is not typing it).
			const typed = (issue: SpellingIssue): boolean =>
				edit.from === edit.to ? issue.from <= edit.from && edit.from <= issue.to : edit.from < issue.to && issue.from < edit.to

			suppressed = this.issues.find(issue => issue.from <= caret && caret <= issue.to && typed(issue))

			// The caret left the edited word: later visits to it are not typing.
			if (suppressed === undefined && (caret < edit.from - 1 || caret > edit.to + 1)) this.lastEdit = undefined
		}

		if (!issuesChanged && suppressed === this.suppressed) return

		this.suppressed = suppressed

		// A new array: the issues move in place.
		const at = suppressed === undefined ? -1 : this.issues.indexOf(suppressed)

		this.marksValue = at < 0 ? this.issues.slice() : this.issues.slice(0, at).concat(this.issues.slice(at + 1))

		for (const listener of [...this.listeners]) listener(this.marksValue)
	}
}
