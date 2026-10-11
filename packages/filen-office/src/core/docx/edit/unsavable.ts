import type { DocxEditorSession } from "./session"

// Decision 5 at save time: when the writer refuses an edited document (its own checks failed), the edited
// blocks it cannot write are found, and the save writes them as they were last saved (DocxEditorSession
// .withReverted; .revert once the file is stored). `writes` tries the writer on stand-in sessions without
// storing anything.

export interface UnsavableBlocks {
	readonly session: DocxEditorSession
	// Positions in session.state.doc (from editedBlocks).
	readonly positions: readonly number[]
}

export interface UnsavableSearchOptions {
	// Writer runs at most (each writes the whole document); default 24.
	readonly maxAttempts?: number
	// Time the search may take (ms, from its start); default 20 s.
	readonly budgetMs?: number
	readonly now?: () => number
	// Between runs: the page's turn (input, painting).
	readonly pause?: () => Promise<void>
	// Before each run, with the number of runs so far.
	readonly onAttempt?: (attempt: number) => void
	// The writer is known to refuse the sessions as they are (it just did, for this reason): not tried again.
	readonly refused?: boolean
}

interface Candidate {
	readonly session: number
	readonly pos: number
}

class SearchExhausted extends Error {}

// The edited blocks to take back so that the writer accepts the document: a small set, found by halving the
// edited blocks. Null when taking edited blocks back does not help (the refusal is not about one of them),
// or the search ran out of runs or time; the caller then keeps the document as it is.
export async function findUnsavableBlocks(
	sessions: readonly DocxEditorSession[],
	writes: (sessions: readonly DocxEditorSession[]) => Promise<boolean>,
	options: UnsavableSearchOptions = {}
): Promise<UnsavableBlocks[] | null> {
	const maxAttempts = options.maxAttempts ?? 24
	const budget = options.budgetMs ?? 20_000
	const now = options.now ?? (() => Date.now())
	const started = now()
	const editable = sessions.map(session => (session.readOnly ? [] : session.editedBlocks()))
	const all: Candidate[] = editable.flatMap((positions, session) => positions.map(pos => ({ session, pos })))
	let attempts = 0

	if (all.length === 0) return null

	const passes = async (reverted: readonly Candidate[]): Promise<boolean> => {
		if (attempts >= maxAttempts || (attempts > 0 && now() - started > budget)) throw new SearchExhausted()

		if (attempts > 0) await options.pause?.()

		options.onAttempt?.(attempts)
		attempts++

		return writes(
			sessions.map((session, index) => {
				const positions = reverted.filter(candidate => candidate.session === index).map(candidate => candidate.pos)

				return positions.length === 0 ? session : session.withReverted(positions)
			})
		)
	}

	// Blocks of `candidates` that must go back, given that taking back `base` and all of `candidates` passes and
	// `base` alone does not.
	const search = async (candidates: readonly Candidate[], base: readonly Candidate[]): Promise<Candidate[]> => {
		if (candidates.length === 1) return [...candidates]

		const half = candidates.length >> 1
		const a = candidates.slice(0, half)
		const b = candidates.slice(half)

		if (await passes([...base, ...a])) return search(a, base)
		if (await passes([...base, ...b])) return search(b, base)

		const fromA = await search(a, [...base, ...b])

		return [...fromA, ...(await search(b, [...base, ...fromA]))]
	}

	try {
		// Refused for something else (the zip, a part no edit made), or not even with every edited block back.
		if ((options.refused !== true && (await passes([]))) || !(await passes(all))) return null

		const found = await search(all, [])

		return sessions
			.map((session, index) => ({
				session,
				positions: found.filter(candidate => candidate.session === index).map(candidate => candidate.pos)
			}))
			.filter(entry => entry.positions.length > 0)
	} catch (error) {
		if (error instanceof SearchExhausted) return null

		throw error
	}
}
