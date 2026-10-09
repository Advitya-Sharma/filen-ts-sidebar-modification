// Scroll positions of the Videos views by history entry, so Back returns to where the list was. Session memory only,
// bounded to the most recent entries.

const MAX_ENTRIES = 200

const positions = new Map<string, number>()

export function rememberScroll(key: string, value: number): void {
	positions.delete(key)
	positions.set(key, value)

	if (positions.size > MAX_ENTRIES) {
		const oldest = positions.keys().next().value

		if (oldest !== undefined) {
			positions.delete(oldest)
		}
	}
}

export function recallScroll(key: string): number | undefined {
	return positions.get(key)
}
