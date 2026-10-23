// The name of a converted copy (no imports: worker modules use it too).

// "budget.xls" → "budget.xlsx", or "budget (2).xlsx", "(3)"… while `taken` says the name is in use.
export async function convertedCopyName(original: string, extension: string, taken: (name: string) => Promise<boolean>): Promise<string> {
	const dot = original.lastIndexOf(".")
	const base = dot > 0 ? original.slice(0, dot) : original

	for (let attempt = 1; ; attempt++) {
		const candidate = attempt === 1 ? `${base}.${extension}` : `${base} (${String(attempt)}).${extension}`

		if (!(await taken(candidate))) {
			return candidate
		}
	}
}
