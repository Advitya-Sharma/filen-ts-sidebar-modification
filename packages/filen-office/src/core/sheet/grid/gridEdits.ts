export interface CellEntry {
	row: number
	col: number
	input: string
}

// The name an added sheet proposes: numbered after the sheets there are, or the first number past it no
// sheet's name takes (sheet names are unique whatever their case).
export function newSheetName(names: readonly string[], nameFor: (number: number) => string): string {
	const taken = new Set(names.map(name => name.toLowerCase()))
	let number = names.length + 1

	while (taken.has(nameFor(number).toLowerCase())) {
		number++
	}

	return nameFor(number)
}
