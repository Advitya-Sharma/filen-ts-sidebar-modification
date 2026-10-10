// A growable Int32 column for model tables: rows are appended while one part is built, then trimmed.
export class Column {
	private array: Int32Array
	length = 0

	constructor(capacity: number = 16) {
		this.array = new Int32Array(Math.max(capacity, 4))
	}

	push(value: number): number {
		if (this.length === this.array.length) {
			const grown = new Int32Array(this.array.length * 2)

			grown.set(this.array)
			this.array = grown
		}

		this.array[this.length] = value

		return this.length++
	}

	get(index: number): number {
		return this.array[index] ?? -1
	}

	set(index: number, value: number): void {
		this.array[index] = value
	}

	trim(): Int32Array {
		return this.array.length === this.length ? this.array : this.array.slice(0, this.length)
	}
}

// Interned strings (style ids, authors, dates): one copy per distinct value for the whole model.
export class StringTable {
	readonly values: string[] = []
	private readonly index = new Map<string, number>()

	intern(value: string | undefined): number {
		if (value === undefined) return -1

		let id = this.index.get(value)

		if (id === undefined) {
			id = this.values.length
			this.values.push(value)
			this.index.set(value, id)
		}

		return id
	}

	get(id: number): string | undefined {
		return id < 0 ? undefined : this.values[id]
	}
}
