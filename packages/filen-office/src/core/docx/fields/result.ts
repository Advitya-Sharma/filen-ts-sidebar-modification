import { FieldCharType, Item, type StoryPart } from "../model/index"
import { Ns } from "../model/namespaces"

// The cached result of a source field as the editor reads it: what follows the separator (a simple field's
// content), nested field codes left out. A result Word updated while tracking changes holds the old result in
// w:del and the new one in w:ins (each update recorded as a deletion and an insertion); read as the result
// after them and the field's tracked update: the result before them and who updated it. Read the same by the
// builder and the writer, so a result no command changed is told apart.

// A tracked update of a field's result: the result before it, who updated it and when, and the change's id
// (w:id of its deletion, else insertion).
export interface FieldUpdate {
	readonly result: string
	readonly author: string
	readonly date: string
	readonly id: string
}

export interface StoredResult {
	readonly result: string
	readonly update: FieldUpdate | null
}

const NOT_DELETED = 1
const NOT_INSERTED = 2

export function storedResult(part: StoryPart, first: number, last: number, simple: boolean): StoredResult {
	const field = simple ? (part.aux[first] ?? -1) : part.fieldOf(first + 1)
	const separate = simple || field < 0 ? first : (part.fields.separate[field] ?? -1)

	if (separate < 0) return { result: "", update: null }

	// What holds the field's runs: tracked changes outside it (a field inserted or deleted whole) are not its.
	const container = simple ? first : part.parent(first)
	const end = part.end(last)
	const codes: boolean[] = []
	let after = ""
	let before = ""
	let revision = -1
	let depth = 0
	let code = false
	// Whether the text of `item` shows after the change (NOT_DELETED) and before it (NOT_INSERTED).
	const shows = (item: number): number => {
		let out = NOT_DELETED | NOT_INSERTED

		for (let at = part.parent(item); at > container; at = part.parent(at)) {
			const kind = part.kinds[at]

			if (kind === Item.Deletion || kind === Item.MoveFrom) out &= ~NOT_DELETED
			else if (kind === Item.Insertion || kind === Item.MoveTo) out &= ~NOT_INSERTED
			else continue

			if (revision < 0 || (part.kinds[revision] !== Item.Deletion && kind === Item.Deletion)) revision = at
		}

		return out
	}
	const add = (item: number, text: string): void => {
		const visible = shows(item)

		if ((visible & NOT_DELETED) !== 0) after += text
		if ((visible & NOT_INSERTED) !== 0) before += text
	}

	for (let item = separate + 1; item < end; item++) {
		switch (part.kinds[item]) {
			case Item.FieldChar: {
				const type = part.fieldCharType(item)

				if (type === FieldCharType.Begin) {
					codes.push(code)
					depth++
					code = true
				} else if (type === FieldCharType.Separate && depth > 0) {
					code = codes[codes.length - 1] ?? false
				} else if (type === FieldCharType.End && depth > 0) {
					depth--
					code = codes.pop() ?? false
				}

				break
			}
			case Item.Text:
			case Item.DeletedText:
				if (!code) add(item, part.text(item))

				break
			case Item.Tab:
				if (!code) add(item, "\t")

				break
			case Item.Break:
			case Item.CarriageReturn:
				if (!code) add(item, "\n")

				break
		}
	}

	if (revision < 0) return { result: after, update: null }

	const node = part.node(revision)
	const reader = part.reader

	return {
		result: after,
		update: {
			result: before,
			author: reader.attr(node, Ns.W, "author") ?? "",
			date: reader.attr(node, Ns.W, "date") ?? "",
			id: reader.attr(node, Ns.W, "id") ?? ""
		}
	}
}
