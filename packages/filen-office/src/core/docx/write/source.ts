import type { Mark, MarkType, Node } from "prosemirror-model"
import { NO_SOURCE } from "../edit/schema"
import { Item, type StoryPart } from "../model/index"
import { DocxWriteError } from "./values"

export const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
export const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
export const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"

// The source item a node or mark stands for, or NO_SOURCE.
export function src(node: Node | Mark): number {
	const value: unknown = node.attrs["src"]

	return typeof value === "number" ? value : NO_SOURCE
}

// Whether source item `item` lies inside a table cell.
export function inCell(part: StoryPart, item: number): boolean {
	for (let at = part.parent(item); at >= 0; at = part.parent(at)) if (part.kinds[at] === Item.Cell) return true

	return false
}

// Whether source item `item` lies in a table cell inside source item `container`.
export function cellWithin(part: StoryPart, item: number, container: number): boolean {
	for (let at = part.parent(item); at >= 0 && at !== container; at = part.parent(at)) if (part.kinds[at] === Item.Cell) return true

	return false
}

// Whether a paragraph ending with the section break of source paragraph `sectPr` writes it, in a table cell
// or not: a break written in a cell is no section break (the model ignores it), so one is written only
// where its source paragraph had it, in a cell or out of one.
export function sectionWritten(part: StoryPart, sectPr: number, cell: boolean): boolean {
	return sectPr !== NO_SOURCE && inCell(part, sectPr) === cell
}

// What a feature records on a block as a node mark reaches the file through the block's writer, or the save
// is refused: `written` is the one mark that writer writes.
export function refuseBlockMarks(node: Node, written?: MarkType): void {
	for (const mark of node.marks) {
		if (mark.type !== written) throw new DocxWriteError("missingSource", `A ${mark.type.name} mark on a ${node.type.name}`)
	}
}
