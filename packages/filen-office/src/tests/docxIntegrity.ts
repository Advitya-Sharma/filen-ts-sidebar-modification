import { NodeKind } from "@filen/office/core/xml"
import { Item, type DocxModel, type StoryPart } from "@filen/office/core/docx"

// Structural checks over a built model: every item maps to an element of its part, items nest as their
// elements do and appear in document order, siblings never overlap, and every element child of a
// container is accounted for (an item, a property element its container recorded, or proofErr).
// Returns human-readable problems (empty when sound).
export function modelProblems(model: DocxModel, limit: number = 20): string[] {
	const problems: string[] = []

	for (const part of model.parts) {
		partProblems(part, problems, limit)

		if (problems.length >= limit) break
	}

	sectionProblems(model, problems)

	return problems.slice(0, limit)
}

// Sections tile the body without gaps, and each lists exactly the wrappers open at its start.
function sectionProblems(model: DocxModel, problems: string[]): void {
	const story = model.body

	if (story === undefined) return

	const part = model.partOf(story)
	let expected = story.item + 1

	for (const section of model.sections) {
		if (section.start !== expected) problems.push(`section ${section.index} starts at ${section.start}, not ${expected}`)

		const open: number[] = []

		if (section.start < section.end) {
			for (let ancestor = part.parent(section.start); ancestor > story.item; ancestor = part.parent(ancestor)) open.unshift(ancestor)
		}

		if (open.join() !== section.openWrappers.join()) {
			problems.push(`section ${section.index} lists open wrappers [${section.openWrappers.join()}], not [${open.join()}]`)
		}

		expected = section.end
	}

	if (model.sections.length > 0 && expected !== part.end(story.item)) problems.push(`sections end at ${expected}, not the body's end`)
}

const NO_COVERAGE = new Set<number>([Item.Drawing, Item.VmlPicture, Item.EmbeddedObject, Item.Ruby])

function partProblems(part: StoryPart, problems: string[], limit: number): void {
	const document = part.document
	const report = (item: number, message: string): void => {
		if (problems.length < limit) problems.push(`${part.name} item ${item} (${document.name(part.node(item))}): ${message}`)
	}
	const itemOfNode = new Map<number, number>()
	let previousStart = -1

	for (let item = 0; item < part.count; item++) {
		const node = part.node(item)
		const parent = part.parent(item)
		const end = part.end(item)

		if (node < 0 || node >= document.count || document.kinds[node] !== NodeKind.Element) {
			report(item, "does not map to an element")

			continue
		}

		if (itemOfNode.has(node)) report(item, "shares its element with another item")

		itemOfNode.set(node, item)

		if (document.start(node) <= previousStart) report(item, "is out of document order")

		previousStart = document.start(node)

		if (end <= item || end > part.count) report(item, `has an invalid end ${end}`)

		if (parent >= item) report(item, "has a parent after it")

		if (parent >= 0) {
			const parentNode = part.node(parent)

			if (end > part.end(parent)) report(item, "ends after its parent")
			if (document.start(node) < document.openEnd(parentNode) || document.end(node) > document.end(parentNode)) {
				report(item, "lies outside its parent's element")
			}
		}

		const next = part.nextSibling(item)

		if (next >= 0 && document.start(part.node(next)) < document.end(node)) report(item, "overlaps its next sibling")
	}

	for (let item = 0; item < part.count; item++) {
		const kind = part.kind(item)

		if (part.end(item) === item + 1 || NO_COVERAGE.has(kind)) continue

		const owned = ownedNodes(part, item)
		const containers = contentNodes(part, item)

		for (const container of containers) {
			for (let child = document.firstChildElement(container); child >= 0; child = document.nextSiblingElement(child)) {
				if (owned.has(child) || containers.includes(child) || document.localName(child) === "proofErr") continue

				const childItem = itemOfNode.get(child)

				if (childItem === undefined || part.parent(childItem) !== item) {
					report(item, `child <${document.name(child)}> is not accounted for`)
				}
			}
		}
	}
}

function contentNodes(part: StoryPart, item: number): number[] {
	const node = part.node(item)

	switch (part.kind(item)) {
		case Item.Sdt: {
			const row = part.aux[item] ?? -1
			const content = part.sdts.content[row] ?? -1

			return content < 0 ? [node] : [node, content]
		}
		case Item.AlternateContent: {
			const selected = part.alternates.selected[part.aux[item] ?? -1] ?? -1

			return selected < 0 ? [] : [selected]
		}
		default:
			return [node]
	}
}

function ownedNodes(part: StoryPart, item: number): Set<number> {
	const owned = new Set<number>()
	const row = part.aux[item] ?? -1
	const add = (value: number | undefined): void => {
		if (value !== undefined && value >= 0) owned.add(value)
	}

	switch (part.kind(item)) {
		case Item.Paragraph:
			add(part.paragraphs.pPr[row])

			break
		case Item.Run:
			add(part.runs.rPr[row])

			break
		case Item.Table:
			add(part.tables.tblPr[row])
			add(part.tables.tblGrid[row])

			break
		case Item.Row:
			add(part.rows.trPr[row])
			add(part.rows.tblPrEx[row])

			break
		case Item.Cell:
			add(part.cells.tcPr[row])

			break
		case Item.Sdt:
			add(part.sdts.sdtPr[row])
			add(part.sdts.sdtEndPr[row])

			break
		case Item.SmartTag:
		case Item.CustomXml:
			add(row)

			break
	}

	return owned
}
