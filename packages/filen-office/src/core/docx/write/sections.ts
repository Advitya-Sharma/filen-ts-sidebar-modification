import type { Node } from "prosemirror-model"
import { partKey } from "../../opc/index"
import { finalSectionEdit, sectionMarkup, type SectionEdit } from "../stories/sections"
import type { StoryWriting } from "./story"
import { DocxWriteError } from "./values"

// Section properties as edited (see stories/sections.ts): the w:sectPr a section edit makes, with the main
// part's relationship to each header and footer it refers to (a new story's part gets one from the save).

const HEADER_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
const FOOTER_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
const STRICT_BASE = "http://purl.oclc.org/ooxml/officeDocument/relationships/"

// The relationship id the body's part has (or the save gives) to header or footer story `story`.
function storyRelationship(w: StoryWriting, story: number): string {
	const context = w.env.context

	for (const [added, made] of context.newStories) {
		if (added.index !== story) continue

		const id = w.env.parts?.storyRelationship(made.target)

		// Gate 1 and the layout of edited blocks write without a save: the reference is what counts there.
		if (id !== undefined) return id
		if (w.env.parts === undefined) return `filenStory${String(story)}`

		throw new DocxWriteError("unsupported", `No relationship to the new ${made.kind} ${made.target}`)
	}

	const source = context.model.stories[story]
	const part = source === undefined ? undefined : context.model.parts[source.part]

	if (source === undefined || part === undefined || (source.kind !== "header" && source.kind !== "footer"))
		throw new DocxWriteError("missingSource", `A section refers to story ${String(story)}, no header or footer`)

	const key = partKey(part.name)
	const found = w.part.relationships.all.find(
		relationship =>
			!relationship.external &&
			relationship.partName !== undefined &&
			partKey(relationship.partName) === key &&
			(relationship.type === HEADER_TYPE ||
				relationship.type === FOOTER_TYPE ||
				relationship.type === `${STRICT_BASE}header` ||
				relationship.type === `${STRICT_BASE}footer`)
	)

	if (found === undefined) throw new DocxWriteError("missingSource", `No relationship to ${part.name}`)

	return found.id
}

// The w:sectPr a paragraph's (or the body's own) section edit writes.
export function sectionEditMarkup(w: StoryWriting, edit: SectionEdit): string {
	const context = w.env.context

	return sectionMarkup(
		{
			part: w.part,
			names: w.names,
			opened: context.model.sections,
			relationship: story => storyRelationship(w, story),
			props: node => context.store.section(w.part.document, node),
			revisionId: () => w.env.nextRevisionId()
		},
		edit
	)
}

// The body's own w:sectPr where an edit changed it (made where the body had none), else undefined.
export function finalSectionMarkup(w: StoryWriting, doc: Node): string | undefined {
	const edit = finalSectionEdit(doc)

	return edit === undefined ? undefined : sectionEditMarkup(w, edit)
}
