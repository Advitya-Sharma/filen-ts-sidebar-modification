import type { ListLabel } from "./lists"
import type { DocModel, DocStoryKind } from "./model"
import type { DocPara } from "./props"

// Plain text of a story, as LibreOffice's text export writes it: field results (never instructions), list labels
// followed by a space, note reference numbers, one line per paragraph and per table cell, symbol characters as
// their code point, pictures and other anchors left out.

export interface DocTextOptions {
	// Default "main".
	readonly story?: DocStoryKind
	// Include text formatted as hidden (vanish). Default false.
	readonly hidden?: boolean
	// Include deleted revision text. Default true, as LibreOffice's text export does.
	readonly deleted?: boolean
	// Include list labels and note reference numbers (generated text). Default true.
	readonly labels?: boolean
	// Include positioned paragraphs (Word frames) where they are in the text. Default false, as LibreOffice's text
	// export leaves them out.
	readonly framed?: boolean
}

// What characters with a meaning of their own (§2.3.1 ff.) contribute to plain text; undefined = the character.
function plainChar(code: number): string | undefined {
	switch (code) {
		case 0x0b:
		case 0x0c:
		case 0x0e:
			return "\n"
		case 0x09:
			return "\t"
		case 0x1e:
			return "\u2011"
		case 0x1f:
			return "\u00ad"
		default:
			return code < 0x20 ? "" : undefined
	}
}

// Plain text has no font for a bullet: levels take the bullets LibreOffice writes (•, ◦, ▪ in turn).
const LEVEL_BULLETS = ["•", "◦", "▪"]

function listLabel(label: ListLabel, ilvl: number): string {
	if (label.bullet !== undefined) return LEVEL_BULLETS[ilvl % 3] ?? "•"

	return label.text
}

// Positioned paragraphs (Word frames) are not body text; LibreOffice exports them as frames, outside the text.
function framed(props: DocPara): boolean {
	const frame = props.frame

	return frame !== undefined && (frame.x !== 0 || frame.y !== 0 || frame.width !== 0 || frame.height !== 0)
}

const FIELD_BEGIN = 0x13
const FIELD_SEPARATOR = 0x14
const FIELD_END = 0x15
const MAX_FIELD_DEPTH = 64

export function docText(model: DocModel, options: DocTextOptions = {}): string {
	const kind = options.story ?? "main"
	const story = model.stories.find(entry => entry.kind === kind)

	if (story === undefined || story.end <= story.start) return ""

	const { text, runs, paragraphs } = model
	const notes = new Map<number, string>()

	for (const note of [...model.footnotes, ...model.endnotes]) if (note.label !== undefined) notes.set(note.reference, note.label)

	// Field marks from the PLCs, plus fSpec field characters a PLC misses.
	const fieldMarks = new Set<number>()

	for (const field of model.fields) {
		fieldMarks.add(field.begin).add(field.end)

		if (field.separator !== undefined) fieldMarks.add(field.separator)
	}

	const parts: string[] = []
	// Per open field: true while in its instruction part.
	const fields: boolean[] = []
	let inInstruction = 0
	let paragraph = Math.max(
		0,
		paragraphs.findIndex(entry => entry.end > story.start)
	)
	let labelled = -1

	for (const run of runs) {
		if (run.end <= story.start) continue
		if (run.start >= story.end) break

		const hidden = (run.props.vanish && options.hidden !== true) || (run.props.deleted !== undefined && options.deleted === false)

		for (let cp = Math.max(run.start, story.start); cp < Math.min(run.end, story.end); cp++) {
			while (paragraph < paragraphs.length && (paragraphs[paragraph]?.end ?? 0) <= cp) paragraph++

			const owner = paragraphs[paragraph]
			const code = text.charCodeAt(cp)

			if ((code === FIELD_BEGIN || code === FIELD_SEPARATOR || code === FIELD_END) && (run.props.special || fieldMarks.has(cp))) {
				if (code === FIELD_BEGIN && fields.length < MAX_FIELD_DEPTH) {
					fields.push(true)
					inInstruction++
				} else if (code === FIELD_SEPARATOR && fields.at(-1) === true) {
					fields[fields.length - 1] = false
					inInstruction--
				} else if (code === FIELD_END && fields.length > 0) {
					if (fields.pop() === true) inInstruction--
				}

				continue
			}

			const outside = options.framed !== true && owner !== undefined && owner.depth === 0 && framed(owner.props)

			if (owner !== undefined && labelled !== paragraph && cp >= owner.start && inInstruction === 0 && !outside) {
				labelled = paragraph

				if (options.labels !== false && owner.list !== undefined && owner.list.text !== "")
					parts.push(listLabel(owner.list, owner.props.ilvl), " ")
			}

			if (owner !== undefined && cp === owner.end - 1 && owner.ending !== "none") {
				// A row mark follows its last cell's mark, which already ended the line.
				if (owner.ending !== "row" && !outside) parts.push("\n")

				continue
			}

			if (inInstruction > 0 || hidden || outside) continue

			if (run.props.special) {
				const note = code === 0x02 && options.labels !== false ? notes.get(cp) : undefined

				if (note !== undefined) parts.push(note)

				continue
			}

			if (run.props.symbol !== undefined) {
				parts.push(String.fromCharCode(run.props.symbol.char))
				continue
			}

			parts.push(plainChar(code) ?? text.charAt(cp))
		}
	}

	return parts.join("")
}
