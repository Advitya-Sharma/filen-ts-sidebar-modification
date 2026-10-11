import type { Mark, Node } from "prosemirror-model"
import type { StoryPart } from "../model/index"
import { EMPTY_PARAGRAPH, type ParagraphProps, type ResolvedRunProps, type RunProps } from "../props/index"
import type { ParagraphFormat } from "../styles/index"
import type { DocxEditContext } from "./context"
import { NO_SOURCE, docxSchema, type IndentAttrs } from "./schema"

export type ListKind = "bullet" | "numbered"

// A list a command started: its kind, and the existing w:num whose definition (w:abstractNum) the new
// w:num takes over, restarting the count; null for a new definition.
export interface NewList {
	readonly kind: ListKind
	readonly basedOn: number | null
}

export type RunToggle = "bold" | "italic" | "underline" | "strike"

export const RUN_TOGGLES: readonly RunToggle[] = ["bold", "italic", "underline", "strike"]
export const TOGGLE_ON = 1
export const TOGGLE_STYLE = 16

const marks = docxSchema.marks

// Effective formatting of editing nodes: a paragraph's direct properties are its source w:pPr with the
// attributes commands edit laid over; a run's toggles are its direct marks, else what the paragraph and
// character styles give. Everything is cached by identity (attribute objects are shared by unchanged node
// copies, resolver results are interned).
export class FormatResolver {
	private readonly context: DocxEditContext
	private readonly part: StoryPart
	private readonly paragraphs = new WeakMap<object, ParagraphFormat>()
	private readonly styleParagraphs = new Map<string | null, ParagraphFormat>()
	private readonly characterStyles = new Map<string | null, RunProps>()
	private readonly overrides = new WeakMap<ParagraphProps, Map<string, ParagraphProps>>()
	private readonly toggleCache = new WeakMap<readonly Mark[], { paragraph: ParagraphFormat; bits: number }>()
	// Lists created in this session (negative numIds) and their kind.
	readonly newLists = new Map<number, NewList>()

	constructor(context: DocxEditContext, part: StoryPart) {
		this.context = context
		this.part = part
	}

	private sourceProps(node: Node): ParagraphProps {
		const src = node.attrs["src"] as number

		if (src === NO_SOURCE || this.part.kinds[src] === undefined) return EMPTY_PARAGRAPH

		const pPr = this.part.paragraphs.pPr[this.part.aux[src] ?? -1] ?? -1

		return this.context.store.paragraph(this.part.document, pPr)
	}

	// The direct paragraph properties of a paragraph node: its source w:pPr's own (interned) object while
	// the attributes still say what it says, so the resolver's cache serves every unedited paragraph;
	// otherwise one interned object per source and override set.
	direct(node: Node): ParagraphProps {
		const source = this.sourceProps(node)
		const styleId = node.attrs["styleId"] as string | null
		const numId = node.attrs["numId"] as number | null
		const ilvl = node.attrs["ilvl"] as number | null
		const align = node.attrs["align"] as string | null
		const indent = node.attrs["indent"] as IndentAttrs | null

		if (
			styleId === (source.styleId ?? null) &&
			numId === (source.numbering?.numId ?? null) &&
			ilvl === (source.numbering?.ilvl ?? null) &&
			align === (source.alignment ?? null) &&
			indent === (source.indent ?? null)
		) {
			return source
		}

		const key = JSON.stringify([styleId, numId, ilvl, align, indent])
		let bySource = this.overrides.get(source)

		if (bySource === undefined) {
			bySource = new Map()
			this.overrides.set(source, bySource)
		}

		let direct = bySource.get(key)

		if (direct === undefined) {
			const { styleId: _styleId, numbering: _numbering, alignment: _alignment, indent: _indent, ...rest } = source

			direct = {
				...rest,
				...(styleId !== null ? { styleId } : {}),
				...(numId !== null
					? { numbering: ilvl === null ? { numId } : { numId, ilvl } }
					: ilvl !== null
						? { numbering: { ilvl } }
						: {}),
				...(align !== null ? { alignment: align } : {}),
				...(indent !== null ? { indent } : {})
			}
			bySource.set(key, direct)
		}

		return direct
	}

	paragraph(node: Node): ParagraphFormat {
		const key = node.attrs
		let format = this.paragraphs.get(key)

		if (format === undefined) {
			format = this.context.resolver.paragraph(this.direct(node))
			this.paragraphs.set(key, format)
		}

		return format
	}

	// What the paragraph style alone gives (no direct properties).
	styleParagraph(styleId: string | null): ParagraphFormat {
		let format = this.styleParagraphs.get(styleId)

		if (format === undefined) {
			format = this.context.resolver.paragraph(styleId === null ? EMPTY_PARAGRAPH : { styleId })
			this.styleParagraphs.set(styleId, format)
		}

		return format
	}

	// A run's properties from its paragraph and character style, without direct formatting.
	styleRun(paragraph: ParagraphFormat, marksOf: readonly Mark[]): ResolvedRunProps {
		const styleId = (marks.char_style.isInSet(marksOf)?.attrs["styleId"] as string | undefined) ?? null
		let direct = this.characterStyles.get(styleId)

		if (direct === undefined) {
			direct = styleId === null ? {} : { styleId }
			this.characterStyles.set(styleId, direct)
		}

		return this.context.resolver.run(paragraph, direct)
	}

	// Whether content with `marksOf` in `paragraph` is complex script (w:rtl or w:cs in its base or
	// styles): then w:bCs and w:iCs are what count for bold and italic.
	private complex(style: ResolvedRunProps, marksOf: readonly Mark[]): boolean {
		if (style.rtl || style.cs) return true

		const base = marks.run.isInSet(marksOf)?.attrs["base"] as number | undefined
		const props = base === undefined ? undefined : this.context.bases.get(base)?.props

		return props?.rtl === true || props?.cs === true
	}

	// The four toggles of content with `marksOf` in `paragraph` as bits: TOGGLE_ON << k when toggle k is
	// on, TOGGLE_STYLE << k when the styles alone give it. Cached per mark array (unchanged nodes keep
	// theirs), so a snapshot over a large selection reads a map per node.
	toggles(paragraph: ParagraphFormat, marksOf: readonly Mark[]): number {
		const cached = this.toggleCache.get(marksOf)

		if (cached !== undefined && cached.paragraph === paragraph) return cached.bits

		const style = this.styleRun(paragraph, marksOf)
		const complex = this.complex(style, marksOf)
		let bits = 0

		RUN_TOGGLES.forEach((toggle, k) => {
			let on: boolean
			let fromStyle: boolean

			switch (toggle) {
				case "bold":
				case "italic": {
					fromStyle = toggle === "bold" ? (complex ? style.boldCs : style.bold) : complex ? style.italicCs : style.italic

					const mark = marks[toggle].isInSet(marksOf)
					const value = mark?.attrs[complex ? "cs" : "value"] as boolean | null | undefined

					on = value ?? fromStyle

					break
				}
				case "underline": {
					fromStyle = style.underline !== undefined

					const mark = marks.underline.isInSet(marksOf)

					on = mark === undefined ? fromStyle : mark.attrs["style"] !== "none"

					break
				}
				case "strike": {
					fromStyle = style.strike

					const value = marks.strike.isInSet(marksOf)?.attrs["value"] as boolean | null | undefined

					on = value ?? fromStyle

					break
				}
			}

			if (on) bits |= TOGGLE_ON << k
			if (fromStyle) bits |= TOGGLE_STYLE << k
		})

		this.toggleCache.set(marksOf, { paragraph, bits })

		return bits
	}

	// Whether `toggle` is on for content with `marksOf` in `paragraph`, and what the styles alone give.
	toggle(toggle: RunToggle, paragraph: ParagraphFormat, marksOf: readonly Mark[]): { readonly on: boolean; readonly style: boolean } {
		const bits = this.toggles(paragraph, marksOf)
		const k = RUN_TOGGLES.indexOf(toggle)

		return { on: (bits & (TOGGLE_ON << k)) !== 0, style: (bits & (TOGGLE_STYLE << k)) !== 0 }
	}

	// The list a paragraph is in, by its effective numbering.
	listOf(node: Node): { readonly kind: ListKind; readonly numId: number; readonly ilvl: number } | undefined {
		const numbering = this.paragraph(node).props.numbering
		const direct = node.attrs["numId"] as number | null
		const ilvl = (node.attrs["ilvl"] as number | null) ?? numbering?.ilvl ?? 0

		if (direct !== null && direct < 0) {
			const list = this.newLists.get(direct)

			return list === undefined ? undefined : { kind: list.kind, numId: direct, ilvl }
		}

		if (numbering === undefined || numbering.numId <= 0) return undefined

		const level = this.context.numbering?.level(numbering.numId, numbering.ilvl)

		const kind = level === undefined ? undefined : listKind(level.level.format)

		return kind === undefined ? undefined : { kind, numId: numbering.numId, ilvl: numbering.ilvl }
	}

	// The first w:num of the document whose first level is a list of `kind`.
	firstList(kind: ListKind): number | undefined {
		const numbering = this.context.numbering

		if (numbering === undefined) return undefined

		for (const numId of [...numbering.instances.keys()].sort((a, b) => a - b)) {
			const level = numbering.level(numId, 0)

			if (level !== undefined && listKind(level.level.format) === kind) return numId
		}

		return undefined
	}
}

function listKind(format: string): ListKind | undefined {
	return format === "none" ? undefined : format === "bullet" ? "bullet" : "numbered"
}

// ST_Jc to the four alignments commands set (logical: Word reads left/right of a right-to-left paragraph
// as its start/end).
export function alignmentOf(jc: string): "start" | "center" | "end" | "justify" | "other" {
	switch (jc) {
		case "left":
		case "start":
			return "start"
		case "right":
		case "end":
			return "end"
		case "center":
			return "center"
		case "both":
		case "distribute":
		case "lowKashida":
		case "mediumKashida":
		case "highKashida":
		case "thaiDistribute":
			return "justify"
		default:
			return "other"
	}
}
