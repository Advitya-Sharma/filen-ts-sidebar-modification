import type { Mark, MarkType, Node } from "prosemirror-model"
import type { EditorState, Transaction } from "prosemirror-state"
import type { StoryPart } from "../model/index"
import {
	EMPTY_PARAGRAPH,
	type ParagraphProps,
	type ResolvedRunProps,
	type ParagraphBordersProps,
	type RunProps,
	type SpacingProps
} from "../props/index"
import type { ParagraphFormat } from "../styles/index"
import { HIGHLIGHT_COLORS } from "../theme/color"
import type { DocxHighlightColor } from "./commands/formatting"
import type { DocxEditContext } from "./context"
import { runPropsOf } from "./formatting"
import { MarkChanges, MarkSetChanges, sameSet } from "./marks"
import { holdsFormatting } from "../fields/format"
import { FORMATTING_MARKS, NO_SOURCE, docxSchema, inlineLevel, type ColorAttrs, type IndentAttrs } from "./schema"

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
const nodes = docxSchema.nodes
const PROPERTY_TYPES: ReadonlySet<MarkType> = new Set(FORMATTING_MARKS.filter(name => name !== "plain").map(name => marks[name]))
const HIGHLIGHT_NAMES = new Map(Object.entries(HIGHLIGHT_COLORS).map(([name, hex]) => [hex, name as DocxHighlightColor]))

// The formatting values of run content (what the toolbar shows and what value commands compare with).
// Interned: equal values are one object.
export interface RunValues {
	// The font of the content's script (w:ascii, or w:cs for complex script).
	readonly font: string
	// Half-points.
	readonly size: number
	// "auto", "RRGGBB" or "theme:<name>" with "/tint:<n>" or "/shade:<n>" (0–255) as the direct colour says;
	// what the styles give is "auto" or "RRGGBB".
	readonly color: string
	// The colour shown: "auto" or "RRGGBB".
	readonly colorRgb: string
	readonly highlight: DocxHighlightColor
	readonly vertAlign: "baseline" | "superscript" | "subscript"
	readonly strike: boolean
	readonly doubleStrike: boolean
	readonly caps: boolean
	readonly smallCaps: boolean
	// Character spacing, twips.
	readonly spacing: number
	// Complex script content (w:rtl or w:cs): size and font are the w:szCs and w:cs ones.
	readonly complex: boolean
}

interface Formatting {
	readonly props: RunProps
	// The character style alone.
	readonly style: RunProps
	// The source w:rPr marks complex script content (w:rtl, w:cs).
	readonly complexBase: boolean
	readonly color: string | undefined
	readonly values: Map<ParagraphFormat, RunValues>
	readonly styleValues: Map<ParagraphFormat, RunValues>
}

interface TrieNode {
	next: Map<Mark, TrieNode> | undefined
	// Without and with a complex script base.
	readonly leaf: [Formatting | undefined, Formatting | undefined]
}

// The colour key of a direct w:color (see RunValues.color).
export function colorKey(attrs: ColorAttrs): string {
	if (attrs.themeColor !== null && attrs.themeColor !== "none") {
		const tone = attrs.themeTint !== null ? `/tint:${attrs.themeTint}` : attrs.themeShade !== null ? `/shade:${attrs.themeShade}` : ""

		return `theme:${attrs.themeColor}${tone}`
	}

	return attrs.value === null || attrs.value === "auto" ? "auto" : attrs.value.toUpperCase()
}

// Built-in styles commands add whose type is not paragraph, or whose id Word does not make from the name.
const BUILT_IN_STYLES: Readonly<Record<string, { readonly type: "paragraph" | "character"; readonly id: string }>> = {
	"annotation text": { type: "paragraph", id: "CommentText" },
	"annotation reference": { type: "character", id: "CommentReference" },
	"footnote reference": { type: "character", id: "FootnoteReference" },
	"endnote reference": { type: "character", id: "EndnoteReference" },
	"Placeholder Text": { type: "character", id: "PlaceholderText" },
	...Object.fromEntries(
		[1, 2, 3, 4, 5, 6, 7, 8, 9].map(level => [`toc ${String(level)}`, { type: "paragraph", id: `TOC${String(level)}` } as const])
	)
}

export function builtInStyleType(name: string): "paragraph" | "character" {
	return BUILT_IN_STYLES[name]?.type ?? "paragraph"
}

// The id of built-in style `name` (a name the writer has a definition for): the document's own, else one
// added to the styles the writer defines when content uses it (Word's id, made unique against the document's).
export function builtInStyleId(context: DocxEditContext, name: string): string {
	const type = builtInStyleType(name)
	const styles = context.styles
	const existing = styles.byName(name, type)

	if (existing !== undefined) return existing.id

	for (const [id, missing] of context.newStyles) if (missing === name) return id

	const base = BUILT_IN_STYLES[name]?.id ?? name.replace(/(^|\s)(\S)/g, (_match, _space: string, letter: string) => letter.toUpperCase())
	let id = base

	for (let n = 1; styles.byId(id) !== undefined || context.newStyles.has(id); n++) id = `${base}${n}`

	context.newStyles.set(id, name)

	return id
}

// Run-level content of a selection as [from, to) pieces with their paragraph (what formatting changes: not
// markers or locked blocks); `cell` when the paragraph is in a table (a table style may format it,
// which the resolver does not see).
export interface RunPiece {
	readonly from: number
	readonly to: number
	readonly node: Node
	readonly paragraph: Node
	readonly cell: boolean
}

export function runPieces(state: EditorState, ranges?: readonly { readonly from: number; readonly to: number }[]): RunPiece[] {
	return runPiecesIn(state.doc, ranges ?? state.selection.ranges.map(range => ({ from: range.$from.pos, to: range.$to.pos })))
}

export function runPiecesIn(doc: Node, ranges: readonly { readonly from: number; readonly to: number }[]): RunPiece[] {
	const out: RunPiece[] = []

	for (const { from, to } of ranges) {
		// Ends of the tables around the content visited (nodesBetween visits a range's ancestors too).
		const tables: number[] = []

		doc.nodesBetween(from, to, (node, pos, parent) => {
			if (!node.isInline) {
				if (node.type === nodes.table) tables.push(pos + node.nodeSize)

				return node.type !== nodes.locked_block
			}

			// A field takes formatting as its runs do (§17.16.4.3; the writer gives them what changed).
			if (
				parent !== null &&
				parent.type === nodes.paragraph &&
				(inlineLevel(node) === "run" || (node.type === nodes.field && holdsFormatting(node)))
			) {
				const start = Math.max(from, pos)
				const end = Math.min(to, pos + node.nodeSize)

				while (tables.length > 0 && (tables[tables.length - 1] ?? 0) <= pos) tables.pop()

				if (end > start) out.push({ from: start, to: end, node, paragraph: parent, cell: tables.length > 0 })
			}

			return false
		})
	}

	return out
}

export interface ParagraphAt {
	readonly node: Node
	readonly pos: number
	// In a table (see RunPiece.cell).
	readonly cell: boolean
}

// The paragraphs the selection touches, in document order.
export function selectedParagraphs(state: EditorState): ParagraphAt[] {
	const out: ParagraphAt[] = []
	const seen = new Set<number>()

	for (const range of state.selection.ranges) {
		const tables: number[] = []

		state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, pos) => {
			if (node.type === nodes.paragraph) {
				while (tables.length > 0 && (tables[tables.length - 1] ?? 0) <= pos) tables.pop()

				if (!seen.has(pos)) {
					seen.add(pos)
					out.push({ node, pos, cell: tables.length > 0 })
				}

				return false
			}

			if (node.type === nodes.table) tables.push(pos + node.nodeSize)

			return !node.isInline && node.type !== nodes.locked_block
		})
	}

	return out
}

// Sets, piece by piece, the mark of `type` that `next` gives (undefined: none), as one step (see SetMarksStep:
// undo restores what each piece had).
export function changeMarks(
	tr: Transaction,
	pieces: readonly RunPiece[],
	type: MarkType,
	next: (piece: RunPiece, index: number) => Mark | undefined
): void {
	const changes = new MarkChanges()
	let sample: Mark | undefined

	pieces.forEach((piece, index) => {
		const existing = type.isInSet(piece.node.marks)
		const wanted = next(piece, index)

		if (existing === undefined ? wanted === undefined : wanted !== undefined && existing.eq(wanted)) return

		sample ??= wanted ?? existing
		changes.add(piece.from, piece.to, wanted ?? null)
	})

	if (sample !== undefined) changes.step(tr, sample)
}

// Sets, piece by piece, the marks of `types` that `next` gives (a set as addToSet makes it), as one step (see
// SetMarkSetsStep); `mark` is what readers of an AddMarkStep see.
export function changeMarkSets(
	tr: Transaction,
	pieces: readonly RunPiece[],
	types: readonly MarkType[],
	next: (piece: RunPiece) => readonly Mark[],
	mark: Mark
): void {
	const changes = new MarkSetChanges(types)

	for (const piece of pieces) {
		const wanted = next(piece)

		if (
			!sameSet(
				piece.node.marks.filter(own => types.includes(own.type)),
				wanted
			)
		)
			changes.add(piece.from, piece.to, wanted)
	}

	changes.step(tr, mark)
}

// The marks `marksOf` with the mark of `type` replaced by `mark` (removed when undefined).
export function withMark(marksOf: readonly Mark[], type: MarkType, mark: Mark | undefined): readonly Mark[] {
	const without = type.removeFromSet(marksOf)

	return mark === undefined ? without : mark.addToSet(without)
}

// Effective formatting of editing nodes: a paragraph's direct properties are its source w:pPr with the
// attributes commands edit laid over; a run's toggles are its direct marks, else what the paragraph and
// character styles give. Everything is cached by identity (attribute objects are shared by unchanged node
// copies, resolver results are interned).
export class FormatResolver {
	private readonly context: DocxEditContext
	private readonly part: StoryPart
	private readonly paragraphs = new WeakMap<object, ParagraphFormat>()
	private readonly sources: (ParagraphProps | undefined)[] = []
	private readonly styleParagraphs = new Map<string | null, ParagraphFormat>()
	private readonly characterStyles = new Map<string | null, RunProps>()
	private readonly overrides = new WeakMap<ParagraphProps, Map<string, ParagraphProps>>()
	private readonly toggleCache = new WeakMap<readonly Mark[], { paragraph: ParagraphFormat; bits: number }>()
	// Distinct direct formattings (by the property marks' attributes and the base's script), and their values
	// per paragraph format: keyed by content, not by mark array (every source run has its own array), so a
	// large selection adds nothing per run.
	private readonly markKeys = new WeakMap<Mark, string>()
	private readonly formattings = new Map<string, Formatting>()
	private readonly trie: TrieNode = { next: undefined, leaf: [undefined, undefined] }
	private readonly interned = new WeakMap<ResolvedRunProps, Map<string, RunValues>>()
	// Lists created in this session (negative numIds) and their kind.
	readonly newLists: Map<number, NewList>

	constructor(context: DocxEditContext, part: StoryPart) {
		this.context = context
		this.part = part
		this.newLists = context.newLists
	}

	// The paragraph properties of a paragraph node's source w:pPr (none for a new paragraph), by source item:
	// the store interns them by their bytes, which a command over a long document would otherwise hash again
	// for every paragraph.
	sourceProps(node: Node): ParagraphProps {
		const src = node.attrs["src"] as number

		if (src === NO_SOURCE || this.part.kinds[src] === undefined) return EMPTY_PARAGRAPH

		let props = this.sources[src]

		if (props === undefined) {
			props = this.context.store.paragraph(this.part.document, this.part.paragraphs.pPr[this.part.aux[src] ?? -1] ?? -1)
			this.sources[src] = props
		}

		return props
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
		const spacing = node.attrs["spacing"] as SpacingProps | null
		const borders = node.attrs["borders"] as ParagraphBordersProps | null

		if (
			styleId === (source.styleId ?? null) &&
			numId === (source.numbering?.numId ?? null) &&
			ilvl === (source.numbering?.ilvl ?? null) &&
			align === (source.alignment ?? null) &&
			indent === (source.indent ?? null) &&
			spacing === (source.spacing ?? null) &&
			borders === (source.borders ?? null)
		) {
			return source
		}

		const key = JSON.stringify([styleId, numId, ilvl, align, indent, spacing, borders])
		let bySource = this.overrides.get(source)

		if (bySource === undefined) {
			bySource = new Map()
			this.overrides.set(source, bySource)
		}

		let direct = bySource.get(key)

		if (direct === undefined) {
			const {
				styleId: _styleId,
				numbering: _numbering,
				alignment: _alignment,
				indent: _indent,
				spacing: _spacing,
				borders: _borders,
				...rest
			} = source

			direct = {
				...rest,
				...(styleId !== null ? { styleId } : {}),
				...(numId !== null
					? { numbering: ilvl === null ? { numId } : { numId, ilvl } }
					: ilvl !== null
						? { numbering: { ilvl } }
						: {}),
				...(align !== null ? { alignment: align } : {}),
				...(indent !== null ? { indent } : {}),
				...(spacing !== null ? { spacing } : {}),
				...(borders !== null ? { borders } : {})
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
		return this.context.resolver.run(
			paragraph,
			this.characterStyle((marks.char_style.isInSet(marksOf)?.attrs["styleId"] as string | undefined) ?? null)
		)
	}

	private characterStyle(styleId: string | null): RunProps {
		let direct = this.characterStyles.get(styleId)

		if (direct === undefined) {
			direct = styleId === null ? {} : { styleId }
			this.characterStyles.set(styleId, direct)
		}

		return direct
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

	private markKey(mark: Mark): string {
		let key = this.markKeys.get(mark)

		if (key === undefined) {
			key = `${mark.type.name}${JSON.stringify(mark.attrs)}`
			this.markKeys.set(mark, key)
		}

		return key
	}

	// The formatting of content with `marksOf`: found through a trie over the property marks' identities
	// (source runs share their base's mark objects), so a lookup allocates nothing; a new path is matched to
	// an existing formatting by content.
	private formatting(marksOf: readonly Mark[]): Formatting {
		let complexBase = false
		let node = this.trie

		for (const mark of marksOf) {
			if (PROPERTY_TYPES.has(mark.type)) {
				let next = node.next?.get(mark)

				if (next === undefined) {
					next = { next: undefined, leaf: [undefined, undefined] }
					node.next ??= new Map()
					node.next.set(mark, next)
				}

				node = next
			} else if (mark.type === marks.run) {
				const props = this.context.bases.get(mark.attrs["base"] as number)?.props

				complexBase = props?.rtl === true || props?.cs === true
			}
		}

		const slot = complexBase ? 1 : 0
		let formatting = node.leaf[slot]

		if (formatting === undefined) {
			const own = marksOf.filter(mark => PROPERTY_TYPES.has(mark.type))
			const key = own.map(mark => this.markKey(mark)).join("") + (complexBase ? "|cs" : "")

			formatting = this.formattings.get(key)

			if (formatting === undefined) {
				const color = marks.color.isInSet(own)
				const props = runPropsOf(own)

				formatting = {
					props,
					style: this.characterStyle(props.styleId ?? null),
					complexBase,
					color: color === undefined ? undefined : colorKey(color.attrs as ColorAttrs),
					values: new Map(),
					styleValues: new Map()
				}
				this.formattings.set(key, formatting)
			}

			node.leaf[slot] = formatting
		}

		return formatting
	}

	// The direct properties the formatting marks of `marksOf` stand for, one object per distinct set.
	runProps(marksOf: readonly Mark[]): RunProps {
		return this.formatting(marksOf).props
	}

	// The effective formatting of content with `marksOf` in `paragraph`.
	resolvedRun(paragraph: ParagraphFormat, marksOf: readonly Mark[]): ResolvedRunProps {
		return this.context.resolver.run(paragraph, this.runProps(marksOf))
	}

	// The formatting values of content with `marksOf` in `paragraph`.
	values(paragraph: ParagraphFormat, marksOf: readonly Mark[]): RunValues {
		const formatting = this.formatting(marksOf)
		let values = formatting.values.get(paragraph)

		if (values === undefined) {
			const resolved = this.context.resolver.run(paragraph, formatting.props)

			values = this.intern(resolved, formatting.color, formatting.complexBase || resolved.rtl || resolved.cs)
			formatting.values.set(paragraph, values)
		}

		return values
	}

	// What the paragraph and character styles alone give content with `marksOf` (no direct formatting).
	styleValues(paragraph: ParagraphFormat, marksOf: readonly Mark[]): RunValues {
		const formatting = this.formatting(marksOf)
		let values = formatting.styleValues.get(paragraph)

		if (values === undefined) {
			const resolved = this.context.resolver.run(paragraph, formatting.style)

			values = this.intern(resolved, undefined, formatting.complexBase || resolved.rtl || resolved.cs)
			formatting.styleValues.set(paragraph, values)
		}

		return values
	}

	private intern(resolved: ResolvedRunProps, directColor: string | undefined, complex: boolean): RunValues {
		let byKey = this.interned.get(resolved)

		if (byKey === undefined) {
			byKey = new Map()
			this.interned.set(resolved, byKey)
		}

		const key = `${directColor ?? ""}|${complex ? 1 : 0}`
		let values = byKey.get(key)

		if (values === undefined) {
			const rgb = resolved.color ?? "auto"

			values = {
				font: complex ? resolved.fonts.cs : resolved.fonts.ascii,
				size: complex ? resolved.sizeCs : resolved.size,
				color: directColor ?? rgb,
				colorRgb: rgb,
				highlight: resolved.highlight === undefined ? "none" : (HIGHLIGHT_NAMES.get(resolved.highlight) ?? "none"),
				vertAlign: resolved.vertAlign,
				strike: resolved.strike,
				doubleStrike: resolved.doubleStrike,
				caps: resolved.caps,
				smallCaps: resolved.smallCaps,
				spacing: resolved.spacing,
				complex
			}
			byKey.set(key, values)
		}

		return values
	}

	// The paragraph `node` would resolve to with `attrs` laid over its attributes.
	paragraphWith(node: Node, attrs: Readonly<Record<string, unknown>>): ParagraphFormat {
		return this.context.resolver.paragraph(this.direct(node.type.create({ ...node.attrs, ...attrs })))
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
