import {
	cssFont,
	forEachSlotRun,
	isChineseLanguage,
	symbolTable,
	symbolText,
	type FontRequest,
	type FontSlot,
	type ResolvedFont,
	type SymbolTable
} from "../../fonts/index"
import { Caps, Strike, verticalAlignment, type StyleTable, type TextStyle } from "../../text/index"
import type { ResolvedBorder, ResolvedRunProps, ResolvedShading } from "../props/types"
import { shadingColor } from "../theme/color"
import { eighthPoints, halfPoints, points, twips } from "./units"

// What layout needs from a font manager (FontManager satisfies it): faces are loaded by prepare() and
// resolve() must only be trusted for requests that were prepared.
export interface FontSource {
	prepare(requests: Iterable<FontRequest>): Promise<void>
	resolve(request: FontRequest): ResolvedFont
}

// Tracked-change look of a run in the Markup view.
export const TrackedLook = { NONE: 0, INSERTED: 1, DELETED: 2, MOVED_FROM: 3, MOVED_TO: 4 } as const

export type TrackedLook = (typeof TrackedLook)[keyof typeof TrackedLook]

export type DecorationValue =
	| { readonly kind: "color"; readonly color: string }
	| { readonly kind: "underline"; readonly style: string; readonly color: string | undefined }
	| {
			readonly kind: "border"
			readonly style: string
			// px
			readonly width: number
			readonly space: number
			readonly color: string | undefined
			readonly shadow: boolean
	  }
	// Tracked-change marks: the renderer picks the author's colour.
	| { readonly kind: "revision"; readonly mark: TrackedLook }

// Interned decoration values; key 0 means none, so keys are 1-based indices into `values` + 1.
export class DecorationTable {
	readonly values: DecorationValue[] = []
	private readonly keys = new Map<string, number>()
	// Resolved borders and shadings are interned by the resolver: their keys are looked up by identity.
	private readonly borders = new WeakMap<ResolvedBorder, number>()
	private readonly shadings = new WeakMap<ResolvedShading, number>()

	private intern(signature: string, value: DecorationValue): number {
		let key = this.keys.get(signature)

		if (key === undefined) {
			this.values.push(value)
			key = this.values.length
			this.keys.set(signature, key)
		}

		return key
	}

	get(key: number): DecorationValue | undefined {
		return this.values[key - 1]
	}

	color(color: string | undefined): number {
		return color === undefined ? 0 : this.intern(`c${color}`, { kind: "color", color })
	}

	shading(shading: ResolvedShading | undefined): number {
		if (shading === undefined) return 0

		let key = this.shadings.get(shading)

		if (key === undefined) {
			key = this.color(shadingColor(shading.pattern, shading.color, shading.fill))
			this.shadings.set(shading, key)
		}

		return key
	}

	underline(style: string, color: string | undefined): number {
		return this.intern(`u${style}|${color ?? ""}`, { kind: "underline", style, color })
	}

	border(border: ResolvedBorder | undefined): number {
		if (border === undefined) return 0

		const cached = this.borders.get(border)

		if (cached !== undefined) return cached

		const width = borderWidth(border)
		const key = this.intern(`b${border.style}|${width}|${border.space}|${border.color ?? ""}|${border.shadow ? 1 : 0}`, {
			kind: "border",
			style: border.style,
			width,
			space: points(border.space),
			color: border.color,
			shadow: border.shadow
		})

		this.borders.set(border, key)

		return key
	}

	revision(mark: TrackedLook): number {
		return mark === TrackedLook.NONE ? 0 : this.intern(`r${mark}`, { kind: "revision", mark })
	}
}

// Drawn thickness of a border in px. Double and multi-line styles take more room than their stroke width.
export function borderWidth(border: ResolvedBorder): number {
	const size = eighthPoints(border.size > 0 ? border.size : 2)

	switch (border.style) {
		case "double":
			return size * 3
		case "triple":
			return size * 5
		case "thinThickSmallGap":
		case "thickThinSmallGap":
		case "thinThickMediumGap":
		case "thickThinMediumGap":
		case "thinThickLargeGap":
		case "thickThinLargeGap":
			return size * 2
		default:
			return size
	}
}

// What the renderer needs per text style id besides TextStyle.
export interface RunStyleInfo {
	readonly css: string
	readonly color: string | undefined
	readonly run: ResolvedRunProps | undefined
	readonly revision: TrackedLook
}

const SLOT_INDEX: Readonly<Record<FontSlot, number>> = { ascii: 0, hAnsi: 1, eastAsia: 2, cs: 3 }
const NON_ASCII = /[^ -~\t\n\r]/

// Maps resolved run formatting to core/text styles, one TextStyle per distinct look, and tracks the fonts
// that were needed before they were prepared.
export class RunStyles {
	readonly table: StyleTable
	readonly decorations: DecorationTable
	readonly info: RunStyleInfo[] = []
	readonly missing: FontRequest[] = []
	private readonly fonts: FontSource
	private readonly byKey = new Map<number, number>()
	private readonly byFont = new Map<string, number>()
	private readonly prepared = new Set<string>()
	private readonly pending = new Set<string>()
	// Per style, the code table of a symbol font drawn without the font itself, whose characters must be
	// drawn as the Unicode characters they show.
	private readonly symbols: (SymbolTable | undefined)[] = []

	constructor(table: StyleTable, fonts: FontSource, decorations: DecorationTable) {
		this.table = table
		this.fonts = fonts
		this.decorations = decorations
	}

	get hasMissing(): boolean {
		return this.missing.length > 0
	}

	// Loads the missing fonts; styles needing them can be made afterwards.
	async prepareMissing(): Promise<void> {
		if (this.missing.length === 0) return

		const requests = this.missing.splice(0)

		await this.fonts.prepare(requests)

		for (const key of this.pending) this.prepared.add(key)

		this.pending.clear()
	}

	// Requests for fonts layout will need, prepared ahead.
	async prepare(requests: readonly FontRequest[]): Promise<void> {
		for (const request of requests) this.need(request)

		await this.prepareMissing()
	}

	// The style for a run's text drawn from `slot`; `fontOverride` names the family instead (w:sym).
	style(run: ResolvedRunProps, slot: FontSlot, revision: TrackedLook = TrackedLook.NONE, fontOverride?: string): number {
		if (fontOverride !== undefined) {
			const key = `${run.id}|${revision}|${fontOverride}`
			const cached = this.byFont.get(key)

			if (cached !== undefined) return cached

			const made = this.make(run, slot, revision, fontOverride)

			if (made.cacheable) this.byFont.set(key, made.id)

			return made.id
		}

		// Slots drawn with the same family share a style, so text switching between them stays one fragment.
		if (slot === "hAnsi" && run.fonts.hAnsi === run.fonts.ascii) slot = "ascii"
		else if (slot === "eastAsia" && run.fonts.eastAsiaUsesAscii) slot = "ascii"

		const key = (run.id * 4 + SLOT_INDEX[slot]) * 8 + revision
		const cached = this.byKey.get(key)

		if (cached !== undefined) return cached

		const made = this.make(run, slot, revision, undefined)

		if (made.cacheable) this.byKey.set(key, made.id)

		return made.id
	}

	// `text` as `style` draws it: a symbol font's codes become Unicode characters when the font is missing.
	drawn(style: number, text: string): string {
		const table = this.symbols[style]

		return table === undefined ? text : symbolText(table, text)
	}

	// Calls `out` for each stretch of `text` drawn by one style.
	forEachStyle(
		text: string,
		run: ResolvedRunProps,
		revision: TrackedLook,
		out: (start: number, end: number, style: number) => void
	): void {
		if (run.cs || run.rtl) {
			out(0, text.length, this.style(run, "cs", revision))

			return
		}

		if (!NON_ASCII.test(text)) {
			out(0, text.length, this.style(run, "ascii", revision))

			return
		}

		forEachSlotRun(
			text,
			{
				hint: run.fonts.hint,
				complexScript: false,
				chineseLanguage: isChineseLanguage(run.lang.eastAsia),
				chineseFontCharset: false
			},
			(start, end, slot) => out(start, end, this.style(run, slot, revision))
		)
	}

	private need(request: FontRequest): boolean {
		const key = `${request.family.toLowerCase()}\u0000${request.bold ? 1 : 0}${request.italic ? 1 : 0}${request.slot}`

		if (this.prepared.has(key)) return true

		if (!this.pending.has(key)) {
			this.pending.add(key)
			this.missing.push(request)
		}

		return false
	}

	private make(
		run: ResolvedRunProps,
		slot: FontSlot,
		revision: TrackedLook,
		fontOverride: string | undefined
	): { id: number; cacheable: boolean } {
		const complex = slot === "cs"
		const family =
			fontOverride ??
			(slot === "ascii"
				? run.fonts.ascii
				: slot === "hAnsi"
					? run.fonts.hAnsi
					: slot === "eastAsia"
						? run.fonts.eastAsiaUsesAscii
							? run.fonts.ascii
							: run.fonts.eastAsia
						: run.fonts.cs)
		const request: FontRequest = {
			family,
			bold: complex ? run.boldCs : run.bold,
			italic: complex ? run.italicCs : run.italic,
			slot
		}
		const ready = this.need(request)
		const font = this.fonts.resolve(request)
		const halfPointSize = complex ? run.sizeCs : run.size
		let size = halfPoints(halfPointSize)
		let shift = halfPoints(run.position)

		if (run.vertAlign !== "baseline") {
			const aligned = verticalAlignment(size, run.vertAlign)

			size = aligned.size
			shift += aligned.baselineShift
		}

		const decorations = this.decorations
		let underline = run.underline === undefined ? 0 : decorations.underline(run.underline.style, run.underline.color ?? run.color)
		let strike: Strike = run.doubleStrike ? Strike.DOUBLE : run.strike ? Strike.SINGLE : Strike.NONE

		if (revision === TrackedLook.INSERTED || revision === TrackedLook.MOVED_TO) {
			underline = decorations.underline(revision === TrackedLook.MOVED_TO ? "double" : "single", undefined)
		} else if (revision === TrackedLook.DELETED || revision === TrackedLook.MOVED_FROM) {
			strike = revision === TrackedLook.MOVED_FROM ? Strike.DOUBLE : Strike.SINGLE
		}

		const style: TextStyle = {
			font,
			size,
			kerning: run.kern > 0 && halfPointSize >= run.kern,
			letterSpacing: twips(run.spacing),
			scaleX: run.scale > 0 ? run.scale / 100 : 1,
			caps: run.caps ? Caps.CAPS : run.smallCaps ? Caps.SMALL_CAPS : Caps.NONE,
			baselineShift: shift,
			underline,
			underlineWords: run.underline?.style === "words",
			strike,
			highlight: decorations.color(run.highlight),
			shading: decorations.shading(run.shading),
			border: decorations.border(run.border)
		}
		const id = this.table.add(style)

		this.info[id] = { css: cssFont(font, size), color: run.color, run, revision }
		this.symbols[id] = font.source === "embedded" || font.source === "local" ? undefined : symbolTable(family)

		return { id, cacheable: ready }
	}
}
