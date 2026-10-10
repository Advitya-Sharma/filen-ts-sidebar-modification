import { activeCueIndices, type CueStore } from "@filen/shared/videos"

// How subtitles look over the picture, and which cues show where. Pure: the style record is persisted per
// device (subtitleStyle.ts), the overlay component turns this module's output into DOM.

export const SUBTITLE_SIZES = [0.8, 1, 1.36, 1.72, 2, 2.2] as const
export const SUBTITLE_FONTS = ["sans", "serif", "mono"] as const
export const SUBTITLE_COLORS = ["white", "yellow", "cyan", "green", "magenta"] as const
export const SUBTITLE_BACKGROUNDS = ["none", "translucent", "solid"] as const
export const SUBTITLE_EDGES = ["none", "shadow", "outline", "raised", "depressed"] as const
// Distance of the lowest line above the picture's bottom edge, percent of the picture height.
export const SUBTITLE_POSITIONS = [2, 6, 10, 15, 20] as const

export type SubtitleFont = (typeof SUBTITLE_FONTS)[number]
export type SubtitleColor = (typeof SUBTITLE_COLORS)[number]
export type SubtitleBackground = (typeof SUBTITLE_BACKGROUNDS)[number]
export type SubtitleEdge = (typeof SUBTITLE_EDGES)[number]

export interface SubtitleStyle {
	// Index into SUBTITLE_SIZES.
	size: number
	font: SubtitleFont
	color: SubtitleColor
	background: SubtitleBackground
	edge: SubtitleEdge
	// Index into SUBTITLE_POSITIONS.
	position: number
}

export const DEFAULT_SUBTITLE_STYLE: SubtitleStyle = {
	size: 2,
	font: "sans",
	color: "white",
	background: "none",
	edge: "outline",
	position: 1
}

function clampIndex(value: number, length: number, fallback: number): number {
	return Number.isInteger(value) && value >= 0 && value < length ? value : fallback
}

export function normalizeSubtitleStyle(style: SubtitleStyle): SubtitleStyle {
	return {
		size: clampIndex(style.size, SUBTITLE_SIZES.length, DEFAULT_SUBTITLE_STYLE.size),
		font: SUBTITLE_FONTS.includes(style.font) ? style.font : DEFAULT_SUBTITLE_STYLE.font,
		color: SUBTITLE_COLORS.includes(style.color) ? style.color : DEFAULT_SUBTITLE_STYLE.color,
		background: SUBTITLE_BACKGROUNDS.includes(style.background) ? style.background : DEFAULT_SUBTITLE_STYLE.background,
		edge: SUBTITLE_EDGES.includes(style.edge) ? style.edge : DEFAULT_SUBTITLE_STYLE.edge,
		position: clampIndex(style.position, SUBTITLE_POSITIONS.length, DEFAULT_SUBTITLE_STYLE.position)
	}
}

const COLOR_VALUES: Readonly<Record<SubtitleColor, string>> = {
	white: "#ffffff",
	yellow: "#fde047",
	cyan: "#67e8f9",
	green: "#86efac",
	magenta: "#f0abfc"
}

const FONT_VALUES: Readonly<Record<SubtitleFont, string>> = {
	sans: '"Inter Variable", system-ui, sans-serif',
	serif: 'Georgia, "Times New Roman", serif',
	mono: 'ui-monospace, "SF Mono", Menlo, Consolas, monospace'
}

const BACKGROUND_VALUES: Readonly<Record<SubtitleBackground, string>> = {
	none: "transparent",
	translucent: "rgb(0 0 0 / 0.55)",
	solid: "rgb(0 0 0 / 0.9)"
}

// Edges are drawn with text-shadow so they scale with the text and never need a second layer.
const EDGE_VALUES: Readonly<Record<SubtitleEdge, string>> = {
	none: "none",
	shadow: "0 0.05em 0.12em rgb(0 0 0 / 0.9), 0 0 0.3em rgb(0 0 0 / 0.5)",
	outline: "-0.05em -0.05em 0 #000, 0.05em -0.05em 0 #000, -0.05em 0.05em 0 #000, 0.05em 0.05em 0 #000, 0 0 0.08em #000",
	raised: "-0.04em -0.04em 0 rgb(255 255 255 / 0.5), 0.05em 0.05em 0 #000",
	depressed: "0.04em 0.04em 0 rgb(255 255 255 / 0.5), -0.05em -0.05em 0 #000"
}

export interface SubtitleCss {
	// Multiplier of the overlay's base size (a share of the picture height).
	scale: number
	fontFamily: string
	color: string
	background: string
	textShadow: string
	// Percent of the picture height.
	bottom: number
}

export function subtitleCss(style: SubtitleStyle): SubtitleCss {
	const normalized = normalizeSubtitleStyle(style)

	return {
		scale: SUBTITLE_SIZES[normalized.size] ?? 1,
		fontFamily: FONT_VALUES[normalized.font],
		color: COLOR_VALUES[normalized.color],
		background: BACKGROUND_VALUES[normalized.background],
		textShadow: EDGE_VALUES[normalized.edge],
		bottom: SUBTITLE_POSITIONS[normalized.position] ?? 6
	}
}

export type SubtitleRegionKey = "bottom" | "middle" | "top"
export type SubtitleAlign = "left" | "center" | "right"

export interface SubtitleRegion {
	key: SubtitleRegionKey
	align: SubtitleAlign
	// Cue payloads (escaped text with <i>, <b> and <u> only), one per cue, in start order.
	cues: string[]
}

// Numpad alignment ({\anN}): 1-3 bottom, 4-6 middle, 7-9 top; left, centre, right within each row. 0 is bottom centre.
function regionOf(hint: number): { key: SubtitleRegionKey; align: SubtitleAlign } {
	if (hint < 1 || hint > 9) {
		return { key: "bottom", align: "center" }
	}

	const row = Math.floor((hint - 1) / 3)
	const column = (hint - 1) % 3

	return {
		key: row === 0 ? "bottom" : row === 1 ? "middle" : "top",
		align: column === 0 ? "left" : column === 1 ? "center" : "right"
	}
}

// What the overlay shows at `time` (on the video's timeline) with the cues shifted by `offset` seconds
// (positive = later): at most MAX_CONCURRENT_CUES, which is all a cue store lets show at once. Regions in a stable
// order, empty ones left out.
export function subtitleRegionsAt(store: CueStore, time: number, offset: number, scratch: number[] = []): SubtitleRegion[] {
	const indices = activeCueIndices(store, time - offset, scratch)

	if (indices.length === 0) {
		return []
	}

	const regions: SubtitleRegion[] = []

	for (const index of indices) {
		const text = store.texts[index]

		if (text === undefined || text.length === 0) {
			continue
		}

		const { key, align } = regionOf(store.hints?.[index] ?? 0)
		const region = regions.find(entry => entry.key === key && entry.align === align)

		if (region === undefined) {
			regions.push({ key, align, cues: [text] })
		} else {
			region.cues.push(text)
		}
	}

	return regions
}

// One string per showing state, so a re-render can be skipped when nothing changed.
export function subtitleRegionsKey(regions: readonly SubtitleRegion[]): string {
	return regions.map(region => `${region.key}:${region.align}:${region.cues.join("\u0000")}`).join("\u0001")
}

// Subtitle delay: ±30 s in 0.1 s steps.
export const MAX_SUBTITLE_DELAY_SECONDS = 30

// "+0.3 s", "−1.2 s", "0.0 s".
export function formatSubtitleDelay(seconds: number): string {
	const value = clampSubtitleDelay(seconds)

	return `${value > 0 ? "+" : value < 0 ? "−" : ""}${Math.abs(value).toFixed(1)} s`
}

export function clampSubtitleDelay(seconds: number): number {
	if (!Number.isFinite(seconds)) {
		return 0
	}

	const rounded = Math.round(seconds * 10) / 10

	return Math.max(-MAX_SUBTITLE_DELAY_SECONDS, Math.min(MAX_SUBTITLE_DELAY_SECONDS, rounded)) || 0
}

export interface CueSegment {
	text: string
	italic: boolean
	bold: boolean
	underline: boolean
}

const CUE_ENTITIES: Readonly<Record<string, string>> = { amp: "&", lt: "<", gt: ">", nbsp: "\u00a0", lrm: "\u200e", rlm: "\u200f" }
const CUE_TOKEN = /<(\/?)([ibu])>|&(amp|lt|gt|nbsp|lrm|rlm);/g

// A cue payload (escaped text with <i>, <b>, <u>) as styled text runs: rendered as text nodes, never as HTML.
export function cueSegments(payload: string): CueSegment[] {
	const out: CueSegment[] = []
	let italic = 0
	let bold = 0
	let underline = 0
	let text = ""
	let last = 0

	function flush(): void {
		if (text.length > 0) {
			out.push({ text, italic: italic > 0, bold: bold > 0, underline: underline > 0 })
			text = ""
		}
	}

	for (const match of payload.matchAll(CUE_TOKEN)) {
		text += payload.slice(last, match.index)
		last = match.index + match[0].length

		const tag = match[2]
		const entity = match[3]

		if (entity !== undefined) {
			text += CUE_ENTITIES[entity] ?? ""

			continue
		}

		flush()

		const step = match[1] === "/" ? -1 : 1

		if (tag === "i") {
			italic = Math.max(0, italic + step)
		} else if (tag === "b") {
			bold = Math.max(0, bold + step)
		} else {
			underline = Math.max(0, underline + step)
		}
	}

	text += payload.slice(last)
	flush()

	return out
}
