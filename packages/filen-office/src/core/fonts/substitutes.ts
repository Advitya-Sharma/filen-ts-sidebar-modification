// Free substitutes for the fonts documents name most. Arimo, Tinos, Cousine and Gelasio have measured
// advance widths identical to Arial, Times New Roman, Courier New and Georgia, so lines break where they
// break in Word. Carlito claims Calibri's widths, unverified against Calibri itself; Caladea makes no
// claim for Cambria and is only the closest free match; Selawik is Microsoft's own open fallback for
// Segoe UI, widths unverified. Substitutes apply even when the original is installed, so layout is the
// same on every device. This module is plain data so the table generator can read it.

export type BundledFamily = "carlito" | "caladea" | "arimo" | "tinos" | "cousine" | "gelasio" | "selawik"

export type GenericFamily = "serif" | "sans-serif" | "monospace"

export interface Substitution {
	readonly family: BundledFamily
	readonly generic: GenericFamily
	// Font-size factor bringing the substitute's widths closer to the original's.
	readonly scale: number
	// The weight a family name implies ("Segoe UI Semibold"); undefined for regular families.
	readonly weight: number | undefined
	// Use an installed font of the original name first: there is no free width-compatible clone, and only
	// the real font lays out exactly.
	readonly preferLocal: boolean
}

// Size factors for setting Aptos text in Carlito. No free font matches Aptos, and nothing derived from
// Microsoft's font files may ship, so these stay 1 until measured from page breaks recorded in publicly
// available Word documents.
export const UNCALIBRATED_APTOS_SCALE = 1
export const UNCALIBRATED_APTOS_DISPLAY_SCALE = 1
export const UNCALIBRATED_APTOS_NARROW_SCALE = 1

function substitute(
	family: BundledFamily,
	generic: GenericFamily,
	options: { scale?: number; weight?: number; preferLocal?: boolean } = {}
): Substitution {
	return { family, generic, scale: options.scale ?? 1, weight: options.weight, preferLocal: options.preferLocal ?? false }
}

// Keyed by the lower-cased family name as documents write it.
export const SUBSTITUTIONS: ReadonlyMap<string, Substitution> = new Map([
	["calibri", substitute("carlito", "sans-serif")],
	// Carlito has no light weight: Calibri Light headings come out in Carlito Regular.
	["calibri light", substitute("carlito", "sans-serif")],
	["cambria", substitute("caladea", "serif")],
	["arial", substitute("arimo", "sans-serif")],
	// Arial was drawn to Helvetica's widths.
	["helvetica", substitute("arimo", "sans-serif")],
	["times new roman", substitute("tinos", "serif")],
	["courier new", substitute("cousine", "monospace")],
	["georgia", substitute("gelasio", "serif")],
	["segoe ui", substitute("selawik", "sans-serif")],
	["segoe ui light", substitute("selawik", "sans-serif", { weight: 300 })],
	["segoe ui semilight", substitute("selawik", "sans-serif", { weight: 350 })],
	["segoe ui semibold", substitute("selawik", "sans-serif", { weight: 600 })],
	["aptos", substitute("carlito", "sans-serif", { scale: UNCALIBRATED_APTOS_SCALE, preferLocal: true })],
	["aptos display", substitute("carlito", "sans-serif", { scale: UNCALIBRATED_APTOS_DISPLAY_SCALE, preferLocal: true })],
	["aptos narrow", substitute("carlito", "sans-serif", { scale: UNCALIBRATED_APTOS_NARROW_SCALE, preferLocal: true })],
	["aptos light", substitute("carlito", "sans-serif", { scale: UNCALIBRATED_APTOS_SCALE, weight: 300, preferLocal: true })],
	["aptos semibold", substitute("carlito", "sans-serif", { scale: UNCALIBRATED_APTOS_SCALE, weight: 600, preferLocal: true })],
	["aptos serif", substitute("tinos", "serif", { preferLocal: true })],
	["aptos mono", substitute("cousine", "monospace", { preferLocal: true })]
])

// w:family in a font table (ECMA-376 ST_FontFamily).
export type FontClass = "roman" | "swiss" | "modern" | "script" | "decorative" | "auto"

// w:pitch (ST_Pitch).
export type FontPitch = "fixed" | "variable" | "default"

export interface FontHints {
	readonly fontClass: FontClass | undefined
	readonly pitch: FontPitch | undefined
	// The 10 PANOSE bytes, when the document gives them.
	readonly panose: Uint8Array | undefined
	// w:charset (a Windows character set number).
	readonly charset: number | undefined
}

export const NO_HINTS: FontHints = { fontClass: undefined, pitch: undefined, panose: undefined, charset: undefined }

export interface Fallback {
	// undefined: no bundled family fits; the platform's generic family draws the text.
	readonly family: BundledFamily | undefined
	readonly generic: GenericFamily
}

// Windows character sets the bundled families cover (ANSI, default, Mac, Greek, Turkish, Vietnamese,
// Baltic, Cyrillic, Central European). Fonts declared for anything else (symbol, CJK, Hebrew, Arabic,
// Thai ...) are left to the platform, whose fonts draw those scripts.
const LATIN_CHARSETS = new Set([0, 1, 77, 161, 162, 163, 186, 204, 238])

// PANOSE (Latin text): byte 0 family kind, byte 1 serif style, byte 3 proportion.
const PANOSE_LATIN_TEXT = 2
const PANOSE_MONOSPACED = 9
// Normal, obtuse and perpendicular sans, flared, rounded.
const PANOSE_FIRST_SANS = 11
const PANOSE_LAST_SANS = 15

// The bundled family to stand in for a font we have no substitute for, from the font table's hints. Text
// in the East Asian or complex-script slots goes to the platform: the bundled families are Latin fonts.
export function fallbackFor(name: string, hints: FontHints, latinSlot: boolean): Fallback {
	const lower = name.toLowerCase()
	const panose = hints.panose
	const latin = latinSlot && (hints.charset === undefined || LATIN_CHARSETS.has(hints.charset))
	const monospaced =
		hints.pitch === "fixed" ||
		hints.fontClass === "modern" ||
		(panose !== undefined && panose[0] === PANOSE_LATIN_TEXT && panose[3] === PANOSE_MONOSPACED) ||
		lower.includes("mono") ||
		lower.includes("courier")

	if (monospaced) return { family: latin ? "cousine" : undefined, generic: "monospace" }

	let serif: boolean

	if (hints.fontClass === "roman") serif = true
	else if (hints.fontClass === "swiss" || hints.fontClass === "script" || hints.fontClass === "decorative") serif = false
	else if (panose !== undefined && panose[0] === PANOSE_LATIN_TEXT && (panose[1] ?? 0) >= 2) {
		serif = (panose[1] ?? 0) < PANOSE_FIRST_SANS || (panose[1] ?? 0) > PANOSE_LAST_SANS
	} else {
		serif = !lower.includes("sans")
	}

	const generic: GenericFamily = serif ? "serif" : "sans-serif"

	return { family: latin ? (serif ? "tinos" : "arimo") : undefined, generic }
}
