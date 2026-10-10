import { normalizeLanguage } from "./languages"
import { MAX_SUBTITLE_CHARS } from "./subtitles/cues"

interface TextDecoderLike {
	decode(input: Uint8Array): string
}

type TextDecoderCtor = new (label: string, options?: { fatal?: boolean }) => TextDecoderLike

export interface DecodedText {
	text: string
	// WHATWG encoding label actually used.
	encoding: string
	// false when the bytes were not valid UTF-8/UTF-16 and a legacy codepage was assumed (or the caller forced one).
	confident: boolean
}

// Encodings offered by a manual picker (label = WHATWG name).
export const SUBTITLE_ENCODINGS: readonly { label: string; name: string }[] = [
	{ label: "utf-8", name: "Unicode (UTF-8)" },
	{ label: "utf-16le", name: "Unicode (UTF-16 LE)" },
	{ label: "utf-16be", name: "Unicode (UTF-16 BE)" },
	{ label: "windows-1250", name: "Central European (Windows-1250)" },
	{ label: "windows-1251", name: "Cyrillic (Windows-1251)" },
	{ label: "windows-1252", name: "Western European (Windows-1252)" },
	{ label: "windows-1253", name: "Greek (Windows-1253)" },
	{ label: "windows-1254", name: "Turkish (Windows-1254)" },
	{ label: "windows-1255", name: "Hebrew (Windows-1255)" },
	{ label: "windows-1256", name: "Arabic (Windows-1256)" },
	{ label: "windows-1257", name: "Baltic (Windows-1257)" },
	{ label: "windows-1258", name: "Vietnamese (Windows-1258)" },
	{ label: "windows-874", name: "Thai (Windows-874)" },
	{ label: "koi8-r", name: "Cyrillic (KOI8-R)" },
	{ label: "gb18030", name: "Chinese Simplified (GB18030)" },
	{ label: "big5", name: "Chinese Traditional (Big5)" },
	{ label: "shift_jis", name: "Japanese (Shift JIS)" },
	{ label: "euc-jp", name: "Japanese (EUC-JP)" },
	{ label: "euc-kr", name: "Korean (EUC-KR)" }
]

const KNOWN_LABELS: ReadonlySet<string> = new Set(SUBTITLE_ENCODINGS.map(encoding => encoding.label))

// A Map, so a language code never resolves an inherited object key.
const LEGACY_BY_LANGUAGE: ReadonlyMap<string, string> = new Map([
	["ru", "windows-1251"],
	["uk", "windows-1251"],
	["bg", "windows-1251"],
	["sr", "windows-1251"],
	["mk", "windows-1251"],
	["be", "windows-1251"],
	["kk", "windows-1251"],
	["mn", "windows-1251"],
	["el", "windows-1253"],
	["tr", "windows-1254"],
	["az", "windows-1254"],
	["he", "windows-1255"],
	["yi", "windows-1255"],
	["ar", "windows-1256"],
	["fa", "windows-1256"],
	["ur", "windows-1256"],
	["et", "windows-1257"],
	["lv", "windows-1257"],
	["lt", "windows-1257"],
	["pl", "windows-1250"],
	["cs", "windows-1250"],
	["sk", "windows-1250"],
	["hu", "windows-1250"],
	["hr", "windows-1250"],
	["sl", "windows-1250"],
	["ro", "windows-1250"],
	["sq", "windows-1250"],
	["bs", "windows-1250"],
	["vi", "windows-1258"],
	["th", "windows-874"],
	["zh", "gb18030"],
	["yue", "big5"],
	["ja", "shift_jis"],
	["ko", "euc-kr"]
])

const TRADITIONAL_HINT = /traditional|(?:^|[-_ (])(?:tw|hk|mo|hant)(?:$|[-_ )])/

// Legacy codepage most likely for a language hint (BCP 47 tag, 639 code or English name); windows-1252 when unknown.
export function legacyEncodingForLanguage(hint: string | null | undefined): string {
	if (hint === null || hint === undefined || hint.length > 64) {
		return "windows-1252"
	}

	const lowered = hint.trim().toLowerCase()

	if (lowered === "zht" || lowered === "cht") {
		return "big5"
	}

	const lang = normalizeLanguage(lowered)

	if (lang === "zh" && TRADITIONAL_HINT.test(lowered)) {
		return "big5"
	}

	if (lang === "sr" && /(?:^|[-_ (])latn(?:$|[-_ )])|latin/.test(lowered)) {
		return "windows-1250"
	}

	return lang === null ? "windows-1252" : (LEGACY_BY_LANGUAGE.get(lang) ?? "windows-1252")
}

function getDecoder(): TextDecoderCtor {
	return (globalThis as unknown as { TextDecoder: TextDecoderCtor }).TextDecoder
}

function tryDecode(label: string, bytes: Uint8Array, fatal: boolean): string | null {
	try {
		return new (getDecoder())(label, { fatal }).decode(bytes)
	} catch {
		return null
	}
}

// UTF-16 without a BOM: ASCII-heavy text leaves every other byte zero.
function sniffUtf16(bytes: Uint8Array): "utf-16le" | "utf-16be" | null {
	const sample = Math.min(bytes.length, 512) & ~1

	if (sample < 16) {
		return null
	}

	let even = 0
	let odd = 0

	for (let i = 0; i < sample; i += 2) {
		if (bytes[i] === 0) {
			even++
		}

		if (bytes[i + 1] === 0) {
			odd++
		}
	}

	const half = sample / 2

	if (odd > half * 0.3 && even <= half * 0.05) {
		return "utf-16le"
	}

	if (even > half * 0.3 && odd <= half * 0.05) {
		return "utf-16be"
	}

	return null
}

function stripBom(text: string): string {
	return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text
}

export interface DecodeSubtitleOptions {
	// Language of the track (sidecar name token), picks the legacy codepage fallback.
	languageHint?: string | null
	// Manual picker choice (a label from `SUBTITLE_ENCODINGS`); skips detection.
	encoding?: string | null
}

// BOM -> strict UTF-8 -> legacy codepage by language hint. null when the bytes exceed the 8 MiB cap.
export function decodeSubtitleBytes(bytes: Uint8Array, options: DecodeSubtitleOptions = {}): DecodedText | null {
	if (bytes.length > MAX_SUBTITLE_CHARS) {
		return null
	}

	let body = bytes
	let bomLabel: string | null = null

	if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
		bomLabel = "utf-8"
		body = bytes.subarray(3)
	} else if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
		bomLabel = "utf-16le"
		body = bytes.subarray(2)
	} else if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
		bomLabel = "utf-16be"
		body = bytes.subarray(2)
	}

	const forced = options.encoding?.toLowerCase()

	if (forced !== undefined && KNOWN_LABELS.has(forced)) {
		// A BOM of another encoding is a marker, not text: decoded as a legacy codepage it would survive stripBom.
		const text = tryDecode(forced, bomLabel !== null && bomLabel !== forced ? body : bytes, false)

		if (text !== null) {
			return {
				text: stripBom(text),
				encoding: forced,
				confident: false
			}
		}
	}

	if (bomLabel !== null) {
		const strict = tryDecode(bomLabel, body, true)

		if (strict !== null) {
			return {
				text: stripBom(strict),
				encoding: bomLabel,
				confident: true
			}
		}

		if (bomLabel !== "utf-8") {
			// The BOM promised UTF-16 but the payload is malformed: keep what decodes, flagged.
			return {
				text: stripBom(tryDecode(bomLabel, body, false) ?? ""),
				encoding: bomLabel,
				confident: false
			}
		}
	}

	const utf16 = bomLabel === null ? sniffUtf16(bytes) : null

	if (utf16 !== null) {
		const strict = tryDecode(utf16, bytes, true)

		if (strict !== null) {
			return {
				text: strict,
				encoding: utf16,
				confident: true
			}
		}

		return {
			text: tryDecode(utf16, bytes, false) ?? "",
			encoding: utf16,
			confident: false
		}
	}

	const utf8 = tryDecode("utf-8", body, true)

	if (utf8 !== null) {
		return {
			text: stripBom(utf8),
			encoding: "utf-8",
			confident: true
		}
	}

	const legacy = legacyEncodingForLanguage(options.languageHint)
	// `body`: a UTF-8 BOM in front of a legacy payload would decode as "ï»¿" and survive stripBom.
	const text = tryDecode(legacy, body, false) ?? tryDecode("windows-1252", body, false)

	return {
		text: stripBom(text ?? ""),
		encoding: text === null ? "windows-1252" : legacy,
		confident: false
	}
}
