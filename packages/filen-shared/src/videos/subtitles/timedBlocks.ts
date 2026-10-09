import {
	CueBuilder,
	SubtitleLineReader,
	MAX_SUBTITLE_CUE_CHARS,
	MAX_SUBTITLE_CHARS,
	normalizeCueText,
	parseSubtitleTimestamp,
	type CueStore
} from "./cues"

function vttLineHint(settings: string): number {
	const at = settings.indexOf("line:")

	if (at === -1) {
		return 0
	}

	let end = at + 5

	while (end < settings.length && settings.charCodeAt(end) > 32 && settings.charCodeAt(end) !== 44) {
		end++
	}

	const value = settings.slice(at + 5, end)
	const percent = value.endsWith("%")
	const n = Number(percent ? value.slice(0, -1) : value)

	if (!Number.isFinite(n)) {
		return 0
	}

	if (percent) {
		return n <= 33 ? 8 : n <= 66 ? 5 : 2
	}

	return n >= 0 && n <= 2 ? 8 : 0
}

interface Timing {
	start: number
	end: number
	hint: number
}

// `start --> end [settings]`; null when the line is not a timing line.
function parseTimingLine(line: string, vtt: boolean): Timing | null {
	const arrow = line.indexOf("-->")

	if (arrow === -1) {
		return null
	}

	let i = 0

	while (i < arrow && line.charCodeAt(i) <= 32) {
		i++
	}

	const left = parseSubtitleTimestamp(line, i)

	if (left === null) {
		return null
	}

	i = arrow + 3

	while (i < line.length && line.charCodeAt(i) <= 32) {
		i++
	}

	const right = parseSubtitleTimestamp(line, i)

	if (right === null) {
		return null
	}

	return {
		start: left.seconds,
		end: right.seconds,
		hint: vtt ? vttLineHint(line.slice(right.end)) : 0
	}
}

function isIndexLine(line: string): boolean {
	const len = line.length

	if (len === 0 || len > 10) {
		return false
	}

	for (let i = 0; i < len; i++) {
		const d = line.charCodeAt(i) - 48

		if (d < 0 || d > 9) {
			return false
		}
	}

	return true
}

// Shared by SRT and WebVTT: blocks of an optional identifier, a timing line and text lines up to a blank line.
// Returns null when the input exceeds the size cap.
export function parseTimedBlocks(input: string, mode: "srt" | "vtt"): CueStore | null {
	if (input.length > MAX_SUBTITLE_CHARS) {
		return null
	}

	const vtt = mode === "vtt"
	const reader = new SubtitleLineReader(input.charCodeAt(0) === 0xfeff ? input.slice(1) : input)
	const builder = new CueBuilder()
	const rawLimit = MAX_SUBTITLE_CUE_CHARS * 2
	let timing: Timing | null = null
	let raw = ""

	const flush = () => {
		if (timing !== null && raw.length > 0) {
			const cleaned = normalizeCueText(raw, mode)

			builder.add(timing.start, timing.end, cleaned.text, cleaned.hint !== 0 ? cleaned.hint : timing.hint)
		}

		timing = null
		raw = ""
	}

	for (let line = reader.next(); line !== null; line = reader.next()) {
		if (line.indexOf("-->") !== -1) {
			const parsed = parseTimingLine(line, vtt)

			if (parsed !== null) {
				flush()
				timing = parsed

				continue
			}
		}

		if (timing === null) {
			continue
		}

		if (line.trim().length === 0) {
			flush()

			continue
		}

		// An index line directly followed by a timing line ends a cue that lacks its blank separator.
		if (!vtt && isIndexLine(line)) {
			const upcoming = reader.peek()

			if (upcoming !== null && parseTimingLine(upcoming, false) !== null) {
				flush()

				continue
			}
		}

		if (raw.length < rawLimit) {
			raw = raw.length === 0 ? line : `${raw}\n${line}`
		}
	}

	flush()

	return builder.finish()
}
