import {
	CueBuilder,
	SubtitleLineReader,
	MAX_SUBTITLE_CHARS,
	normalizeCueText,
	parseSubtitleTimestamp,
	type CueStore,
	type NormalizedCueText
} from "./cues"

const DEFAULT_COLUMNS = ["layer", "start", "end", "style", "name", "marginl", "marginr", "marginv", "effect", "text"]

// First `count - 1` comma-separated fields, then the rest as the final field (the text may contain commas).
function splitFields(line: string, from: number, count: number): string[] | null {
	const out: string[] = []
	let start = from

	for (let k = 0; k < count - 1; k++) {
		const comma = line.indexOf(",", start)

		if (comma === -1) {
			return null
		}

		out.push(line.slice(start, comma))
		start = comma + 1
	}

	out.push(line.slice(start))

	return out
}

// ASS/SSA reduced to plain text: `{...}` overrides are removed except `\anN` (cue alignment hint), `\N`/`\n` become line breaks,
// vector drawings are dropped. Returns null when over the size cap.
export function parseAss(input: string): CueStore | null {
	if (input.length > MAX_SUBTITLE_CHARS) {
		return null
	}

	const reader = new SubtitleLineReader(input.charCodeAt(0) === 0xfeff ? input.slice(1) : input)
	const builder = new CueBuilder()
	let inEvents = false
	let columns = DEFAULT_COLUMNS
	let startIdx = 1
	let endIdx = 2
	let textIdx = 9

	for (let line = reader.next(); line !== null; line = reader.next()) {
		const t = line.trimStart()

		if (t.charCodeAt(0) === 91) {
			inEvents = t.slice(0, 8).toLowerCase() === "[events]"

			continue
		}

		if (!inEvents) {
			continue
		}

		const colon = t.indexOf(":")

		if (colon === -1 || colon > 10) {
			continue
		}

		const key = t.slice(0, colon).toLowerCase()

		if (key === "format") {
			const names = t
				.slice(colon + 1)
				.split(",", 32)
				.map(name => name.trim().toLowerCase())
			const s = names.indexOf("start")
			const e = names.indexOf("end")
			const x = names.indexOf("text")

			if (s !== -1 && e !== -1 && x !== -1) {
				columns = names
				startIdx = s
				endIdx = e
				textIdx = x
			}

			continue
		}

		if (key !== "dialogue") {
			continue
		}

		const fields = splitFields(t, colon + 1, columns.length)

		if (fields === null) {
			continue
		}

		const start = parseSubtitleTimestamp((fields[startIdx] as string).trim())
		const end = parseSubtitleTimestamp((fields[endIdx] as string).trim())

		if (start === null || end === null) {
			continue
		}

		const cleaned = normalizeCueText((fields[textIdx] as string).trimStart(), "ass")

		if (!cleaned.drawing) {
			builder.add(start.seconds, end.seconds, cleaned.text, cleaned.hint)
		}
	}

	return builder.finish()
}

export interface MatroskaAssBlock {
	readOrder: number
	layer: number
	style: string
	name: string
	marginL: string
	marginR: string
	marginV: string
	effect: string
	// Raw override-laden text.
	text: string
}

// A Matroska S_TEXT/ASS block payload: `ReadOrder,Layer,Style,Name,MarginL,MarginR,MarginV,Effect,Text`.
export function parseMatroskaAssBlock(payload: string): MatroskaAssBlock | null {
	const fields = splitFields(payload, 0, 9)

	if (fields === null) {
		return null
	}

	const readOrder = Number(fields[0])
	const layer = Number(fields[1])

	if (!Number.isInteger(readOrder) || !Number.isInteger(layer)) {
		return null
	}

	return {
		readOrder,
		layer,
		style: fields[2] as string,
		name: fields[3] as string,
		marginL: fields[4] as string,
		marginR: fields[5] as string,
		marginV: fields[6] as string,
		effect: fields[7] as string,
		text: fields[8] as string
	}
}

// Plain-text cue for an ASS block; null for unparseable payloads and vector drawings.
export function matroskaAssBlockToCue(payload: string): (NormalizedCueText & { readOrder: number }) | null {
	const block = parseMatroskaAssBlock(payload)

	if (block === null) {
		return null
	}

	const cleaned = normalizeCueText(block.text, "ass")

	return cleaned.drawing || cleaned.text.length === 0
		? null
		: {
				...cleaned,
				readOrder: block.readOrder
			}
}

export function formatAssTime(seconds: number): string {
	const cs = Math.max(0, Math.round(seconds * 100))
	const h = Math.floor(cs / 360000)
	const m = Math.floor((cs % 360000) / 6000)
	const s = Math.floor((cs % 6000) / 100)

	return `${String(h)}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}.${String(cs % 100).padStart(2, "0")}`
}

// Rebuilds the full `Dialogue:` line (standard ASS column order) a renderer expects from a Matroska block and its times.
export function matroskaAssBlockToDialogue(payload: string, startSeconds: number, endSeconds: number): string | null {
	const block = parseMatroskaAssBlock(payload)

	if (block === null) {
		return null
	}

	return `Dialogue: ${String(block.layer)},${formatAssTime(startSeconds)},${formatAssTime(endSeconds)},${block.style},${block.name},${block.marginL},${block.marginR},${block.marginV},${block.effect},${block.text}`
}
