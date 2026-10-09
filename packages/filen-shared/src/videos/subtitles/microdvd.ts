import { CueBuilder, SubtitleLineReader, MAX_SUBTITLE_CUE_CHARS, MAX_SUBTITLE_CHARS, normalizeCueText, type CueStore } from "./cues"

export const DEFAULT_MICRODVD_FPS = 23.976

// `{digits}{digits}` at `from`; returns the two numbers and the index after.
function readFrames(line: string, from: number): { a: number; b: number; end: number } | null {
	const read = (at: number): { value: number; end: number } | null => {
		if (line.charCodeAt(at) !== 123) {
			return null
		}

		let i = at + 1
		let n = 0
		let digits = 0

		while (i < line.length) {
			const d = line.charCodeAt(i) - 48

			if (d < 0 || d > 9) {
				break
			}

			if (digits < 10) {
				n = n * 10 + d
			}

			digits++
			i++
		}

		return digits > 0 && line.charCodeAt(i) === 125
			? {
					value: n,
					end: i + 1
				}
			: null
	}

	const first = read(from)
	const second = first === null ? null : read(first.end)

	return first === null || second === null
		? null
		: {
				a: first.value,
				b: second.value,
				end: second.end
			}
}

// The first non-empty line must start with `{n}{n}` (this is what separates MicroDVD from VobSub `.sub`).
export function isMicroDvd(text: string): boolean {
	const reader = new SubtitleLineReader(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text)

	for (let line = reader.next(); line !== null; line = reader.next()) {
		const t = line.trimStart()

		if (t.length > 0) {
			return readFrames(t, 0) !== null
		}
	}

	return false
}

// Control codes `{y:i}` `{c:$..}` `{f:..}`: italics/bold/underline whole-line flags become markup, the rest is dropped.
function convertControlCodes(text: string): string {
	let open = ""
	let close = ""
	let out = ""
	let i = 0

	while (i < text.length) {
		if (text.charCodeAt(i) === 123) {
			let end = -1

			for (let j = i + 1; j < Math.min(text.length, i + 41); j++) {
				if (text.charCodeAt(j) === 125) {
					end = j

					break
				}
			}

			if (end !== -1 && text.charCodeAt(i + 2) === 58) {
				const key = text.charCodeAt(i + 1) | 32
				const value = text.slice(i + 3, end).toLowerCase()

				if (key === 121) {
					for (const ch of value) {
						if (ch === "i" || ch === "b" || ch === "u") {
							open += `<${ch}>`
							close = `</${ch}>${close}`
						}
					}
				}

				i = end + 1

				continue
			}
		}

		out += text[i]
		i++
	}

	return open + out + close
}

// `{startFrame}{endFrame}text` with `|` as line break. A leading `{1}{1}<fps>` line sets the frame rate; otherwise `fps` is used.
// Returns null when over the size cap.
export function parseMicroDvd(input: string, fps: number = DEFAULT_MICRODVD_FPS): CueStore | null {
	if (input.length > MAX_SUBTITLE_CHARS) {
		return null
	}

	let rate = Number.isFinite(fps) && fps > 0 ? fps : DEFAULT_MICRODVD_FPS
	const reader = new SubtitleLineReader(input.charCodeAt(0) === 0xfeff ? input.slice(1) : input)
	const builder = new CueBuilder()
	let first = true

	for (let line = reader.next(); line !== null; line = reader.next()) {
		const t = line.trimStart()
		const frames = readFrames(t, 0)

		if (frames === null) {
			continue
		}

		const body = t.slice(frames.end, frames.end + MAX_SUBTITLE_CUE_CHARS * 2)

		if (first) {
			first = false

			if (frames.a <= 1 && frames.b <= 1) {
				const header = Number(body.trim())

				if (Number.isFinite(header) && header >= 1 && header <= 240) {
					rate = header

					continue
				}
			}
		}

		const cleaned = normalizeCueText(convertControlCodes(body.replace(/\|/g, "\n")), "srt")

		builder.add(frames.a / rate, Math.max(frames.a, frames.b) / rate, cleaned.text, cleaned.hint)
	}

	return builder.finish()
}
