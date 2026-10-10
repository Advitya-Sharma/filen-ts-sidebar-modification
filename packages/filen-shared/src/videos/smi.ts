import { normalizeLanguage } from "./languages"
import { CueBuilder, MAX_SUBTITLE_CUE_CHARS, MAX_SUBTITLE_CHARS, normalizeCueText, type CueStore } from "./subtitles/cues"

export interface SmiTrack {
	// The `Class` attribute of the `<P>` tags, "" when absent.
	className: string
	// From the style block (`lang:` or `Name:`), normalised; null when unknown.
	lang: string | null
	store: CueStore
}

const TAIL_SECONDS = 5
// Real files carry a handful of classes; every class holds its own cue builder, so entries of any further class are dropped.
const MAX_SMI_CLASSES = 32

function indexOfCI(text: string, needle: string, from: number): number {
	const first = needle.charCodeAt(0)
	const last = text.length - needle.length

	for (let i = from; i <= last; i++) {
		if ((text.charCodeAt(i) | 32) !== first) {
			continue
		}

		let j = 1

		while (j < needle.length && (text.charCodeAt(i + j) | 32) === needle.charCodeAt(j)) {
			j++
		}

		if (j === needle.length) {
			return i
		}
	}

	return -1
}

// Value of `name=value` / `name="value"` inside a tag body, case-insensitive; "" when absent.
function attribute(tag: string, name: string): string {
	let from = 0

	for (;;) {
		const at = indexOfCI(tag, name, from)

		if (at === -1) {
			return ""
		}

		from = at + name.length

		const before = at === 0 ? 32 : tag.charCodeAt(at - 1)

		if (before > 32 && before !== 34 && before !== 39) {
			continue
		}

		let i = from

		while (i < tag.length && tag.charCodeAt(i) <= 32) {
			i++
		}

		if (tag.charCodeAt(i) !== 61) {
			continue
		}

		i++

		while (i < tag.length && tag.charCodeAt(i) <= 32) {
			i++
		}

		const quote = tag.charCodeAt(i)

		if (quote === 34 || quote === 39) {
			const end = tag.indexOf(tag[i] as string, i + 1)

			return tag.slice(i + 1, end === -1 ? tag.length : end)
		}

		let end = i

		while (end < tag.length && tag.charCodeAt(end) > 32) {
			end++
		}

		return tag.slice(i, end)
	}
}

// Class -> language from a style block such as `.ENCC { Name: English; lang: en-US; }`.
function parseStyleLanguages(style: string, out: Map<string, string | null>): void {
	let i = 0

	while (i < style.length) {
		const dot = style.indexOf(".", i)

		if (dot === -1) {
			break
		}

		const open = style.indexOf("{", dot)

		if (open === -1) {
			break
		}

		const close = style.indexOf("}", open)
		const body = style.slice(open + 1, close === -1 ? style.length : close)
		const className = style.slice(dot + 1, open).trim()

		if (className.length > 0 && className.length <= 64) {
			const lang = readDeclaration(body, "lang")
			const name = readDeclaration(body, "name")

			out.set(className, normalizeLanguage(lang) ?? normalizeLanguage(name))
		}

		i = close === -1 ? style.length : close + 1
	}
}

function readDeclaration(body: string, key: string): string {
	const at = indexOfCI(body, key, 0)

	if (at === -1) {
		return ""
	}

	const colon = body.indexOf(":", at)

	if (colon === -1) {
		return ""
	}

	const semi = body.indexOf(";", colon)

	return body.slice(colon + 1, semi === -1 ? body.length : semi).trim()
}

interface ClassState {
	builder: CueBuilder
	pendingStart: number
	pendingRaw: string
	hasPending: boolean
}

// SAMI: `<SYNC Start=ms>` points carrying `<P Class=X>` entries, one track per class. A cue ends at the next sync that
// carries an entry of the same class (usually an empty `&nbsp;` one). Returns null when over the size cap.
export function parseSmi(input: string): SmiTrack[] | null {
	if (input.length > MAX_SUBTITLE_CHARS) {
		return null
	}

	const len = input.length
	const classes = new Map<string, ClassState>()
	const order: string[] = []
	const languages = new Map<string, string | null>()
	let syncMs = 0
	let currentClass: string | null = null
	let inSync = false
	let segmentStart = 0
	let syncHadEntry = false

	const stateFor = (className: string): ClassState | null => {
		let state = classes.get(className)

		if (state === undefined) {
			if (classes.size >= MAX_SMI_CLASSES) {
				return null
			}

			state = {
				builder: new CueBuilder(),
				pendingStart: 0,
				pendingRaw: "",
				hasPending: false
			}

			classes.set(className, state)
			order.push(className)
		}

		return state
	}

	const closePending = (state: ClassState, endMs: number) => {
		if (state.hasPending) {
			const cleaned = normalizeCueText(state.pendingRaw, "smi")

			state.builder.add(state.pendingStart / 1000, Math.max(state.pendingStart, endMs) / 1000, cleaned.text, cleaned.hint)
			state.hasPending = false
			state.pendingRaw = ""
		}
	}

	// A sync point without any entry ends the active cue of every class.
	const endBlankSync = () => {
		if (inSync && !syncHadEntry) {
			for (const state of classes.values()) {
				closePending(state, syncMs)
			}
		}
	}

	// Text collected since the last P/SYNC belongs to the current class.
	const commit = (end: number) => {
		if (!inSync || end <= segmentStart) {
			return
		}

		const raw = input.slice(segmentStart, Math.min(end, segmentStart + MAX_SUBTITLE_CUE_CHARS * 4))

		if (currentClass === null) {
			if (normalizeCueText(raw, "smi").text.length === 0) {
				return
			}

			currentClass = ""
		}

		const state = stateFor(currentClass)

		syncHadEntry = true

		if (state === null) {
			return
		}

		closePending(state, syncMs)
		state.pendingStart = syncMs
		state.pendingRaw = raw
		state.hasPending = true
	}

	let i = 0

	while (i < len) {
		const lt = input.indexOf("<", i)

		if (lt === -1) {
			break
		}

		// Comments carry the style block's declarations in most files; skip them whole.
		if (input.startsWith("<!--", lt)) {
			const end = input.indexOf("-->", lt + 4)

			i = end === -1 ? len : end + 3

			continue
		}

		const gt = input.indexOf(">", lt + 1)

		if (gt === -1) {
			break
		}

		let nameEnd = input.charCodeAt(lt + 1) === 47 ? lt + 2 : lt + 1

		while (nameEnd < gt && input.charCodeAt(nameEnd) > 32 && input.charCodeAt(nameEnd) !== 47) {
			nameEnd++
		}

		const name = input.slice(lt + 1, nameEnd).toLowerCase()

		if (name === "style") {
			const close = indexOfCI(input, "</style", gt)
			const end = close === -1 ? len : close

			parseStyleLanguages(input.slice(gt + 1, end), languages)
			i = close === -1 ? len : input.indexOf(">", close) + 1 || len

			continue
		}

		if (name === "sync") {
			commit(lt)
			endBlankSync()

			const start = Number(attribute(input.slice(lt, gt), "start"))

			if (Number.isFinite(start) && start >= 0) {
				syncMs = start
			}

			inSync = true
			syncHadEntry = false
			currentClass = null
			segmentStart = gt + 1
		} else if (name === "p" && inSync) {
			commit(lt)
			currentClass = attribute(input.slice(lt, gt), "class")
			segmentStart = gt + 1
		} else if (name === "/p" || name === "/sync" || name === "/body" || name === "/sami") {
			commit(lt)
			currentClass = name === "/p" ? currentClass : null
			segmentStart = len + 1
		}

		i = gt + 1
	}

	commit(len)
	endBlankSync()

	const tracks: SmiTrack[] = []

	for (const className of order) {
		const state = classes.get(className) as ClassState

		closePending(state, state.pendingStart + TAIL_SECONDS * 1000)

		const store = state.builder.finish()

		if (store.texts.length > 0) {
			tracks.push({
				className,
				lang: languages.get(className) ?? null,
				store
			})
		}
	}

	return tracks
}
