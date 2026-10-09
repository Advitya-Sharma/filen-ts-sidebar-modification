import { type CueStore } from "./cues"
import { parseTimedBlocks } from "./timedBlocks"

// SubRip. `<i>/<b>/<u>` stay as markup, other tags are dropped, `{\anN}` becomes the cue alignment hint.
export function parseSrt(text: string): CueStore | null {
	return parseTimedBlocks(text, "srt")
}
