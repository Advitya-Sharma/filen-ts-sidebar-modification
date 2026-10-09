import { type CueStore } from "./cues"
import { parseTimedBlocks } from "./timedBlocks"

// WebVTT. Cue settings: only `line:` (as an alignment hint). Voice/class/timestamp tags are stripped to their text;
// NOTE, STYLE and REGION blocks are ignored. Returns null when over the size cap.
export function parseVtt(text: string): CueStore | null {
	return parseTimedBlocks(text, "vtt")
}
