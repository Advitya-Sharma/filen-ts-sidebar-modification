import { BlockSource, OutOfFileReadError, ReadBudgetError, type ReadRange } from "@/lib/media/blockSource"
import { MalformedError, UnsupportedContainerError, ascii, readBytes, uint32 } from "@/lib/media/probe/bytes"
import { LEADING_BOXES, walkIsobmff } from "@/lib/media/probe/isobmff"
import { EBML_ID, walkMatroska } from "@/lib/media/probe/matroska"
import type { ParsedContainer, ProbeTier, ProbeVideo, VideoProbe } from "@/lib/media/probe/types"

// The container probe: which tier reads how much, and the record each tier yields.

// A card shows badges for files scrolling past, so it never pays for more than this; a full or cover
// probe starts here and is raised by the walk once the header's size is known (an MP4 `moov` read
// header by header, a Matroska SeekHead jump, a hop between attachments).
export const CARD_BLOCK_READS = 4
const INITIAL_BLOCK_READS: Readonly<Record<ProbeTier, number>> = { card: CARD_BLOCK_READS, full: 4, cover: 4 }

// Tracks kept per kind, so a file with dozens of subtitle tracks still makes a small record.
const MAX_TRACKS_PER_KIND = 64

export function initialBlockReads(tier: ProbeTier): number {
	return INITIAL_BLOCK_READS[tier]
}

// The block source one probe of one file reads through, budgeted for its tier.
export function probeSource(size: number, readRange: ReadRange, tier: ProbeTier): BlockSource {
	return new BlockSource(size, readRange, { maxBlockReads: initialBlockReads(tier) })
}

// The container's raw walk, or null when the file is no container this reads. The first 16 bytes decide,
// so anything else costs one block. A walk that runs out of budget or into a malformed structure
// returns what it found, flagged partial; failed downloads throw.
export async function walkContainer(source: BlockSource, tier: ProbeTier): Promise<ParsedContainer | null> {
	if (source.size < 8) {
		return null
	}

	const head = await readBytes(source, 0, Math.min(16, source.size))
	const matroska = uint32(head, 0) === EBML_ID

	if (!matroska && !LEADING_BOXES.has(ascii(head, 4, 8))) {
		return null
	}

	const result: ParsedContainer = {
		container: matroska ? "mkv" : "mp4",
		durationSec: null,
		index: null,
		tracks: [],
		tracksFound: false,
		chapters: [],
		cover: null,
		partial: false
	}

	try {
		await (matroska ? walkMatroska(source, tier, result) : walkIsobmff(source, tier, result))
	} catch (error) {
		if (error instanceof UnsupportedContainerError) {
			return null
		}

		// A read past the file's end means a size the file cannot hold: malformed, like any other.
		if (!(error instanceof ReadBudgetError || error instanceof MalformedError || error instanceof OutOfFileReadError)) {
			throw error
		}

		result.partial = true
	}

	return result
}

// The probe record for one tier. A cover probe fills only `container` and `cover`.
function toRecord(parsed: ParsedContainer, tier: ProbeTier): VideoProbe {
	// The first enabled video track is the picture; a disabled one only stands in when nothing else does.
	let video: ProbeVideo | null = null
	let videoEnabled = false
	const audio: VideoProbe["audio"] = []
	const subs: VideoProbe["subs"] = []

	for (const track of parsed.tracks) {
		if (track.kind === "video") {
			if (video === null || (track.enabled && !videoEnabled)) {
				video = track.video
				videoEnabled = track.enabled
			}
		} else if (track.kind === "audio" && audio.length < MAX_TRACKS_PER_KIND) {
			audio.push(track.audio)
		} else if (track.kind === "subtitle" && subs.length < MAX_TRACKS_PER_KIND) {
			subs.push(track.sub)
		}
	}

	if (tier === "cover") {
		return {
			v: 1,
			container: parsed.container,
			durationSec: null,
			index: null,
			video: null,
			audio: [],
			subs: [],
			chapters: [],
			cover: parsed.cover,
			partial: parsed.partial
		}
	}

	return {
		v: 1,
		container: parsed.container,
		durationSec: parsed.durationSec,
		index: parsed.index,
		video,
		audio,
		subs,
		chapters: tier === "full" ? parsed.chapters : [],
		cover: tier === "full" ? parsed.cover : null,
		partial: parsed.partial || !parsed.tracksFound
	}
}

// Reads one tier of a file's probe record through `source` (from probeSource). Null for a file that is
// no MP4/MOV/Matroska/WebM; throws BlockReadError when a download fails.
export async function probeVideo(source: BlockSource, tier: ProbeTier): Promise<VideoProbe | null> {
	const parsed = await walkContainer(source, tier)

	return parsed === null ? null : toRecord(parsed, tier)
}
