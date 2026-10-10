import { describe, it, expect } from "vitest"
import { playerTrackVerdict, recordPlayerVerdict } from "@/lib/playerVerdict"
import fileKindVerdicts from "@/lib/fileKindVerdicts"

const TRACK = {}

function loaded(video: number, audio: number) {
	return {
		availableVideoTracks: Array.from({ length: video }, () => TRACK) as never[],
		availableAudioTracks: Array.from({ length: audio }, () => TRACK) as never[]
	}
}

describe("playerTrackVerdict", () => {
	it("makes a listed picture track a video", () => {
		expect(playerTrackVerdict(loaded(1, 1), false)).toBe("video")
		expect(playerTrackVerdict(loaded(1, 0), false)).toBe("video")
	})

	it("makes listed audio tracks without a picture audio", () => {
		expect(playerTrackVerdict(loaded(0, 1), false)).toBe("audio")
		expect(playerTrackVerdict(loaded(0, 2), false)).toBe("audio")
	})

	it("says nothing when no track is listed, whatever else loaded", () => {
		expect(playerTrackVerdict(loaded(0, 0), false)).toBeNull()
	})

	it("says nothing for a failed load", () => {
		expect(playerTrackVerdict(loaded(0, 1), true)).toBeNull()
		expect(playerTrackVerdict(loaded(1, 0), true)).toBeNull()
	})
})

describe("recordPlayerVerdict", () => {
	it("keeps audio for an .mp4, and nothing that confirms its name or rests on no evidence", () => {
		recordPlayerVerdict({ uuid: "voice", name: "voice.mp4", mime: "video/mp4", size: 10 }, loaded(0, 1), false)
		recordPlayerVerdict({ uuid: "clip", name: "clip.mp4", mime: "video/mp4", size: 10 }, loaded(1, 1), false)
		recordPlayerVerdict({ uuid: "broken", name: "broken.mp4", mime: null, size: 10 }, loaded(0, 1), true)
		recordPlayerVerdict({ uuid: "silent", name: "silent.mp4", mime: null, size: 10 }, loaded(0, 0), false)

		expect(fileKindVerdicts.get("voice")).toBe("audio")
		expect(fileKindVerdicts.get("clip")).toBeUndefined()
		expect(fileKindVerdicts.get("broken")).toBeUndefined()
		expect(fileKindVerdicts.get("silent")).toBeUndefined()
	})
})
