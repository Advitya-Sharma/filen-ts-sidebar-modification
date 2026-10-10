import { describe, expect, it } from "vitest"
import { browserCanPlay, queryAlternatives } from "@/features/preview/lib/containerTracks"

describe("browserCanPlay", () => {
	it("asks for every spelling of the codec, since Firefox knows MP3 in MP4 only as mp3", () => {
		const query = 'audio/mp4; codecs="mp4a.6B"'
		const media = {
			canPlayType: (asked: string) => (asked === 'audio/mp4; codecs="mp3"' ? "maybe" : "")
		} as unknown as HTMLMediaElement

		expect(queryAlternatives(query)).toEqual([
			'audio/mp4; codecs="mp4a.6B"',
			'audio/mp4; codecs="mp4a.69"',
			'audio/mp4; codecs="mp3"',
			'audio/mp4; codecs="mp4a.6b"'
		])
		expect(queryAlternatives("audio/mpeg")).toEqual(["audio/mpeg"])
		expect(browserCanPlay(media)(query)).toBe(true)
		expect(browserCanPlay(media)('audio/mp4; codecs="ac-3"')).toBe(false)
	})
})
