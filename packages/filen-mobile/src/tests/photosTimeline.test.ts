import { vi, describe, it, expect } from "vitest"

vi.mock("@/constants", () => ({
	EXPO_IMAGE_SUPPORTED_EXTENSIONS: new Set<string>([".jpg"]),
	EXPO_VIDEO_SUPPORTED_EXTENSIONS: new Set<string>([".mp4"]),
	EXPO_AUDIO_SUPPORTED_EXTENSIONS: new Set<string>([".mp3"])
}))

vi.mock("@/lib/serializer", () => ({
	serialize: (value: unknown) => JSON.stringify(value)
}))

import { buildPhotoTimeline } from "@/features/photos/utils"

describe("buildPhotoTimeline", () => {
	it("records one marker per year and one entry per month at the first (newest) photo", () => {
		const times = [
			new Date(2024, 2, 20).getTime(),
			new Date(2024, 2, 1).getTime(),
			new Date(2024, 0, 5).getTime(),
			new Date(2023, 11, 31).getTime(),
			new Date(2021, 5, 1).getTime()
		]
		const timeline = buildPhotoTimeline(times, t => t)

		expect(Array.from(timeline.timestamps)).toEqual(times)
		expect(timeline.yearMarkers).toEqual([
			{ label: "2024", index: 0 },
			{ label: "2023", index: 3 },
			{ label: "2021", index: 4 }
		])
		expect(timeline.months).toEqual([
			{ year: 2024, month: 2, index: 0 },
			{ year: 2024, month: 0, index: 2 },
			{ year: 2023, month: 11, index: 3 },
			{ year: 2021, month: 5, index: 4 }
		])
	})

	it("handles an empty listing", () => {
		const timeline = buildPhotoTimeline<number>([], t => t)

		expect(timeline.timestamps.length).toBe(0)
		expect(timeline.yearMarkers).toEqual([])
		expect(timeline.months).toEqual([])
	})
})
