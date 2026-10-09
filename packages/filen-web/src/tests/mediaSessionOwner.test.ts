import { describe, expect, it, vi } from "vitest"
import { MediaSessionOwners, type MediaPositionState, type MediaSessionHandler, type MediaSessionLike } from "@/lib/media/mediaSessionOwner"
import { createMediaSessionPublisher, bindMediaSessionActions } from "@/features/audio/lib/mediaSession"

function fakeSession() {
	const handlers = new Map<string, MediaSessionHandler>()
	const positions: (MediaPositionState | undefined)[] = []
	const session: MediaSessionLike = {
		metadata: null,
		playbackState: "none",
		setActionHandler: (action, handler) => {
			if (action === "skipad") {
				throw new Error("NotSupportedError")
			}

			handlers.set(action, handler)
		},
		setPositionState: state => {
			positions.push(state)
		}
	}

	return { session, handlers, positions }
}

describe("MediaSessionOwners", () => {
	it("lets only the topmost owner write, and gives the session back on release", () => {
		const { session, handlers, positions } = fakeSession()
		const owners = new MediaSessionOwners(session)
		const pauseAudio = vi.fn()
		const audio = owners.create("audio", { onPreempt: pauseAudio })
		const audioPlay = vi.fn()

		audio.acquire()
		audio.metadata = "song"
		audio.playbackState = "playing"
		audio.setActionHandler("play", audioPlay)
		audio.setActionHandler("nexttrack", () => undefined)
		audio.setPositionState({ duration: 200, position: 10, playbackRate: 1 })

		expect(session.metadata).toBe("song")
		expect(handlers.get("play")).toBe(audioPlay)

		const video = owners.create("video")
		const videoPlay = vi.fn()

		video.acquire()

		expect(pauseAudio).toHaveBeenCalledTimes(1)
		// The new owner starts clean: nothing of the audio's is left on the session.
		expect(session.metadata).toBeNull()
		expect(session.playbackState).toBe("none")
		expect(handlers.get("play")).toBeNull()
		expect(handlers.get("nexttrack")).toBeNull()
		expect(positions.at(-1)).toBeUndefined()

		video.setMetadata("movie")
		video.setPlaybackState("playing")
		video.setActionHandler("play", videoPlay)
		// The paused engine's own writes stay recorded but don't reach the session.
		audio.playbackState = "paused"

		expect(session.metadata).toBe("movie")
		expect(session.playbackState).toBe("playing")
		expect(handlers.get("play")).toBe(videoPlay)
		expect(audio.active).toBe(false)
		expect(video.active).toBe(true)

		video.release()

		expect(session.metadata).toBe("song")
		expect(session.playbackState).toBe("paused")
		expect(handlers.get("play")).toBe(audioPlay)
		expect(positions.at(-1)).toEqual({ duration: 200, position: 10, playbackRate: 1 })
		expect(audio.active).toBe(true)
	})

	it("survives an action the browser refuses and an owner released twice", () => {
		const { session } = fakeSession()
		const owners = new MediaSessionOwners(session)
		const audio = owners.create("audio")

		audio.acquire()
		expect(() => {
			audio.setActionHandler("skipad", () => undefined)
		}).not.toThrow()
		audio.release()
		audio.release()
		expect(session.playbackState).toBe("none")
	})

	it("is a no-op without Media Session support", () => {
		const owners = new MediaSessionOwners(null)
		const owner = owners.create("video")

		owner.acquire()
		owner.setMetadata("x")
		owner.setPositionState({ duration: 1, position: 0, playbackRate: 1 })
		owner.release()

		expect(owner.metadata).toBe("x")
	})

	it("carries the audio engine's publisher and action binder", () => {
		const { session, handlers } = fakeSession()
		const owners = new MediaSessionOwners(session)
		const audio = owners.create("audio")
		const target = { resume: vi.fn(), pause: vi.fn(), skipNext: vi.fn(), skipPrevious: vi.fn(), seek: vi.fn() }

		audio.acquire()
		createMediaSessionPublisher(audio).setPlaybackState("playing")
		bindMediaSessionActions(target, () => 5, audio)

		expect(session.playbackState).toBe("playing")

		handlers.get("seekforward")?.({ action: "seekforward" })
		expect(target.seek).toHaveBeenCalledWith(15)

		owners.create("video").acquire()
		createMediaSessionPublisher(audio).setPlaybackState("paused")
		expect(session.playbackState).toBe("none")
	})
})
