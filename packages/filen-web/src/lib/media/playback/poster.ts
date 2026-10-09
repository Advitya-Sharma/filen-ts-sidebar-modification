// The frame a player shows, kept as its poster while the next source loads, so a restart (a demotion to
// another rung, a track change the session cannot make in place) never flashes black.

// Wider frames are scaled down: the poster only bridges a moment, and a 4K frame would hold ~30 MB.
const MAX_POSTER_WIDTH = 1920
const POSTER_QUALITY = 0.85

// An object URL of the element's current frame, null when it has none or it cannot be read. The caller
// revokes it.
export async function captureFramePoster(video: HTMLVideoElement): Promise<string | null> {
	if (video.readyState < video.HAVE_CURRENT_DATA || video.videoWidth === 0 || video.videoHeight === 0) {
		return null
	}

	const scale = Math.min(1, MAX_POSTER_WIDTH / video.videoWidth)
	const width = Math.round(video.videoWidth * scale)
	const height = Math.round(video.videoHeight * scale)

	try {
		const canvas = new OffscreenCanvas(width, height)
		const context = canvas.getContext("2d")

		if (context === null) {
			return null
		}

		context.drawImage(video, 0, 0, width, height)

		return URL.createObjectURL(await canvas.convertToBlob({ type: "image/jpeg", quality: POSTER_QUALITY }))
	} catch {
		// A canvas lost under memory pressure, or a frame the engine will not hand out.
		return null
	}
}
