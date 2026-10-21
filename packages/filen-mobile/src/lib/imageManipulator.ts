import * as ImageManipulator from "expo-image-manipulator"

// Decode `uri` (already in expo's encoded file:// form), optionally resize to `resizeWidth`, and
// save a new file. Every expo-image-manipulator decode goes through here so the two rules below
// can't be forgotten by a new caller.
export async function renderAndSave(
	uri: string,
	save: { format: ImageManipulator.SaveFormat; compress?: number },
	resizeWidth?: number
): Promise<ImageManipulator.ImageResult> {
	// Hold the Context in a local binding across the awaits: it cancels its native coroutine on
	// sharedObjectDidRelease, so a Context that became Hermes-GC-eligible mid-render would reject
	// with JobCancellationException.
	//
	// With a target width the source is also DECODED no wider than it (expo-image-manipulator 58): the
	// full-size bitmap is never materialised just to be scaled down. resize() still runs, since the
	// decode only bounds the width from above.
	const context =
		resizeWidth === undefined
			? ImageManipulator.ImageManipulator.manipulate(uri)
			: ImageManipulator.ImageManipulator.manipulate(uri, {
					maxWidth: resizeWidth
				}).resize({
					width: resizeWidth
				})

	// The Context and the rendered ImageRef wrap decoded native bitmaps Hermes GC does not track;
	// bulk callers (camera upload, thumbnails) decode many images back-to-back, so both are
	// released on every exit path or native memory piles up until the OS kills the app.
	let rendered: ImageManipulator.ImageRef | null = null

	try {
		rendered = await context.renderAsync()

		return await rendered.saveAsync({
			...save,
			base64: false
		})
	} finally {
		rendered?.release()
		context.release()
	}
}
