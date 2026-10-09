// Runs `callback` once `signal` aborts, at once when it already has (an abort listener added then would
// never fire). Returns the unsubscribe; a no-op without a signal.
export function onAbort(signal: AbortSignal | undefined, callback: () => void): () => void {
	if (signal === undefined) {
		return () => undefined
	}

	if (signal.aborted) {
		callback()

		return () => undefined
	}

	signal.addEventListener("abort", callback, { once: true })

	return () => {
		signal.removeEventListener("abort", callback)
	}
}
