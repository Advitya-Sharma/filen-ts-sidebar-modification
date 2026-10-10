import { useEffect, useState, useSyncExternalStore } from "react"
import { watchState } from "@/features/videos/lib/watchState"

// The local watch state is read before the first rows are asked for; this long at most, so a slow or failed read
// never holds the library back.
const WATCH_STATE_WAIT_MS = 300

// Whether this device's watch state (journal and cached file) is loaded, or the wait for it is over.
export function useWatchStateLoaded(): boolean {
	const loaded = useSyncExternalStore(watchState.subscribe, () => watchState.getSnapshot().loaded)
	const [waited, setWaited] = useState(false)

	useEffect(() => {
		if (loaded) {
			return
		}

		const timer = setTimeout(() => {
			setWaited(true)
		}, WATCH_STATE_WAIT_MS)

		return () => {
			clearTimeout(timer)
		}
	}, [loaded])

	return loaded || waited
}
