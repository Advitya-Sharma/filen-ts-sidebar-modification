import type { VideoRootSize } from "@/workers/videoIndex/view.logic"

// A root's walk is one request whose whole body the SDK buffers and parses in wasm memory, which never shrinks
// again, before the worker filters it; the index then keeps ~180 B per video-related file. A walk lists every
// entry below the root, videos or not. At ~800 B of response per entry, 20k entries is a ~16 MB body (a few times
// that while parsed), and 50k approaches the ~150 MB wasm budget for the whole worker. Above these a root is
// still indexed; the UI warns and suggests a narrower directory.
export const ROOT_ENTRIES_LARGE = 20_000
export const ROOT_ENTRIES_HUGE = 50_000

export function rootSize(walkedEntries: number): VideoRootSize {
	return walkedEntries >= ROOT_ENTRIES_HUGE ? "huge" : walkedEntries >= ROOT_ENTRIES_LARGE ? "large" : "normal"
}
