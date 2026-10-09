import type { ActionDef } from "@/lib/keymap/registry"

// The Videos screens' commands: a library view's search box takes the listings' find combo, and `/` opens Videos
// search from anywhere in the module.
export const VIDEOS_ACTIONS: readonly ActionDef[] = [
	{ id: "videos.search", defaultCombo: "mod+f", scope: "videos", descriptionKey: "videos:videosSearchCommand" },
	{ id: "videos.openSearch", defaultCombo: "slash", scope: "videos", descriptionKey: "videos:videosOpenSearchCommand" }
]
