import type { OpfsBlobStore } from "@/lib/storage/opfs"

// The Videos artwork store: posters, backdrops and stills derived from sidecar images, embedded covers
// and generated frames. Its own OPFS tree and cap; drive's thumbnail tree is only ever read.

// Which variant of a source an entry is: poster, backdrop or still.
export type ArtSlotCode = "p" | "b" | "s"

export const ART_DIR_ROOT = ["videos", "art"]

// Bumped when the stored images change shape (sizes, encoding).
export const ART_GENERATION = "v1"

export const ART_DIR = [...ART_DIR_ROOT, ART_GENERATION]

// Most entries are WebP; Safari's canvas encodes JPEG instead, which an <img> sniffs the same. A write still out
// when a sign-out (in any tab) wipes the tree must not bring back an image of the signed-out account.
export const ART_STORE: OpfsBlobStore = { dir: ART_DIR, ext: ".webp", recreateWhenRemoved: false }

// 128 MiB: a few thousand posters and stills at 15-120 KB each.
export const ART_CACHE_CAP = 134_217_728

// What an entry derives from: the source file's own image (a sidecar, a video's embedded cover), a frame
// generated from a video, or a frame the player captured. Each has its own entry, so none overwrites
// another made from the same video.
export type ArtVariant = "image" | "frame" | "capture"

const VARIANT_SUFFIXES: Readonly<Record<ArtVariant, string>> = { image: "", frame: "f", capture: "c" }

// Keyed by the uuid of the file the image derives from, so a content change (a new uuid) invalidates it.
export function artKey(sourceUuid: string, slot: ArtSlotCode, variant: ArtVariant = "image"): string {
	return `${sourceUuid}.${slot}${VARIANT_SUFFIXES[variant]}`
}
