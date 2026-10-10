import { driveItemMime, type DriveItem } from "@/features/drive/lib/item"
import { itemFileKind, needsImageTransform, isStreamedCategory } from "@/features/drive/lib/preview.logic"
import { isAllowedInlineContentType } from "@/lib/sw/protocol"

// The page-side half of the inline-preview allowlist gate (the SW's own independent re-check lives in
// sw.ts via the same isAllowedInlineContentType — see protocol.ts's own doc comment). First derives
// the category from the item's name/extension (previewType — extension-first, mime-fallback, already
// tested), which is trusted code-owned classification, not raw user input; only once that resolves to
// a streamable category does the item's OWN decrypted mime (attacker-controlled for a file someone
// else encrypted and shared in) get checked against the shared allowlist. Neither the category gate
// nor the mime check alone is sufficient — a whole-buffer category with a spoofed video/audio mime
// must still fall through here, and a genuinely video/audio/image file with an unrecognized or absent
// mime has nothing safe to serve as its inline Content-Type. `null` means "not inline-streamable" —
// every caller falls back to the buffered blob path or a plain download, never a hard error.
// Audio types by the extension that settled a file's kind, for a file stored with no type or a generic binary
// one (`application/octet-stream`): code-owned, so the route serves what the extension promises and the browser
// plays it as it is instead of having it decoded. A specific type that is not allowlisted is still refused.
const AUDIO_TYPES_BY_EXTENSION: ReadonlyMap<string, string> = new Map([
	["mp3", "audio/mpeg"],
	["m4a", "audio/mp4"],
	["m4b", "audio/mp4"],
	["aac", "audio/aac"],
	["wav", "audio/wav"],
	["ogg", "audio/ogg"],
	["oga", "audio/ogg"],
	["opus", "audio/ogg"],
	["weba", "audio/webm"],
	["flac", "audio/flac"],
	["aiff", "audio/aiff"],
	["aif", "audio/aiff"],
	["caf", "audio/x-caf"]
])

const GENERIC_TYPES: ReadonlySet<string> = new Set(["", "application/octet-stream", "binary/octet-stream"])

export function allowedMediaContentType(item: DriveItem): string | null {
	const kind = itemFileKind(item)
	const category = kind?.kind ?? "other"

	// "rawImage" is not a streamed category: a RAW container has no browser decoder, so serving one
	// inline would hand the page bytes nothing can render, and a camera's RAW mime (image/x-nikon-nef
	// and friends) must never be treated as a streamable image just because it starts with "image/".
	if (!isStreamedCategory(category)) {
		return null
	}

	// HEIC/HEIF resolve to "image" but can never stream (no browser decodes them inline) — excluded
	// independently of the mime check below, so a spoofed streamable mime on a HEIC-named file can't
	// slip through (defense-in-depth: imageViewer.tsx's own dispatch checks this first too).
	if (needsImageTransform(item)) {
		return null
	}

	const normalized = driveItemMime(item)?.toLowerCase().trim()

	if (normalized !== undefined && isAllowedInlineContentType(normalized)) {
		return normalized
	}

	const generic = normalized === undefined || GENERIC_TYPES.has(normalized)

	return generic && category === "audio" && kind?.source === "extension" ? (AUDIO_TYPES_BY_EXTENSION.get(kind.extension) ?? null) : null
}
