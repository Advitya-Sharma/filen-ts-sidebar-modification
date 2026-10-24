import { formatBytes } from "@filen/shared"
import { formatDateTime, formatShortDate } from "@/lib/formatDate"
import { asDirectoryOrFile, getSharerIdentity, isSharedRootDriveItem, lastModifiedOf, type DriveItem } from "@/features/drive/lib/item"
import { type DriveVariant } from "@/features/drive/lib/preferences"

// Directories carry no real size on the item itself (synthetic 0n — see narrowItem in
// @/features/drive/lib/item); their true recursive size lives only in the useDriveDirectorySizes map, and
// a caller passes this directory's entry as `directorySize`. Undefined (not resolved yet, or no map at
// all) renders as blank, mirroring filen-mobile's own row (its Size component returns null while the
// query is pending, no loading indicator). A shared file reads as a file, a shared directory as a
// directory (asDirectoryOrFile).
export function formatItemSize(item: DriveItem, directorySize?: number): string {
	const base = asDirectoryOrFile(item)

	if (base.type === "file") {
		return formatBytes(Number(base.data.size))
	}

	return directorySize !== undefined ? formatBytes(directorySize) : ""
}

export function formatModifiedDate(item: DriveItem): string {
	return formatShortDate(lastModifiedOf(item))
}

// The info panel's three rows carry the clock time (formatDateTime): the row's date alone cannot tell an
// upload from an edit made the same day.
export function formatModifiedDateTime(item: DriveItem): string {
	return formatDateTime(lastModifiedOf(item))
}

// The info panel's own "Created" row: both item types carry an OPTIONAL `created` field on their
// decrypted meta (a file's `created` is optional, unlike its required `modified`), falling back to
// the item's own raw timestamp — same fallback formatModifiedDate uses for a directory with no
// `created` field, so an item missing this field never renders a blank/undefined date.
export function formatCreatedDateTime(item: DriveItem): string {
	return formatDateTime(item.data.decryptedMeta?.created ?? item.data.timestamp)
}

// The info panel's own "Uploaded" row: the item's raw server-side upload timestamp, distinct from the
// meta-derived Created/Modified dates — a file re-uploaded over an older name can carry an upload time
// later than its own recorded created/modified. Present on every arm (unlike the optional meta
// fields), so this never falls back. Mirrors filen-mobile's rawUploadTimestamp row.
export function formatUploadedDateTime(item: DriveItem): string {
	return formatDateTime(item.data.timestamp)
}

// The shared counterparty a listing row/tile shows as a muted secondary segment on the two shared
// surfaces: on "shared with me" it's who shared the item (driveSharedByLabel), on "shared with others"
// who it's shared with (driveSharedWithLabel). Returns the i18n key to interpolate plus the
// counterparty email; null for a non-shared variant, or a shared item whose role couldn't be read
// (an undecryptable/unknown role) — the caller renders nothing then, never the badge for the four
// non-shared variants.
//
// Root-only, mirroring filen-mobile's shareEmail.tsx: a nested sharedFile/sharedDirectory inside a
// shared subdirectory shares the same counterparty as every sibling under it, so repeating the email
// on each row is redundant — only the ROOT listing (where one row can be a different sender/recipient
// per item) shows it. sharedRootFile/sharedRootDirectory are structurally root-only (narrowItem's
// fetchSharedListing split narrows a `uuid === null` fetch to the two root arms, a nested fetch to
// sharedFile/sharedDirectory — see item.ts), so this type check alone equals "the listing is at the
// root" without threading a uuid/listing prop through the row/tile.
export function sharedIdentityLabel(
	item: DriveItem,
	variant: DriveVariant
): { labelKey: "driveSharedByLabel" | "driveSharedWithLabel"; name: string } | null {
	if (variant !== "sharedIn" && variant !== "sharedOut") {
		return null
	}

	if (!isSharedRootDriveItem(item)) {
		return null
	}

	const identity = getSharerIdentity(item)

	if (identity === null) {
		return null
	}

	return { labelKey: variant === "sharedIn" ? "driveSharedByLabel" : "driveSharedWithLabel", name: identity.email }
}
