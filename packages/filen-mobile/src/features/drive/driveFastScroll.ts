import { type DriveItem } from "@/types"
import { type SortByType, sortTimestampKey } from "@/lib/sort"
import { type FastScrollOptions, type FastScrollMarker } from "@/components/ui/fastScroller"
import { driveItemName, formatBytes } from "@filen/shared"
import { monthYearLabel } from "@/lib/time"
import { isDirectoryItem } from "@/features/drive/driveSelectors"

// Digits and symbols group under "#"; letters of any script show as themselves.
export function nameBucket(name: string): string {
	const char = name.trim().charAt(0).toLocaleUpperCase()

	return char.length === 0 || /^[0-9!-/:-@[-`{-~]$/.test(char) ? "#" : char
}

function yearOf(timestamp: number): number {
	return new Date(timestamp < 10000000000 ? timestamp * 1000 : timestamp).getFullYear()
}

/**
 * Scrubber labels + markers for a drive listing in its current sort:
 * name → first letter (A…Z markers), date sorts → "Mar 2021" (year markers), size → the item's size.
 */
export function buildDriveFastScroll(items: readonly DriveItem[], sort: SortByType): FastScrollOptions {
	const timestampKey = sortTimestampKey(sort)
	const bySize = sort === "sizeAsc" || sort === "sizeDesc"
	const markers: FastScrollMarker[] = []
	let last: string | null = null

	if (!bySize) {
		for (let i = 0; i < items.length; i++) {
			const item = items[i]

			if (item === undefined) {
				continue
			}

			const key = timestampKey ? String(yearOf(timestampKey(item))) : nameBucket(driveItemName(item))

			if (key !== last) {
				markers.push({ label: key, index: i })

				last = key
			}
		}
	}

	return {
		markers,
		getLabel: index => {
			const item = items[index]

			if (item === undefined) {
				return null
			}

			if (bySize) {
				return isDirectoryItem(item) ? null : formatBytes(Number(item.data.size))
			}

			if (timestampKey) {
				return monthYearLabel(timestampKey(item))
			}

			return nameBucket(driveItemName(item))
		}
	}
}
