import { useSyncExternalStore } from "react"
import { useQueries } from "@tanstack/react-query"
import { useTranslation } from "react-i18next"
import { activeLibraryRoots } from "@filen/shared/videos"
import { useLibraryConfigQuery, useVideoRootChainQuery } from "@/features/videos/queries/library"
import { videoRootChainQueryOptions, type VideoRootChain } from "@/features/videos/lib/library"
import { videoLibrary } from "@/features/videos/lib/index"
import { libraryRootViews, type VideoLibraryRootView } from "@/features/videos/lib/libraryView.logic"

// The configuration and the index snapshot, joined per directory. Reads only: the Videos layout is what opens
// the library (useVideoLibrary).
export function useVideoLibraryView() {
	const configQuery = useLibraryConfigQuery()
	const snapshot = useSyncExternalStore(videoLibrary.subscribe, videoLibrary.getSnapshot)
	const roots = configQuery.data === undefined ? [] : libraryRootViews(activeLibraryRoots(configQuery.data.doc), snapshot.roots)

	return {
		configQuery,
		snapshot,
		roots
	}
}

function chainName(chain: VideoRootChain | undefined, driveName: string): string {
	if (chain === undefined || chain === null) {
		return ""
	}

	return chain.length <= 1 ? driveName : (chain.at(-1)?.name ?? "")
}

function knownName(name: string | undefined): name is string {
	return name !== undefined && name.length > 0
}

// A library directory's name: the index's, else (before its first scan, or turned off) the last link of its
// chain, which usually costs no request. The drive root reads as My Drive.
export function useVideoRootName(uuid: string, indexName: string | undefined): string {
	const { t } = useTranslation("drive")
	const known = knownName(indexName)
	const chainQuery = useVideoRootChainQuery(uuid, !known && uuid.length > 0)

	return known ? indexName : chainName(chainQuery.data, t("driveMyDrive"))
}

// useVideoRootName for several directories at once, by uuid.
export function useVideoRootNames(roots: readonly VideoLibraryRootView[]): Map<string, string> {
	const { t } = useTranslation("drive")
	const unnamed = roots.filter(root => !knownName(root.status?.name))
	const chains = useQueries({
		queries: unnamed.map(root => videoRootChainQueryOptions(root.uuid))
	})
	const names = new Map<string, string>()

	for (const root of roots) {
		const indexName = root.status?.name

		names.set(root.uuid, knownName(indexName) ? indexName : chainName(chains[unnamed.indexOf(root)]?.data, t("driveMyDrive")))
	}

	return names
}
