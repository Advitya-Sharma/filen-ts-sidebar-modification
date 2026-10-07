import { sdkApi } from "@/lib/sdk/client"
import { buildSearchResults, type SearchResults } from "@/features/drive/hooks/useDriveSearch.logic"
import { GRACE_MS } from "@/features/drive/lib/searchStatus.logic"

export type DriveSearchOnceResult = SearchResults & { total: bigint; settled: boolean }

// A whole-drive name search for a caller that is not the search box (the assistant). It runs on the
// one shared search engine only while the user's own search is idle, and the user's next keystroke
// supersedes it (searchEngine.ts's query); null means the engine was the user's. Hits become DriveItems
// through the same projection the search box uses.
export async function searchDriveOnce(name: string, capMs: number): Promise<DriveSearchOnceResult | null> {
	const result = await sdkApi.searchQuery({ rootUuid: null, name }, { settleMs: GRACE_MS, capMs })

	if (result.status === "busy") {
		return null
	}

	return { ...buildSearchResults(result.snapshot.hits), total: result.snapshot.total, settled: result.settled }
}
