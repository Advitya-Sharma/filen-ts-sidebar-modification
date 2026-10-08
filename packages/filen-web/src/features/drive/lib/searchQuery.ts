import { sdkApi } from "@/lib/sdk/client"
import { buildSearchResults, type SearchResults } from "@/features/drive/hooks/useDriveSearch.logic"
import { GRACE_MS } from "@/features/drive/lib/searchStatus.logic"

export type DriveSearchOnceResult = SearchResults & { total: bigint; settled: boolean }

// A whole-drive name search for a caller that is not the search box (the assistant). It runs on the
// one shared search engine only while the user's own search is idle, and the user's next keystroke
// supersedes it (searchEngine.ts's query); null means the engine was the user's. Hits become DriveItems
// through the same projection the search box uses.
// `signal` stops the query in the worker; it then answers with what it had.
let nextQueryId = 1

export async function searchDriveOnce(
	name: string,
	capMs: number,
	{ signal, itemType = "all" }: { signal?: AbortSignal | undefined; itemType?: "all" | "file" | "dir" } = {}
): Promise<DriveSearchOnceResult | null> {
	signal?.throwIfAborted()

	const id = nextQueryId++
	const onAbort = () => {
		void sdkApi.searchQueryCancel(id)
	}

	signal?.addEventListener("abort", onAbort, { once: true })

	try {
		const result = await sdkApi.searchQuery({ rootUuid: null, name, itemType }, { settleMs: GRACE_MS, capMs, id })

		if (result.status === "busy") {
			return null
		}

		return { ...buildSearchResults(result.snapshot.hits), total: result.snapshot.total, settled: result.settled }
	} finally {
		signal?.removeEventListener("abort", onAbort)
	}
}
