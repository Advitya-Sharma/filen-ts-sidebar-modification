import { useQuery, type UseQueryResult } from "@tanstack/react-query"
import { type, type Type } from "arktype"
import { kvGetJson, kvSetJson } from "@/lib/storage/adapter"
import { kvPreference } from "@/lib/storage/preference"
import { noDiskPersister } from "@/queries/persist"
import { queryClient } from "@/queries/client"
import { DEFAULT_DENSITY_INDEX, clampDensityIndex } from "@/features/photos/lib/gridDensity"
import { probeFacts } from "@/features/videos/lib/probeFacts"
import {
	DEFAULT_GRID_PREFS,
	gridPrefsKey,
	gridPrefsSchema,
	type GridPrefs,
	type GridSource,
	type ProbeFacts
} from "@/features/videos/lib/videoGrid.logic"

// The library grids' kv-backed state: sort and filters per library view, the card density, and what the stored
// probe records say of runtimes and resolutions (all local reads).

function gridPrefsQueryKey(source: GridSource) {
	return ["videos", "gridPrefs", gridPrefsKey(source)] as const
}

export function useGridPrefsQuery(source: GridSource): UseQueryResult<GridPrefs> {
	return useQuery({
		queryKey: gridPrefsQueryKey(source),
		queryFn: async () => (await kvGetJson(gridPrefsKey(source), gridPrefsSchema)) ?? DEFAULT_GRID_PREFS,
		staleTime: Infinity,
		persister: noDiskPersister
	})
}

// Shown at once; written behind.
export function setGridPrefs(source: GridSource, prefs: GridPrefs): void {
	queryClient.setQueryData(gridPrefsQueryKey(source), prefs)
	void kvSetJson(gridPrefsKey(source), prefs).catch(() => undefined)
}

const densityIndexSchema: Type<number> = type("number.integer >= 0")

export const { get: getVideosGridDensity, set: setVideosGridDensity } = kvPreference({
	key: "videos.gridDensity.v1",
	schema: densityIndexSchema,
	fallback: DEFAULT_DENSITY_INDEX,
	normalize: clampDensityIndex
})

export function useVideosGridDensityQuery(): UseQueryResult<number> {
	return useQuery({
		queryKey: ["videos", "gridDensity"] as const,
		queryFn: getVideosGridDensity,
		persister: noDiskPersister
	})
}

// Runtime and picture height per file uuid (local reads only, see probeFacts), while a sort or filter needs them.
export function useProbeFactsQuery(enabled: boolean): UseQueryResult<ReadonlyMap<string, ProbeFacts>> {
	return useQuery({
		queryKey: ["videos", "probeFacts"] as const,
		queryFn: () => probeFacts.load(),
		enabled,
		staleTime: 0,
		persister: noDiskPersister
	})
}
