import { useSyncExternalStore } from "react"
import { type ContentVerdict } from "@filen/shared"
import fileKindVerdicts, { type FileKindVerdictReader } from "@/lib/fileKindVerdicts"

function noSubscription(): () => void {
	return () => {}
}

/**
 * A file's content verdict (lib/fileKindVerdicts.ts), re-rendering the caller only when that file's verdict
 * changes. Undefined for none, or no uuid (a directory). Pass it on to the previewType.ts classifiers so what
 * the caller computes from it follows the change.
 */
export function useFileKindVerdict(uuid: string | null): ContentVerdict | undefined {
	const subscribe = uuid === null ? noSubscription : (listener: () => void) => fileKindVerdicts.subscribeFile(uuid, listener)

	return useSyncExternalStore(subscribe, () => (uuid === null ? undefined : fileKindVerdicts.get(uuid)))
}

/**
 * Every verdict, as a reader that is a new object only after a change `affects` says touches the caller's files
 * (fileKindVerdicts.watch): for a list filtered by its files' kinds, which then filters again only for its own.
 */
export function useFileKindVerdicts(affects: (changed: readonly string[]) => boolean): FileKindVerdictReader {
	const watch = fileKindVerdicts.watch(affects)

	return useSyncExternalStore(watch.subscribe, watch.getSnapshot)
}

export default useFileKindVerdict
