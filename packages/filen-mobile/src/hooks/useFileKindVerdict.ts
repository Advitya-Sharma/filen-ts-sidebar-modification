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
 * Every verdict, as a reader that is a new object after each change: for a list filtered by its files' kinds,
 * which then filters again only when a verdict changed.
 */
export function useFileKindVerdicts(): FileKindVerdictReader {
	return useSyncExternalStore(
		listener => fileKindVerdicts.subscribeAny(listener),
		() => fileKindVerdicts.getReader()
	)
}

export default useFileKindVerdict
