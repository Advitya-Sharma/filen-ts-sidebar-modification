import { useSyncExternalStore } from "react"
import type { ContentVerdict } from "@filen/shared"
import { fileKindVerdict, subscribeFileKindVerdicts } from "@/features/drive/lib/fileKindVerdicts"

// A file's content verdict (fileKindVerdicts.ts), re-rendering the caller only when that file's verdict
// changes. Undefined for none, or no uuid (a directory).
export function useFileKindVerdict(uuid: string | undefined): ContentVerdict | undefined {
	return useSyncExternalStore(subscribeFileKindVerdicts, () => (uuid === undefined ? undefined : fileKindVerdict(uuid)))
}
