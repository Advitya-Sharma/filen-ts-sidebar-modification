import { sdkApi } from "@/lib/sdk/client"
import { runOp } from "@/lib/actions/outcome"
import { narrowToAnyFile } from "@/features/drive/lib/download"
import { type DriveItem } from "@/features/drive/lib/item"
import { loadPreviewBytes } from "@/features/preview/lib/previewCache"
import { previewCacheScope } from "@/features/preview/lib/accessMode"

// A signed-in drive file's whole buffer through the preview's own read (usePreviewBytes) for a caller
// without its lifecycle hook (the assistant): served from, joined with and stored in the same preview
// byte cache, and cancelled through the same token registry when `signal` aborts.
export async function loadDriveFileBytes(item: DriveItem, signal?: AbortSignal): Promise<Uint8Array> {
	const file = narrowToAnyFile(item)
	const token = crypto.randomUUID()

	function cancel(): void {
		void sdkApi.cancelPreviewDownload(token)
	}

	signal?.addEventListener("abort", cancel, { once: true })

	try {
		return await loadPreviewBytes(
			previewCacheScope("authed", undefined),
			item.data.uuid,
			Number(file.size),
			() => runOp(sdkApi.downloadFileBytes(file, token)),
			signal === undefined ? undefined : { signal }
		)
	} finally {
		signal?.removeEventListener("abort", cancel)
	}
}
