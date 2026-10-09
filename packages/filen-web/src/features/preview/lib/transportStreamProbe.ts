import type { AnyFile } from "@filen/sdk-rs"
import { sdkApi } from "@/lib/sdk/client"
import { PROBE_RANK } from "@/features/videos/lib/probeService.logic"
import { probeService } from "@/features/videos/lib/probeService"
import type { PreviewAccessMode } from "@/features/preview/lib/accessMode"

// Whether a `.ts`/`.mts` file is an MPEG transport stream, from its probe's first block or two: the drive's
// through the probe service (cached per file, shared with Videos), a public link's through the
// unauthenticated client. Null when the probe could not be read.
export async function probeTransportStream(file: AnyFile, access: PreviewAccessMode, signal: AbortSignal): Promise<boolean | null> {
	if (access === "authed") {
		const record = await probeService.request({ uuid: file.uuid, file }, "card", { rank: () => PROBE_RANK.explicit, signal })

		return record === null ? null : record.probe?.container === "ts"
	}

	const token = crypto.randomUUID()
	const stop = (): void => {
		void sdkApi.cancelPreviewDownload(token)
	}

	signal.addEventListener("abort", stop, { once: true })

	try {
		return (await sdkApi.readLinkedVideoProbeAnon(file, "card", token))?.container === "ts"
	} catch {
		return null
	} finally {
		signal.removeEventListener("abort", stop)
	}
}
