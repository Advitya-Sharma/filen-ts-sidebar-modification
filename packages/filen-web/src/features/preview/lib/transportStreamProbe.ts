import type { AnyFile } from "@filen/sdk-rs"
import type { VideoProbe } from "@filen/shared/videos"
import { sdkApi } from "@/lib/sdk/client"
import { PROBE_RANK } from "@/features/videos/lib/probeService.logic"
import { probeService } from "@/features/videos/lib/probeService"
import type { PreviewAccessMode } from "@/features/preview/lib/accessMode"

// A public link's transport stream probe, kept for the player that opens on it next: nothing caches a public
// link's probe, and the player's ladder would otherwise read the same head again. One slot, taken once.
let linkedProbe: { uuid: string; probe: VideoProbe } | null = null

// The probe the gate read for the public link file `uuid`, once.
export function takeLinkedTransportStreamProbe(uuid: string): VideoProbe | null {
	if (linkedProbe?.uuid !== uuid) {
		return null
	}

	const { probe } = linkedProbe

	linkedProbe = null

	return probe
}

// Whether a `.ts`/`.mts` file is an MPEG transport stream, from its probe: the drive's through the probe
// service (cached per file, shared with Videos), a public link's through the unauthenticated client. Read at
// the tier the player's ladder needs, so a stream's probe is read once: anything else is decided by its first
// block whatever the tier. Null when the probe could not be read.
export async function probeTransportStream(file: AnyFile, access: PreviewAccessMode, signal: AbortSignal): Promise<boolean | null> {
	if (access === "authed") {
		const record = await probeService.request({ uuid: file.uuid, file }, "full", { rank: () => PROBE_RANK.explicit, signal })

		return record === null ? null : record.probe?.container === "ts"
	}

	const token = crypto.randomUUID()
	const stop = (): void => {
		void sdkApi.cancelPreviewDownload(token)
	}

	signal.addEventListener("abort", stop, { once: true })

	try {
		const probe = await sdkApi.readLinkedVideoProbeAnon(file, "full", token)

		if (probe?.container !== "ts") {
			return false
		}

		linkedProbe = { uuid: file.uuid, probe }

		return true
	} catch {
		return null
	} finally {
		signal.removeEventListener("abort", stop)
	}
}
