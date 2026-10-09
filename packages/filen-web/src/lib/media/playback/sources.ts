import * as Comlink from "comlink"
import type { AnyFile } from "@filen/sdk-rs"
import { sdkApi } from "@/lib/sdk/client"
import { holdPreviewStreamAlive, previewStreamUrl, releasePreviewStream } from "@/features/preview/lib/previewStream"
import type { DirectSource } from "@/lib/media/playback/directSession"
import type { RemuxSource } from "@/lib/media/playback/remuxSession"

// The app's byte sources for playback sessions: the service worker's range route and the SDK worker's
// range ports.

export function directSource(params: { file: AnyFile; name: string; contentType: string }): DirectSource {
	const { file, name, contentType } = params

	return {
		register: id => previewStreamUrl(file, name, contentType, id),
		release: releasePreviewStream,
		holdAlive: holdPreviewStreamAlive,
		reachable: async url => {
			const response = await fetch(url, { headers: { Range: "bytes=0-1" } })

			await response.body?.cancel()

			return response.status === 206
		}
	}
}

export function remuxSource(file: AnyFile): RemuxSource {
	const token = crypto.randomUUID()

	return {
		openPort: async () => {
			const { port1, port2 } = new MessageChannel()

			await sdkApi.openRangePort(Comlink.transfer(port1, [port1]), file, token)

			return port2
		},
		close: () => {
			void sdkApi.cancelPreviewDownload(token)
		}
	}
}
