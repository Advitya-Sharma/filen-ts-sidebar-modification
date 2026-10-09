import * as Comlink from "comlink"
import type { AnyFile } from "@filen/sdk-rs"
import { sdkApi } from "@/lib/sdk/client"
import { holdPreviewStreamAlive, previewStreamUrl, releasePreviewStream } from "@/features/preview/lib/previewStream"
import type { DirectSource } from "@/lib/media/playback/directSession"
import type { RemuxSource } from "@/lib/media/playback/remuxSession"
import { warmRemuxFile, type RemuxWarmth } from "@/lib/media/mseController"
import type { VideoProbe } from "@/lib/media/probe/types"

// The app's byte sources for playback sessions: the service worker's range route and the SDK worker's
// range ports.

export function directSource(params: { file: AnyFile; name: string; contentType: string }): DirectSource {
	const { file, name, contentType } = params
	const presetId = takeWarmDirect(file.uuid)

	return {
		register: id => previewStreamUrl(file, name, contentType, id),
		release: releasePreviewStream,
		holdAlive: holdPreviewStreamAlive,
		reachable: async url => {
			const response = await fetch(url, { headers: { Range: "bytes=0-1" } })

			await response.body?.cancel()

			return response.status === 206
		},
		...(presetId === null ? {} : { presetId })
	}
}

function createRemuxSource(file: AnyFile, warmKey?: string): RemuxSource {
	const token = crypto.randomUUID()

	return {
		openPort: async () => {
			const { port1, port2 } = new MessageChannel()

			await sdkApi.openRangePort(Comlink.transfer(port1, [port1]), file, token)

			return port2
		},
		close: () => {
			void sdkApi.cancelPreviewDownload(token)
		},
		...(warmKey === undefined ? {} : { warmKey })
	}
}

export function remuxSource(file: AnyFile): RemuxSource {
	return takeWarmRemux(file.uuid) ?? createRemuxSource(file)
}

// ── Prewarm ─────────────────────────────────────────────────────────────────
// The next item, readied while the current one ends (Up Next shows): a Matroska file is opened in the remux
// worker (its head, index and first keyframe interval read into the cache the session then takes over); a
// file played as it is gets its service worker registration up front. One at a time, once per file.

type Warm =
	| { kind: "remux"; fileUuid: string; source: RemuxSource; warmth: Promise<RemuxWarmth> }
	| { kind: "direct"; fileUuid: string; id: string; registered: Promise<unknown> }

let warm: Warm | null = null

const DIRECT_TYPES: Readonly<Record<string, string>> = { mp4: "video/mp4", mov: "video/mp4", webm: "video/webm" }

export function prewarmPlayback(params: { file: AnyFile; name: string; probe: VideoProbe; idrOnlyStart: boolean }): void {
	const { file, name, probe } = params

	if (warm?.fileUuid === file.uuid) {
		return
	}

	cancelPrewarm()

	if (probe.container === "mkv") {
		const key = crypto.randomUUID()
		const source = createRemuxSource(file, key)

		warm = {
			kind: "remux",
			fileUuid: file.uuid,
			source,
			warmth: source.openPort().then(port =>
				warmRemuxFile({
					port,
					size: Number(file.size),
					key,
					idrOnlyStart: params.idrOnlyStart,
					...(probe.index === null ? {} : { index: probe.index })
				})
			)
		}
		warm.warmth.catch(() => {
			cancelPrewarm(file.uuid)
		})

		return
	}

	const contentType = DIRECT_TYPES[probe.container]

	if (contentType !== undefined) {
		const id = crypto.randomUUID()

		warm = { kind: "direct", fileUuid: file.uuid, id, registered: previewStreamUrl(file, name, contentType, id).catch(() => undefined) }
	}
}

// Lets go of what was readied and not taken (Watch credits, a seek back, the player closing).
export function cancelPrewarm(fileUuid?: string): void {
	const current = warm

	if (current === null || (fileUuid !== undefined && current.fileUuid !== fileUuid)) {
		return
	}

	warm = null

	if (current.kind === "direct") {
		releasePreviewStream(current.id)

		return
	}

	current.source.close()
	void current.warmth.then(
		warmth => {
			warmth.cancel()
		},
		() => undefined
	)
}

function takeWarmRemux(fileUuid: string): RemuxSource | null {
	const current = warm

	if (current?.kind !== "remux" || current.fileUuid !== fileUuid) {
		return null
	}

	warm = null
	void current.warmth.then(
		warmth => {
			warmth.handOver()
		},
		() => undefined
	)

	return current.source
}

function takeWarmDirect(fileUuid: string): string | null {
	const current = warm

	if (current?.kind !== "direct" || current.fileUuid !== fileUuid) {
		return null
	}

	warm = null

	return current.id
}
