import { useEffect, useState } from "react"
import { driveItemName } from "@filen/shared"
import type { CueStore } from "@filen/shared/videos"
import type { DriveItem } from "@/features/drive/lib/item"
import { isTapTextCodec } from "@/workers/remux/textTap"
import { usePreviewAccessMode, usePreviewSiblings } from "@/features/preview/lib/accessMode"
import { cachedSiblings, loadSidecar, previewSidecars, type PreviewSubtitleFile } from "@/features/preview/lib/previewSubtitles"
import { embeddedTextContainer } from "@/features/preview/lib/previewPlayback.logic"
import type { PreviewVideo, PreviewVideoSnapshot } from "@/features/preview/lib/previewVideo"

// The subtitles a preview offers: text files beside the video (from a listing already held) and the text
// tracks inside a Matroska file, which the remux reads as it plays. Nothing is read until one is chosen.

export interface PreviewSubtitleOption {
	// "s:<uuid>" for a file beside the video, "e:<track>" for a track inside it.
	id: string
	lang: string | null
	title: string
	forced: boolean
	sdh: boolean
	embedded: boolean
}

export interface PreviewSubtitles {
	options: PreviewSubtitleOption[]
	// "off" or an option's id.
	selected: string
	select: (id: string) => void
	cues: CueStore | null
	// The chosen file could not be read or parsed.
	failed: boolean
}

// A chosen file's outcome, kept for the preview's life: switching back costs no second read.
type SidecarState = { status: "ok"; cues: CueStore } | { status: "failed" }

function sidecarOption(file: PreviewSubtitleFile): PreviewSubtitleOption {
	return { id: `s:${file.uuid}`, lang: file.lang, title: file.title, forced: file.forced, sdh: file.sdh, embedded: false }
}

export function usePreviewSubtitles(
	item: DriveItem,
	controller: PreviewVideo | null,
	snapshot: PreviewVideoSnapshot | null
): PreviewSubtitles {
	const access = usePreviewAccessMode()
	const provided = usePreviewSiblings()
	const [selected, setSelected] = useState("off")
	const [loaded, setLoaded] = useState<ReadonlyMap<string, SidecarState>>(() => new Map())
	// Looked up once per file: the listing held when the preview opened is the one beside it.
	const [sidecars] = useState(() =>
		previewSidecars(driveItemName(item), provided ?? (access === "authed" ? cachedSiblings(item.data.uuid) : []))
	)
	const probe = snapshot?.probe ?? null
	const embedded: PreviewSubtitleOption[] =
		probe === null || !embeddedTextContainer(probe)
			? []
			: probe.subs
					.filter(sub => isTapTextCodec(sub.codecId))
					.map(sub => ({
						id: `e:${String(sub.id)}`,
						lang: sub.lang,
						title: sub.name ?? "",
						forced: sub.forced,
						sdh: sub.hi,
						embedded: true
					}))
	const options = [...sidecars.map(sidecarOption), ...embedded]
	const chosenSidecar =
		selected.startsWith("s:") && !loaded.has(selected) ? sidecars.find(file => `s:${file.uuid}` === selected) : undefined

	useEffect(() => {
		if (chosenSidecar === undefined) {
			return
		}

		const abort = new AbortController()
		const id = `s:${chosenSidecar.uuid}`
		const settle = (state: SidecarState): void => {
			if (!abort.signal.aborted) {
				setLoaded(previous => new Map(previous).set(id, state))
			}
		}

		void loadSidecar(chosenSidecar, access, abort.signal).then(
			result => {
				settle(result.status === "ok" ? { status: "ok", cues: result.cues } : { status: "failed" })
			},
			() => {
				settle({ status: "failed" })
			}
		)

		return () => {
			abort.abort()
		}
	}, [chosenSidecar, access])

	function select(id: string): void {
		setSelected(id)
		controller?.setSubtitle(id === "off" ? "none" : id.startsWith("e:") ? "embedded-text" : "sidecar")
	}

	const state = loaded.get(selected)
	let cues: CueStore | null = null

	if (selected.startsWith("e:") && controller !== null && snapshot !== null) {
		cues = controller.embeddedCues(Number(selected.slice(2)), snapshot.embeddedVersion)
	} else if (state?.status === "ok") {
		cues = state.cues
	}

	return { options, selected, select, cues, failed: state?.status === "failed" }
}
