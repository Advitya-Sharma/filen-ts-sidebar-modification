import { useEffect, useRef, useState } from "react"
import { cn, driveItemName } from "@filen/shared"
import { asDirectoryOrFile, type DriveItem } from "@/features/drive/lib/item"
import { narrowToAnyFile } from "@/features/drive/lib/download"
import { BlockSource, urlReadRange } from "@/lib/media/blockSource"
import { reportMediaFailure, type MediaFailureKind } from "@/lib/media/mediaFailure"
import { openRemuxSession, type RemuxSession } from "@/lib/media/mseController"
import { linkedRemuxSource, remuxSource } from "@/lib/media/playback/sources"
import type { RemuxSource } from "@/lib/media/playback/remuxSession"
import { mediaControlsList, usePreviewAccessMode, usePreviewByteSource, usePreviewDownloadable } from "@/features/preview/lib/accessMode"
import { allowedMediaContentType } from "@/features/preview/lib/mediaType"
import { isMediaStreamAvailable } from "@/features/preview/lib/previewStream"
import { usePreviewStreamUrl } from "@/features/preview/hooks/usePreviewStreamUrl"
import { VideoPlayer } from "@/features/preview/components/videoPlayer"
import {
	MediaCannotStream,
	MediaFailureState,
	MediaFormatFailure,
	MediaPlaybackFailure
} from "@/features/preview/components/mediaFailureState"
import { mseFailureKind, mseOpenFailureKind } from "@/features/preview/lib/previewPlayback.logic"
import { PreviewLoading } from "@/features/preview/components/previewErrorState"
import { useReleaseOnUnmount } from "@/lib/media/useReleaseOnUnmount"

export interface MediaViewerProps {
	item: DriveItem
	category: "video" | "audio"
	alt: string
}

// The preview's <audio> element; `preload="metadata"` avoids streaming the whole file just to show a
// scrubber. Hidden, not unmounted, once playback failed: a fresh element would only load and fail again.
export function PreviewAudioElement({
	alt,
	src,
	hidden,
	onElement,
	onError,
	onTimeUpdate
}: {
	alt: string
	src?: string | undefined
	hidden?: boolean | undefined
	onElement?: ((audio: HTMLAudioElement | null) => void) | undefined
	onError?: ((audio: HTMLAudioElement) => void) | undefined
	onTimeUpdate?: ((audio: HTMLAudioElement) => void) | undefined
}) {
	const downloadable = usePreviewDownloadable()

	return (
		<audio
			ref={onElement}
			controls
			controlsList={mediaControlsList(downloadable)}
			preload="metadata"
			src={src}
			aria-label={alt}
			className={cn("w-full max-w-md", hidden === true && "hidden")}
			onError={
				onError === undefined
					? undefined
					: event => {
							onError(event.currentTarget)
						}
			}
			onTimeUpdate={
				onTimeUpdate === undefined
					? undefined
					: event => {
							onTimeUpdate(event.currentTarget)
						}
			}
		/>
	)
}

// Where audio that moves from one stream to another goes on from.
interface AudioResume {
	time: number
	play: boolean
}

const FROM_START: AudioResume = { time: 0, play: false }

// Streamed through the service worker's range route, as a signed-in page with a worker plays it. A
// stream that cannot be registered or breaks hands over to MediaSource (onStreamFailed); a format the
// browser cannot decode names its codecs.
function StreamedAudio({
	item,
	alt,
	contentType,
	onStreamFailed
}: {
	item: DriveItem
	alt: string
	contentType: string
	// Where the stream stood, so MediaSource goes on from there.
	onStreamFailed: (resume: AudioResume) => void
}) {
	const result = usePreviewStreamUrl(item, driveItemName(item), contentType)
	const [audio, setAudio] = useState<HTMLAudioElement | null>(null)
	const [formatFailed, setFormatFailed] = useState(false)
	const [source, setSource] = useState<BlockSource | null>(null)

	useReleaseOnUnmount(audio)

	useEffect(() => {
		if (result.status === "error") {
			onStreamFailed(FROM_START)
		}
	}, [result.status, onStreamFailed])

	if (result.status !== "success") {
		return <PreviewLoading />
	}

	return (
		<div className="relative flex size-full items-center justify-center px-6">
			<PreviewAudioElement
				alt={alt}
				src={result.url}
				hidden={formatFailed}
				onElement={setAudio}
				onError={element => {
					// Reads the file's first block to tell a format failure from a broken stream.
					const blocks = new BlockSource(Number(asDirectoryOrFile(item).data.size), urlReadRange(result.url))

					reportMediaFailure(element, blocks, kind => {
						if (kind === "other") {
							onStreamFailed({ time: element.currentTime, play: !element.paused })
						} else {
							setSource(blocks)
							setFormatFailed(true)
						}
					})
				}}
			/>
			{formatFailed && audio !== null && source !== null ? (
				<div className="absolute inset-0">
					<MediaFailureState
						kind="format"
						media={audio}
						source={source}
					/>
				</div>
			) : null}
		</div>
	)
}

// Streamed through MediaSource: the remux worker reads the file's ranges from the SDK worker (an
// unauthenticated client for a public link), so this plays with no service worker at all and never
// holds more of the file than its buffer window.
function MediaSourceAudio({ item, alt, resume: initialResume }: { item: DriveItem; alt: string; resume: AudioResume }) {
	const access = usePreviewAccessMode()
	const [audio, setAudio] = useState<HTMLAudioElement | null>(null)
	const [failure, setFailure] = useState<MediaFailureKind | null>(null)
	// Where a retry goes on from: the start, then wherever the element last really was.
	const [resume, setResume] = useState(initialResume)
	const lastPosition = useRef<AudioResume>(initialResume)
	const uuid = item.data.uuid
	const [file] = useState(() => narrowToAnyFile(item))

	useEffect(() => {
		if (audio === null) {
			return
		}

		const lifetime = new AbortController()
		// A function rather than a read: it changes across the awaits below.
		const disposed = (): boolean => lifetime.signal.aborted
		let session: RemuxSession | null = null
		let source: RemuxSource | null = null
		let portOpen = false
		const fail = (kind: MediaFailureKind): void => {
			if (!disposed()) {
				setFailure(kind)
			}
		}

		void (async () => {
			try {
				const opening = access === "anon" ? linkedRemuxSource(file) : remuxSource(file)

				source = opening

				const port = await opening.openPort()

				portOpen = true

				if (disposed()) {
					port.close()
					opening.close()

					return
				}

				const opened = await openRemuxSession({
					port,
					size: Number(file.size),
					startTime: resume.time,
					onFailure: reason => {
						fail(mseFailureKind(reason))
					}
				})

				if (disposed()) {
					opened.dispose()

					return
				}

				session = opened
				await opened.attach(audio)

				if (resume.play && !disposed()) {
					await audio.play().catch(() => undefined)
				}
			} catch (error) {
				fail(mseOpenFailureKind(error))
			}
		})()

		return () => {
			lifetime.abort()
			session?.dispose()

			if (portOpen) {
				source?.close()
			}
		}
	}, [audio, file, uuid, access, resume])

	return (
		<div className="relative flex size-full items-center justify-center px-6">
			<PreviewAudioElement
				alt={alt}
				hidden={failure !== null}
				onElement={setAudio}
				onTimeUpdate={element => {
					if (element.readyState >= HTMLMediaElement.HAVE_METADATA && !element.seeking) {
						lastPosition.current = { time: element.currentTime, play: !element.paused }
					}
				}}
			/>
			{failure === "format" ? (
				<div className="absolute inset-0">
					<MediaFormatFailure codecs={[]} />
				</div>
			) : failure === "other" ? (
				<div className="absolute inset-0">
					<MediaPlaybackFailure
						onRetry={() => {
							setFailure(null)
							setResume({ ...lastPosition.current })
						}}
					/>
				</div>
			) : null}
		</div>
	)
}

function hasMediaSource(): boolean {
	return typeof MediaSource !== "undefined" || "ManagedMediaSource" in globalThis
}

// The service worker's route first, for any file it can serve (a public link's too), so what the browser
// plays natively (AIFF, CAF, ALAC in Safari; any audio where there is no MediaSource) plays as it is; a
// broken or missing route moves to MediaSource where there is one, else says the browser cannot stream it.
function PreviewAudio({ item, alt }: { item: DriveItem; alt: string }) {
	const contentType = allowedMediaContentType(item)
	const [resume, setResume] = useState<AudioResume | null>(() => (contentType !== null && isMediaStreamAvailable() ? null : FROM_START))

	if (resume === null && contentType !== null) {
		return (
			<StreamedAudio
				item={item}
				alt={alt}
				contentType={contentType}
				onStreamFailed={next => {
					setResume(previous => previous ?? next)
				}}
			/>
		)
	}

	if (!hasMediaSource()) {
		return <MediaCannotStream />
	}

	return (
		<MediaSourceAudio
			item={item}
			alt={alt}
			resume={resume ?? FROM_START}
		/>
	)
}

// Video and audio always stream, never as a whole-file buffer: through the service worker's range route
// where it can serve the file, else through MediaSource fed by the remux worker.
export function MediaViewer({ item, category, alt }: MediaViewerProps) {
	const downloadable = usePreviewDownloadable()
	// Only an archive's entry carries a byte source, and archives never open media in the preview.
	const byteSource = usePreviewByteSource()

	if (byteSource !== null) {
		return <MediaFormatFailure codecs={[]} />
	}

	if (category === "video") {
		return (
			<VideoPlayer
				key={item.data.uuid}
				item={item}
				alt={alt}
				downloadable={downloadable}
				positionKey={item.data.uuid}
			/>
		)
	}

	return (
		<PreviewAudio
			key={item.data.uuid}
			item={item}
			alt={alt}
		/>
	)
}
