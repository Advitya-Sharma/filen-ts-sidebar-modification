import { useEffect, useState } from "react"
import { useTranslation } from "react-i18next"
import { errorLabel } from "@/lib/i18n/errorLabel"
import { mediaFailureDTO, type MediaFailureKind } from "@/lib/media/mediaFailure"
import type { BlockSource } from "@/lib/media/blockSource"
import { browserCanPlay, CODECS, readContainerTracks, unplayableTracks, type CodecId } from "@/features/preview/lib/containerTracks"
import { PreviewDownloadButton, PreviewErrorState, PreviewLoading } from "@/features/preview/components/previewErrorState"

// The codecs in `source`'s container header that `media` reports it cannot decode; null while reading.
// A file whose header cannot be read, or that names nothing unplayable, resolves empty.
function useUnplayableCodecs(media: HTMLMediaElement, source: BlockSource): CodecId[] | null {
	const [codecs, setCodecs] = useState<CodecId[] | null>(null)

	useEffect(() => {
		let live = true

		void readContainerTracks(source)
			.then(
				tracks => (tracks === null ? [] : unplayableTracks(tracks, ["audio", "video"], browserCanPlay(media))),
				() => []
			)
			.then(found => {
				if (live) {
					setCodecs(found)
				}
			})

		return () => {
			live = false
		}
	}, [media, source])

	return codecs
}

// A format failure: names the codecs it can and offers the host's Download, since no retry plays it.
export function MediaFormatFailure({ codecs }: { codecs: readonly CodecId[] }) {
	const { t } = useTranslation("preview")

	return (
		<PreviewErrorState
			message={
				codecs.length > 0
					? t("previewMediaFormatUnsupportedCodecs", { codecs: codecs.map(codec => CODECS[codec].label).join(", ") })
					: errorLabel(mediaFailureDTO("format"))
			}
			action={<PreviewDownloadButton />}
		/>
	)
}

function FormatFailure({ media, source }: { media: HTMLMediaElement; source: BlockSource }) {
	const codecs = useUnplayableCodecs(media, source)

	if (codecs === null) {
		return <PreviewLoading />
	}

	return <MediaFormatFailure codecs={codecs} />
}

// Nothing in this browser can stream the file (no service worker route and no MediaSource): says so, and
// offers the host's Download, whatever the file holds.
export function MediaCannotStream() {
	const { t } = useTranslation("preview")

	return (
		<PreviewErrorState
			message={t("previewMediaCannotStream")}
			action={<PreviewDownloadButton />}
		/>
	)
}

// A stream that could not be read, with nothing further to fall back to; Retry plays it again.
export function MediaPlaybackFailure({ onRetry }: { onRetry?: (() => void) | undefined }) {
	return (
		<PreviewErrorState
			message={errorLabel(mediaFailureDTO("other"))}
			{...(onRetry === undefined ? {} : { onRetry })}
		/>
	)
}

// What a preview's audio shows once its element gave up: a format failure names the codecs the file's
// header declares that the element cannot decode; any other failure offers Retry where there is one.
export function MediaFailureState({
	kind,
	media,
	source,
	onRetry
}: {
	kind: MediaFailureKind
	media: HTMLMediaElement
	source: BlockSource
	onRetry?: (() => void) | undefined
}) {
	if (kind === "format") {
		return (
			<FormatFailure
				media={media}
				source={source}
			/>
		)
	}

	return <MediaPlaybackFailure onRetry={onRetry} />
}
