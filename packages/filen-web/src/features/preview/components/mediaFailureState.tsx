import { useTranslation } from "react-i18next"
import { errorLabel } from "@/lib/i18n/errorLabel"
import { mediaFailureDTO } from "@/lib/media/mediaFailure"
import { CODECS, type CodecId } from "@/features/preview/lib/containerTracks"
import { PreviewDownloadButton, PreviewErrorState } from "@/features/preview/components/previewErrorState"

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
