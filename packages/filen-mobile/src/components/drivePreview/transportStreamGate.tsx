import { useEffect } from "react"
import { type UseQueryResult } from "@tanstack/react-query"
import { useRecyclingState } from "@shopify/flash-list"
import useFileKindVerdict from "@/hooks/useFileKindVerdict"
import useDrivePreviewStore from "@/stores/useDrivePreview.store"
import { recordContentVerdict, verdictFileOf } from "@/lib/previewType"
import { readRangeFromUrl, sniffTransportStream, transportStreamOutcome, type TransportStreamSniffer } from "@/lib/transportStreamSniff"
import UnavailableOfflineNotice from "@/components/drivePreview/unavailableOfflineNotice"
import PreviewLoadFailedNotice from "@/components/drivePreview/previewLoadFailedNotice"
import { PreviewSpinner } from "@/components/drivePreview/previewStatus"
import type { DriveItemFileExtracted } from "@/types"

// Holds a `.ts`/`.mts` that may be a transport stream, or an `.m2ts` not yet found playable, until a read of its
// first bytes says what it is, then pins that to its page (pinVerdict), which picks the renderer: the video
// player for a stream the platform plays, no preview for one it does not (transportStreamOutcome), the text
// viewer for TypeScript. A verdict is kept in the verdict store, an unplayable stream for the session
// (pinVerdict), so a later open pins either at once, with no read. The read is one Range request against the
// HTTP provider, or the local copy's head.
const TransportStreamGate = ({
	item,
	fileUrlQuery,
	sniff = sniffTransportStream
}: {
	item: DriveItemFileExtracted
	fileUrlQuery: UseQueryResult<string | null, Error>
	sniff?: TransportStreamSniffer
}) => {
	const uuid = item.data.uuid
	const verdict = useFileKindVerdict(uuid)
	const [attempt, setAttempt] = useRecyclingState<number>(0, [uuid])
	const [failedAttempt, setFailedAttempt] = useRecyclingState<number | null>(null, [uuid])
	const fileUrl = fileUrlQuery.status === "success" ? fileUrlQuery.data : null

	useEffect(() => {
		if (verdict !== undefined) {
			useDrivePreviewStore.getState().pinVerdict(uuid, verdict)

			return
		}

		if (fileUrl === null) {
			return
		}

		const abort = new AbortController()

		sniff(readRangeFromUrl(fileUrl, abort.signal), Number(item.data.size)).then(layout => {
			if (abort.signal.aborted) {
				return
			}

			if (layout === null) {
				setFailedAttempt(attempt)

				return
			}

			const found = transportStreamOutcome(layout)

			if (found !== "unplayable") {
				recordContentVerdict(verdictFileOf(item), found)
			}

			// Pinned from the read itself, whether or not the store kept it.
			useDrivePreviewStore.getState().pinVerdict(uuid, found)
		})

		return () => {
			abort.abort()
		}
	}, [verdict, fileUrl, uuid, item, attempt, sniff, setFailedAttempt])

	if (fileUrlQuery.status === "error") {
		return (
			<PreviewLoadFailedNotice
				className="bg-transparent"
				onRetry={() => fileUrlQuery.refetch()}
			/>
		)
	}

	// Offline, and the device holds no copy to read.
	if (fileUrlQuery.status === "success" && fileUrl === null) {
		return <UnavailableOfflineNotice />
	}

	if (failedAttempt === attempt) {
		return (
			<PreviewLoadFailedNotice
				className="bg-transparent"
				onRetry={() => setAttempt(previous => previous + 1)}
			/>
		)
	}

	return <PreviewSpinner className="bg-transparent" />
}

export default TransportStreamGate
