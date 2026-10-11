import { useEffect, useRef, useState } from "react"
import { useTranslation } from "react-i18next"
import { ChevronLeftIcon, ChevronRightIcon, ZoomInIcon, ZoomOutIcon } from "lucide-react"
import { bytesSource } from "@filen/office/core/opc"
import { isDocxCommandEnabled } from "@filen/office/dom/docx"
import { DocxPages, linkActivator, useDocxDocument, useDocxViewer, type DocxDocument, type DocxSource } from "@filen/office/react/docx"
import "@filen/office/css/docx.css"
import { type DriveItem } from "@/features/drive/lib/item"
import { usePreviewBytes } from "@/features/preview/hooks/usePreviewBytes"
import { DOCX_FONT_URLS, docxHost, prewarmDocxWorker } from "@/features/preview/lib/docxEngine"
import { docxErrorKey, docxFailure } from "@/features/preview/components/docxViewer.logic"
import { DocxPreviewFallback } from "@/features/preview/components/docxPreviewFallback"
import { LoadingState } from "@/components/loadingState"
import { Button } from "@/components/ui/button"
import { PreviewErrorState, PreviewGate } from "@/features/preview/components/previewErrorState"
import { isSafeAbsoluteHref } from "@/lib/safeUrl"
import { log } from "@/lib/log"
import { useLatestRef } from "@/lib/useLatestRef"

export interface DocxViewerProps {
	item: DriveItem
	alt: string
}

// Space around and between pages (px); fit-to-width leaves it on both sides.
const PAGE_GAP = 16
const WHEEL_SENSITIVITY = 0.0015

// Links are judged when clicked (never when painted): http(s) and mailto open in a new tab with no opener
// or referrer, inside the click's user activation so no popup blocker interferes; anything else is ignored.
const openLink = linkActivator({
	allow: isSafeAbsoluteHref,
	open: href => {
		window.open(href, "_blank", "noopener,noreferrer")
	}
})

// The engine reads ranges of the cached buffer and gets a copy of each range it asked for, never of the
// whole file: the buffer stays shared with the preview cache and the Download action.
function sourceFor(bytes: Uint8Array): DocxSource {
	return { host: docxHost, bytes: bytesSource(bytes), options: { fontUrls: DOCX_FONT_URLS } }
}

export function DocxViewer({ item, alt }: DocxViewerProps) {
	const result = usePreviewBytes(item)
	const downloading = result.status === "pending"

	// The worker starts while the file downloads; a buffer already cached, or a failed download, starts none.
	useEffect(() => {
		if (downloading) {
			prewarmDocxWorker()
		}
	}, [downloading])

	return (
		<PreviewGate result={result}>
			{ready => (
				<DocxDocumentView
					bytes={ready.bytes}
					alt={alt}
				/>
			)}
		</PreviewGate>
	)
}

function DocxDocumentView({ bytes, alt }: { bytes: Uint8Array; alt: string }) {
	const { t } = useTranslation("preview")
	// A retry opens the same buffer again under a fresh source (a new document on the shared worker).
	const [opened, setOpened] = useState(() => ({ bytes, source: sourceFor(bytes) }))
	let current = opened

	if (opened.bytes !== bytes) {
		current = { bytes, source: sourceFor(bytes) }
		setOpened(current)
	}

	const document = useDocxDocument(current.source)
	const failure = docxFailure(document, bytes.byteLength)
	const fallback = failure?.fallback === true
	const detail = failure?.detail

	useEffect(() => {
		if (detail !== undefined) {
			log.warn(
				"docx",
				fallback ? "Engine failed, showing docx-preview instead:" : "Engine failed, too large for docx-preview:",
				detail
			)
		}
	}, [detail, fallback])

	if (fallback) {
		return (
			<DocxPreviewFallback
				bytes={bytes}
				alt={alt}
			/>
		)
	}

	if (document.status === "error") {
		return (
			<PreviewErrorState
				message={t(docxErrorKey(document.error?.reason))}
				onRetry={() => {
					setOpened({ bytes, source: sourceFor(bytes) })
				}}
			/>
		)
	}

	return (
		<DocxPageList
			document={document}
			alt={alt}
		/>
	)
}

function DocxPageList({ document, alt }: { document: DocxDocument; alt: string }) {
	const { t } = useTranslation("preview")
	const wrapperRef = useRef<HTMLDivElement | null>(null)
	const viewer = useDocxViewer(document, { gap: PAGE_GAP })
	const { state, dispatch } = viewer
	const dispatchRef = useLatestRef(dispatch)

	// Ctrl/Cmd+wheel zooms (a native non-passive listener: React registers wheel passively, so it could
	// not stop the browser's own page zoom); a plain wheel scrolls.
	useEffect(() => {
		const wrapper = wrapperRef.current

		if (wrapper === null) {
			return
		}

		const onWheel = (event: WheelEvent): void => {
			if (!event.ctrlKey && !event.metaKey) {
				return
			}

			event.preventDefault()
			// Relative, so events faster than renders add up.
			dispatchRef.current({ type: "zoom.by", delta: -event.deltaY * WHEEL_SENSITIVITY })
		}

		wrapper.addEventListener("wheel", onWheel, { passive: false })

		return () => {
			wrapper.removeEventListener("wheel", onWheel)
		}
	}, [dispatchRef])

	return (
		<div className="flex size-full flex-col">
			<div className="flex h-10 shrink-0 items-center justify-center gap-2">
				<Button
					variant="ghost"
					size="icon-sm"
					disabled={!isDocxCommandEnabled(state, "page.previous")}
					aria-label={t("previewDocxPreviousPageAction")}
					onClick={() => {
						dispatch({ type: "page.previous" })
					}}
				>
					<ChevronLeftIcon />
				</Button>
				<span className="text-xs text-muted-foreground tabular-nums">
					{t("previewDocxPageIndicator", { current: Math.min(state.page + 1, state.pageCount), total: state.pageCount })}
				</span>
				<Button
					variant="ghost"
					size="icon-sm"
					disabled={!isDocxCommandEnabled(state, "page.next")}
					aria-label={t("previewDocxNextPageAction")}
					onClick={() => {
						dispatch({ type: "page.next" })
					}}
				>
					<ChevronRightIcon />
				</Button>
				<span
					aria-hidden="true"
					className="mx-1 h-4 w-px bg-border"
				/>
				<Button
					variant="ghost"
					size="icon-sm"
					disabled={!isDocxCommandEnabled(state, "zoom.out")}
					aria-label={t("previewDocxZoomOutAction")}
					onClick={() => {
						dispatch({ type: "zoom.out" })
					}}
				>
					<ZoomOutIcon />
				</Button>
				<span className="w-10 text-center text-xs text-muted-foreground tabular-nums">
					{t("previewDocxZoomIndicator", { percent: Math.round(state.zoom * 100) })}
				</span>
				<Button
					variant="ghost"
					size="icon-sm"
					disabled={!isDocxCommandEnabled(state, "zoom.in")}
					aria-label={t("previewDocxZoomInAction")}
					onClick={() => {
						dispatch({ type: "zoom.in" })
					}}
				>
					<ZoomInIcon />
				</Button>
			</div>
			<div
				ref={wrapperRef}
				className="relative min-h-0 flex-1"
			>
				<DocxPages
					{...viewer.pages}
					className="size-full select-text"
					label={alt}
					pageLabel={(index, total) => t("previewDocxPageIndicator", { current: index + 1, total })}
					onLinkActivate={openLink}
				/>
				{document.status === "loading" ? (
					<LoadingState
						size="lg"
						className="absolute inset-0 text-inherit"
					/>
				) : null}
			</div>
		</div>
	)
}
