import { useEffect, useRef, useState, type RefObject } from "react"
import { useTranslation } from "react-i18next"
import { bytesSource } from "@filen/office/core/opc"
import { DEFAULT_MARKUP_MODE, markupInfo, type DocxMarkupMode } from "@filen/office/dom/docx"
import {
	DocxPages,
	indexAnnotations,
	linkActivator,
	useDocxAnnotations,
	useDocxDocument,
	useDocxViewer,
	type DocxDocument,
	type DocxSource
} from "@filen/office/react/docx"
import "@filen/office/css/docx.css"
import { isLinkedEmbedItem, type DriveItem } from "@/features/drive/lib/item"
import type { DriveVariant } from "@/features/drive/lib/preferences"
import { usePreviewBytes } from "@/features/preview/hooks/usePreviewBytes"
import { usePreviewAccessMode, usePreviewByteSource } from "@/features/preview/lib/accessMode"
import { DOCX_FONT_URLS, docxHost, prewarmDocxWorker } from "@/features/preview/lib/docxEngine"
import {
	BALLOON_COLUMN_WIDTH,
	balloonsFit,
	hasBalloons,
	docxErrorKey,
	docxFailure,
	viewOnlyReasons,
	type DocxGateReason,
	type DocxOpenContext
} from "@/features/preview/components/docxViewer.logic"
import { DocxAnnotationLayer, type DocxBalloonKey } from "@/features/preview/components/docxAnnotations"
import { DocxPreviewFallback } from "@/features/preview/components/docxPreviewFallback"
import { DocxEditView, type DocxSaveSource } from "@/features/preview/components/docxEditor"
import { DOCX_EDITING } from "@/features/preview/lib/docxEditingFlag"
import { DocxToolbar } from "@/features/preview/components/docxToolbar"
import { useTrustedLinkGate } from "@/features/chats/hooks/useTrustedLinkGate"
import { LoadingState } from "@/components/loadingState"
import { PreviewErrorState, PreviewGate } from "@/features/preview/components/previewErrorState"
import { isAnyDialogOpen } from "@/lib/keymap/dialogGuard"
import { IN_EDITORS, useAction } from "@/lib/keymap/useAction"
import { isSafeAbsoluteHref } from "@/lib/safeUrl"
import { log } from "@/lib/log"
import { useElementSize } from "@/lib/useElementSize"
import { useLatestRef } from "@/lib/useLatestRef"

export interface DocxViewerProps {
	item: DriveItem
	alt: string
	// The drive listing the file was opened from (a view-only reason outside the user's own drive).
	variant?: DriveVariant
	// The overlay's editing side for a file the user may edit (only with editing built in): the document
	// opens in the editor when the file can be saved as it mounts.
	editing?: DocxEditingProps
}

export interface DocxEditingProps {
	// Saving is possible now (false after a rename to another format, or a save that failed for good).
	writable: boolean
	renamed: boolean
	onDirtyChange: (dirty: boolean) => void
	saveRef: RefObject<DocxSaveSource | null>
	onAutosave: () => boolean
}

// Space around and between pages (px); fit-to-width leaves it on both sides.
const PAGE_GAP = 16
const WHEEL_SENSITIVITY = 0.0015
const LAYER = "[role='dialog'], [role='alertdialog']"

// The engine reads ranges of the cached buffer and gets a copy of each range it asked for, never of the
// whole file: the buffer stays shared with the preview cache and the Download action. Tracked changes are
// laid out as the markup mode shows them (a new layout per view).
function sourceFor(bytes: Uint8Array, mode: DocxMarkupMode): DocxSource {
	return { host: docxHost, bytes: bytesSource(bytes), options: { fontUrls: DOCX_FONT_URLS, view: markupInfo(mode).view } }
}

export function DocxViewer({ item, alt, variant, editing }: DocxViewerProps) {
	// Editing or viewing is decided as the file mounts. The editor keeps the document it opened across the
	// user's own saves (each makes the file a new version under a new uuid): its bytes stay the ones it read.
	const [edit] = useState(() => DOCX_EDITING && editing?.writable === true)
	const [openedItem] = useState(item)
	const result = usePreviewBytes(edit ? openedItem : item)
	const downloading = result.status === "pending"
	const anon = usePreviewAccessMode() === "anon"
	const archive = usePreviewByteSource() !== null

	// The worker starts while the file downloads; a buffer already cached, or a failed download, starts none.
	// The editor lays out on this thread and needs no worker.
	useEffect(() => {
		if (downloading && !edit) {
			prewarmDocxWorker()
		}
	}, [downloading, edit])

	return (
		<PreviewGate result={result}>
			{ready => {
				const viewer = (
					<DocxDocumentView
						bytes={ready.bytes}
						alt={alt}
						context={{ anon, archive, linked: isLinkedEmbedItem(item), variant }}
					/>
				)

				// DOCX_EDITING first: a build without editing drops the editor entirely.
				return DOCX_EDITING && edit && editing !== undefined ? (
					<DocxEditView
						bytes={ready.bytes}
						alt={alt}
						writable={editing.writable}
						renamed={editing.renamed}
						onDirtyChange={editing.onDirtyChange}
						saveRef={editing.saveRef}
						onAutosave={editing.onAutosave}
						viewer={() => viewer}
					/>
				) : (
					viewer
				)
			}}
		</PreviewGate>
	)
}

interface Opened {
	bytes: Uint8Array
	mode: DocxMarkupMode
	source: DocxSource
}

function DocxDocumentView({ bytes, alt, context }: { bytes: Uint8Array; alt: string; context: DocxOpenContext }) {
	const { t } = useTranslation("preview")
	// A retry, or a markup mode laid out differently, opens the same buffer again under a fresh source (a
	// new document on the shared worker).
	const [opened, setOpened] = useState<Opened>(() => ({
		bytes,
		mode: DEFAULT_MARKUP_MODE,
		source: sourceFor(bytes, DEFAULT_MARKUP_MODE)
	}))
	let current = opened

	if (opened.bytes !== bytes) {
		current = { bytes, mode: DEFAULT_MARKUP_MODE, source: sourceFor(bytes, DEFAULT_MARKUP_MODE) }
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
					setOpened({ bytes, mode: current.mode, source: sourceFor(bytes, current.mode) })
				}}
			/>
		)
	}

	return (
		<DocxPageList
			document={document}
			alt={alt}
			mode={current.mode}
			onModeChange={mode => {
				if (mode === current.mode) {
					return
				}

				// Modes that lay out the same view keep the open document.
				setOpened(
					markupInfo(mode).view === markupInfo(current.mode).view
						? { ...current, mode }
						: { bytes, mode, source: sourceFor(bytes, mode) }
				)
			}}
			context={context}
		/>
	)
}

// What the document says about itself once opened, kept while a markup mode opens it again (the facts do
// not change with the view).
interface Facts {
	gate: readonly DocxGateReason[]
	comments: number
	revisions: number
	formatting: number
}

function DocxPageList({
	document,
	alt,
	mode,
	onModeChange,
	context
}: {
	document: DocxDocument
	alt: string
	mode: DocxMarkupMode
	onModeChange: (mode: DocxMarkupMode) => void
	context: DocxOpenContext
}) {
	const { t } = useTranslation("preview")
	const [wrapper, setWrapper] = useState<HTMLDivElement | null>(null)
	const rootRef = useRef<HTMLDivElement | null>(null)
	const { width } = useElementSize(wrapper)
	const [facts, setFacts] = useState<Facts | null>(null)
	const info = document.info

	if (info !== undefined && facts === null) {
		setFacts({ gate: info.gate, comments: info.comments ?? 0, revisions: info.revisions ?? 0, formatting: info.formatting ?? 0 })
	}

	const reviewed = facts !== null && facts.comments + facts.revisions > 0
	const markup = markupInfo(mode)
	const annotations = useDocxAnnotations(document, reviewed)
	const index = annotations === undefined ? undefined : indexAnnotations(annotations)
	// The column is kept from the start by what the document counts, so the pages do not shrink once the
	// annotations arrive; then only where something will stand in it.
	const column =
		annotations === undefined
			? (markup.comments && (facts?.comments ?? 0) > 0) || (markup.formatting && (facts?.formatting ?? 0) > 0)
			: hasBalloons(annotations, markup)
	const aside = column && balloonsFit(width) ? BALLOON_COLUMN_WIDTH : 0
	const viewer = useDocxViewer(document, { gap: PAGE_GAP, aside })
	const { state, dispatch } = viewer
	const dispatchRef = useLatestRef(dispatch)
	const [active, setActive] = useState<DocxBalloonKey | null>(null)
	const [findFocus, setFindFocus] = useState(0)
	// The page in view when a markup mode opens the document again (a new session): shown again once the
	// new layout has it.
	const restoreRef = useRef<{ from: unknown; page: number } | null>(null)
	const session = document.session
	const linkGate = useTrustedLinkGate()
	// Links are judged when clicked (never when painted): http(s) and mailto open in a new tab with no opener
	// or referrer, inside the click's user activation so no popup blocker interferes, once a web link's
	// domain is trusted (asked once per domain); anything else is ignored.
	const openLink = linkActivator({
		allow: isSafeAbsoluteHref,
		open: href => {
			if (!linkGate.intercept(href)) {
				window.open(href, "_blank", "noopener,noreferrer")
			}
		}
	})

	useEffect(() => {
		const restore = restoreRef.current

		if (restore === null || session === undefined || restore.from === session) {
			return
		}

		if (document.status === "ready" || document.pages.length > restore.page) {
			restoreRef.current = null
			dispatch({ type: "page.goto", page: Math.min(restore.page, Math.max(0, document.pageCount - 1)) })
		}
	}, [session, document.status, document.pages.length, document.pageCount, dispatch])

	// Ctrl/Cmd+wheel zooms (a native non-passive listener: React registers wheel passively, so it could
	// not stop the browser's own page zoom); a plain wheel scrolls.
	useEffect(() => {
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
	}, [wrapper, dispatchRef])

	// Find opens for keys pressed in the layer the document is shown in (the preview overlay, or the page of
	// a public link), never for one over it.
	useAction(
		"editor.find",
		keyboardEvent => {
			const root = rootRef.current

			if (root === null) {
				return
			}

			const layer = root.closest(LAYER)
			const target = keyboardEvent.target instanceof Element ? keyboardEvent.target.closest(LAYER) : null

			if (target !== layer || (layer === null && isAnyDialogOpen())) {
				return
			}

			keyboardEvent.preventDefault()

			if (state.find.open) {
				setFindFocus(focus => focus + 1)
			} else {
				dispatch({ type: "find.open" })
			}
		},
		IN_EDITORS,
		[state.find.open, dispatch]
	)

	return (
		<div
			ref={rootRef}
			className="flex size-full flex-col"
		>
			<DocxToolbar
				state={state}
				dispatch={dispatch}
				markup={reviewed ? mode : null}
				onMarkupChange={next => {
					setActive(null)

					if (markupInfo(next).view !== markup.view) {
						restoreRef.current = { from: session, page: state.page }
					}

					onModeChange(next)
				}}
				viewOnly={facts === null ? [] : viewOnlyReasons(facts.gate, context)}
				findFocus={findFocus}
				onFindClosed={() => {
					wrapper?.querySelector<HTMLElement>(".fo-docx-pages")?.focus()
				}}
			/>
			<div
				ref={setWrapper}
				className="relative min-h-0 flex-1"
			>
				<DocxPages
					{...viewer.pages}
					className="size-full select-text"
					label={alt}
					pageLabel={(page, total) => t("previewDocxPageIndicator", { current: page + 1, total })}
					onLinkActivate={openLink}
					{...(annotations === undefined || index === undefined || !reviewed
						? {}
						: {
								overlay: frames => (
									<DocxAnnotationLayer
										annotations={annotations}
										index={index}
										markup={markup}
										frames={frames}
										pages={document.pages}
										zoom={state.zoom}
										asideWidth={aside}
										active={active}
										onActivate={setActive}
									/>
								)
							})}
				/>
				{document.status === "loading" ? (
					<LoadingState
						size="lg"
						className="absolute inset-0 text-inherit"
					/>
				) : null}
			</div>
			{linkGate.dialog}
		</div>
	)
}
