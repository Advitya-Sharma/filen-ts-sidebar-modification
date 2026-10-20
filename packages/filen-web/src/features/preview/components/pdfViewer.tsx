import { useEffect, useRef, useState } from "react"
import { useTranslation } from "react-i18next"
import { getDocument, GlobalWorkerOptions, PDFDataRangeTransport, TextLayer } from "pdfjs-dist"
import { ChevronLeftIcon, ChevronRightIcon, SearchIcon, ZoomInIcon, ZoomOutIcon } from "lucide-react"
import { isPdfCommandEnabled, type PdfSource, type PdfViewerCommand, type PdfViewerState } from "@filen/office/dom/pdf"
import {
	PdfPages,
	usePdfDocument,
	usePdfViewer,
	type PdfDocumentOptions,
	type PdfDomDocument,
	type PdfJsDom
} from "@filen/office/react/pdf"
import { type DriveItem } from "@/features/drive/lib/item"
import { usePreviewBytes } from "@/features/preview/hooks/usePreviewBytes"
import { DocxFindField } from "@/features/preview/components/docxToolbar"
import type { DocxFindControls } from "@/features/preview/components/docxViewer.logic"
import { useTrustedLinkGate } from "@/lib/trustedLinks/useTrustedLinkGate"
import { isSafeAbsoluteHref } from "@/lib/safeUrl"
import { isAnyDialogOpen } from "@/lib/keymap/dialogGuard"
import { IN_EDITORS, useAction } from "@/lib/keymap/useAction"
import { Kbd } from "@/lib/keymap/kbd"
import { LoadingState } from "@/components/loadingState"
import { Button } from "@/components/ui/button"
import { TooltipIconButton } from "@/components/ui/tooltipIconButton"
import { InputDialog } from "@/components/dialogs/inputDialog"
import { PreviewErrorState, PreviewGate, PreviewLoading } from "@/features/preview/components/previewErrorState"
import "@filen/office/css/pdf.css"
import "@/features/preview/components/pdfTheme.css"
import { recoverIfNewerBuild } from "@/lib/appUpdate"

export interface PdfViewerProps {
	item: DriveItem
	alt: string
}

// Only what the viewer uses, so the rest of pdf.js's exports stay out of the bundle.
const PDFJS: PdfJsDom = { getDocument, GlobalWorkerOptions, PDFDataRangeTransport, TextLayer }

// Vite rewrites `new URL(specifier, import.meta.url)` to a same-origin, content-hashed asset URL, which
// worker-src 'self' covers. A worker script (not one shared Worker) gives every document its own worker,
// destroyed with it: pdf.js tears down the worker of the task it destroys. pdf.js has no eval path, so the CSP
// needs nothing more.
const OPTIONS: PdfDocumentOptions = { worker: { kind: "script", src: new URL("pdfjs-dist/build/pdf.worker.mjs", import.meta.url).href } }

// Find opens for keys pressed in the layer the document is shown in (the preview overlay), never one over it.
const LAYER = "[role='dialog'], [role='alertdialog']"

const CLASS_NAMES = {
	column: "flex flex-col items-center gap-4 py-4",
	page: "flex shrink-0 items-center justify-center bg-white shadow-sm",
	canvas: (rendered: boolean) => (rendered ? "block max-w-full" : "invisible"),
	link: "rounded-xs focus-ring-row"
}

type PasswordAttempt = "initial" | "retry"

// The password prompt: the shared single-field dialog, nested in the preview overlay's own. Dismissing it leaves
// pdf.js waiting (it does not fail the task), so a button stands in to open it again.
function PdfPasswordPrompt({ attempt, onSubmit }: { attempt: PasswordAttempt; onSubmit: (password: string) => void }) {
	const { t } = useTranslation("preview")
	const [dismissed, setDismissed] = useState(false)

	return (
		<div className="flex size-full items-center justify-center px-6">
			{dismissed ? (
				<Button
					variant="outline"
					onClick={() => {
						setDismissed(false)
					}}
				>
					{t("previewPdfPasswordReopen")}
				</Button>
			) : null}
			<InputDialog
				open={!dismissed}
				pending={false}
				title={t("previewPdfPasswordTitle")}
				body={attempt === "retry" ? t("previewPdfPasswordRetryBody") : t("previewPdfPasswordBody")}
				label={t("previewPdfPasswordLabel")}
				secret
				submitLabel={t("previewPdfPasswordSubmit")}
				validate={value => value.length > 0}
				onOpenChange={open => {
					setDismissed(!open)
				}}
				onSubmit={onSubmit}
			/>
		</div>
	)
}

// The find field's view of the viewer: a query shows as searched (with its count) once its search finished.
function findControls(state: PdfViewerState, dispatch: (command: PdfViewerCommand) => void): DocxFindControls {
	return {
		query: state.find.done ? state.find.query : "",
		caseSensitive: state.find.caseSensitive,
		count: state.find.count,
		capped: state.find.capped,
		active: state.find.active,
		search: (query, caseSensitive) => {
			dispatch({ type: "find.query", query, caseSensitive })
		},
		step: direction => {
			dispatch({ type: direction > 0 ? "find.next" : "find.previous" })
		},
		close: () => {
			dispatch({ type: "find.close" })
		}
	}
}

function Navigation({ state, dispatch }: { state: PdfViewerState; dispatch: (command: PdfViewerCommand) => void }) {
	const { t } = useTranslation("preview")

	return (
		<div className="flex min-w-0 items-center gap-2 max-sm:gap-1">
			<Button
				variant="ghost"
				size="icon-sm"
				disabled={!isPdfCommandEnabled(state, "page.previous")}
				aria-label={t("previewPdfPreviousPageAction")}
				onClick={() => {
					dispatch({ type: "page.previous" })
				}}
			>
				<ChevronLeftIcon />
			</Button>
			<span className="text-xs whitespace-nowrap text-muted-foreground tabular-nums">
				{t("previewPdfPageIndicator", { current: state.page, total: state.pageCount })}
			</span>
			<Button
				variant="ghost"
				size="icon-sm"
				disabled={!isPdfCommandEnabled(state, "page.next")}
				aria-label={t("previewPdfNextPageAction")}
				onClick={() => {
					dispatch({ type: "page.next" })
				}}
			>
				<ChevronRightIcon />
			</Button>
			<span
				aria-hidden="true"
				className="mx-1 h-4 w-px bg-border max-sm:mx-0"
			/>
			<Button
				variant="ghost"
				size="icon-sm"
				disabled={!isPdfCommandEnabled(state, "zoom.out")}
				aria-label={t("previewPdfZoomOutAction")}
				onClick={() => {
					dispatch({ type: "zoom.out" })
				}}
			>
				<ZoomOutIcon />
			</Button>
			<span className="w-10 text-center text-xs text-muted-foreground tabular-nums">
				{t("previewPdfZoomIndicator", { percent: state.zoom })}
			</span>
			<Button
				variant="ghost"
				size="icon-sm"
				disabled={!isPdfCommandEnabled(state, "zoom.in")}
				aria-label={t("previewPdfZoomInAction")}
				onClick={() => {
					dispatch({ type: "zoom.in" })
				}}
			>
				<ZoomInIcon />
			</Button>
		</div>
	)
}

// The toolbar row, laid out as the DOCX viewer's: page navigation and zoom in the middle, swapped for the find
// field while finding; find on the right.
function PdfToolbar({
	state,
	dispatch,
	findFocus,
	onFindClosed
}: {
	state: PdfViewerState
	dispatch: (command: PdfViewerCommand) => void
	findFocus: number
	onFindClosed: () => void
}) {
	const { t } = useTranslation("preview")

	return (
		<div className="grid h-10 shrink-0 grid-cols-[1fr_auto_1fr] items-center gap-2 px-2">
			<div />
			{state.find.open ? (
				<DocxFindField
					find={findControls(state, dispatch)}
					focus={findFocus}
					onClosed={onFindClosed}
				/>
			) : (
				<Navigation
					state={state}
					dispatch={dispatch}
				/>
			)}
			<div className="flex min-w-0 items-center justify-end gap-1">
				{state.find.open ? null : (
					<TooltipIconButton
						label={t("previewDocxFindAction")}
						shortcut={<Kbd action="editor.find" />}
						onClick={() => {
							dispatch({ type: "find.open" })
						}}
					>
						<SearchIcon />
					</TooltipIconButton>
				)}
			</div>
		</div>
	)
}

function PdfReader({ document, alt }: { document: PdfDomDocument; alt: string }) {
	const { t } = useTranslation("preview")
	const viewer = usePdfViewer(document)
	const { state, dispatch } = viewer
	const rootRef = useRef<HTMLDivElement | null>(null)
	const [findFocus, setFindFocus] = useState(0)
	const linkGate = useTrustedLinkGate()

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
			<PdfToolbar
				state={state}
				dispatch={dispatch}
				findFocus={findFocus}
				onFindClosed={() => {
					viewer.pages.scrollerRef.current?.focus()
				}}
			/>
			{/* External links open in a new tab once their domain is trusted. */}
			<PdfPages
				pdfjs={PDFJS}
				{...viewer.pages}
				className="min-h-0 flex-1 overflow-y-auto outline-none"
				classNames={CLASS_NAMES}
				pageLabel={(page, total) => `${alt} — ${t("previewPdfPageIndicator", { current: page, total })}`}
				placeholder={
					<LoadingState
						size="sm"
						className="absolute inset-0 text-inherit"
					/>
				}
				allowLink={isSafeAbsoluteHref}
				tabIndex={-1}
				onClick={linkGate.onClick}
				onAuxClick={linkGate.onAuxClick}
			/>
			{linkGate.dialog}
		</div>
	)
}

function PdfDocument({ bytes, alt }: { bytes: Uint8Array; alt: string }) {
	const { t } = useTranslation("preview")
	const source: PdfSource = { kind: "bytes", bytes }
	const { state, password, retry } = usePdfDocument(PDFJS, source, OPTIONS)
	const failed = state.status === "error"

	useEffect(() => {
		if (failed) {
			// pdf.worker is its own file: a deploy that changed it fails documents in older tabs.
			recoverIfNewerBuild()
		}
	}, [failed])

	if (state.status === "loading") {
		return <PreviewLoading />
	}

	if (state.status === "password") {
		return (
			<PdfPasswordPrompt
				attempt={state.reason === "incorrect" ? "retry" : "initial"}
				onSubmit={password}
			/>
		)
	}

	if (state.status === "error") {
		return (
			<PreviewErrorState
				message={t("previewPdfLoadFailed")}
				onRetry={retry}
			/>
		)
	}

	return (
		<PdfReader
			document={state.document}
			alt={alt}
		/>
	)
}

// Gated on the whole-buffer download every buffered preview shares; PdfDocument owns the rest.
export function PdfViewer({ item, alt }: PdfViewerProps) {
	return (
		<PreviewGate result={usePreviewBytes(item)}>
			{ready => (
				<PdfDocument
					bytes={ready.bytes}
					alt={alt}
				/>
			)}
		</PreviewGate>
	)
}
