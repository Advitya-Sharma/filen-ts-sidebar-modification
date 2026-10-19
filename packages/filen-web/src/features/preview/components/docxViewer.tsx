import { useEffect, useRef, useState, type ReactNode, type RefObject } from "react"
import { useTranslation } from "react-i18next"
import { isCompoundFile } from "@filen/office/core/cfb"
import { bytesSource, type ByteSource } from "@filen/office/core/opc"
import {
	DEFAULT_MARKUP_MODE,
	markupInfo,
	type DocxEncryptionRefusal,
	type DocxHost,
	type DocxLockedDocument,
	type DocxMarkupMode,
	type DocxProtectedOpen
} from "@filen/office/dom/docx"
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
import { DOCX_FONT_URLS, docxHost, docxProtection, prewarmDocxWorker } from "@/features/preview/lib/docxEngine"
import {
	BALLOON_COLUMN_WIDTH,
	balloonsFit,
	hasBalloons,
	docxErrorKey,
	docxFailure,
	docxRefusalKey,
	docxSessionKey,
	viewOnlyReasons,
	type DocxGateReason,
	type DocxOpenContext,
	type DocxSessionKey
} from "@/features/preview/components/docxViewer.logic"
import { DocxAnnotationLayer, type DocxBalloonKey } from "@/features/preview/components/docxAnnotations"
import { DocxPreviewFallback } from "@/features/preview/components/docxPreviewFallback"
import { DocxPasswordGate, type PasswordVerdict } from "@/features/preview/components/docxPasswordPrompt"
import { DocxEditView, DocxProtectedEditView, type DocxSaveSource } from "@/features/preview/components/docxEditor"
import { DOCX_EDITING } from "@/features/preview/lib/docxEditingFlag"
import { DocxToolbar, type PasswordIntegrity } from "@/features/preview/components/docxToolbar"
import { useTrustedLinkGate } from "@/lib/trustedLinks/useTrustedLinkGate"
import { LoadingState } from "@/components/loadingState"
import { PreviewDownloadButton, PreviewErrorState, PreviewGate, PreviewLoading } from "@/features/preview/components/previewErrorState"
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

// What the viewer opens: the file as downloaded (the engine reads ranges of the cached buffer and gets a copy of
// each range it asked for, never of the whole file: the buffer stays shared with the preview cache and the
// Download action), or a password-protected file's decrypted package, held in the docx worker by its unlocked
// document (never on the page). `file`: the buffer, for docx-preview's fallback; none for a decrypted package.
interface DocxContent {
	readonly host: DocxHost
	readonly bytes: ByteSource
	readonly file: Uint8Array | null
}

function fileContent(bytes: Uint8Array): DocxContent {
	return { host: docxHost, bytes: bytesSource(bytes), file: bytes }
}

// Tracked changes are laid out as the markup mode shows them (a new layout per view).
function sourceFor(content: DocxContent, mode: DocxMarkupMode): DocxSource {
	return { host: content.host, bytes: content.bytes, options: { fontUrls: DOCX_FONT_URLS, view: markupInfo(mode).view } }
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
			{ready => (
				<DocxPasswordBoundary
					bytes={ready.bytes}
					name={alt}
				>
					{(content, key) => {
						const viewer = (
							<DocxDocumentView
								content={content}
								alt={alt}
								encrypted={key?.document.integrity}
								context={{ anon, archive, linked: isLinkedEmbedItem(item), variant }}
							/>
						)

						if (!DOCX_EDITING || !edit || editing === undefined) {
							return viewer
						}

						const editProps = {
							alt,
							writable: editing.writable,
							renamed: editing.renamed,
							onDirtyChange: editing.onDirtyChange,
							saveRef: editing.saveRef,
							onAutosave: editing.onAutosave,
							viewer: () => viewer
						}

						// DOCX_EDITING first: a build without editing drops the editor entirely. A password-protected
						// document's editor takes its own copy and seals its saves with its key.
						return key === null || content.file !== null ? (
							<DocxEditView
								bytes={content.file ?? ready.bytes}
								{...editProps}
							/>
						) : (
							<DocxProtectedEditView
								document={key.document}
								{...editProps}
							/>
						)
					}}
				</DocxPasswordBoundary>
			)}
		</PreviewGate>
	)
}

// What the file is before anything is shown: a zip opens at once; a compound file is probed for a password
// ([MS-OFFCRYPTO] §2.3.4) in the docx worker: refused (rights management, an unsupported scheme, damage), locked,
// or not encrypted at all (a legacy .doc, which the engine reports).
type Protection =
	| { readonly kind: "probing" }
	| { readonly kind: "open"; readonly content: DocxContent; readonly key: DocxSessionKey | null }
	| { readonly kind: "refused"; readonly reason: DocxEncryptionRefusal }
	| { readonly kind: "locked"; readonly locked: DocxLockedDocument }

function initial(bytes: Uint8Array): Protection {
	return isCompoundFile(bytes) ? { kind: "probing" } : { kind: "open", content: fileContent(bytes), key: null }
}

function protectionOf(bytes: Uint8Array, opened: Exclude<DocxProtectedOpen, { kind: "aborted" }>): Protection {
	switch (opened.kind) {
		case "plain":
			return { kind: "open", content: fileContent(bytes), key: null }
		case "refused":
			return { kind: "refused", reason: opened.reason }
		case "needsPassword":
			return { kind: "locked", locked: opened.locked }
	}
}

// The document behind its password, if it has one: the prompt in its place until it is given, then the viewer
// on the decrypted package where the worker holds it (or the editor on its own copy), released once the preview
// closes, whatever closes it (the overlay, a route change, a new version of the file, an error above). Reopening
// asks again.
function DocxPasswordBoundary({
	bytes,
	name,
	children
}: {
	bytes: Uint8Array
	name: string
	children: (content: DocxContent, key: DocxSessionKey | null) => ReactNode
}) {
	const { t } = useTranslation("preview")
	const [state, setState] = useState(() => ({ bytes, protection: initial(bytes) }))
	// Every document this preview unlocked and still holds: released when it unmounts, also one whose state
	// never committed.
	const [keys] = useState(() => new Set<DocxSessionKey>())
	let current = state

	if (state.bytes !== bytes) {
		current = { bytes, protection: initial(bytes) }
		setState(current)
	}

	const { protection } = current
	const probing = protection.kind === "probing"
	const key = protection.kind === "open" ? protection.key : null
	const locked = protection.kind === "locked" ? protection.locked : null

	useEffect(() => {
		if (!probing) {
			return
		}

		const controller = new AbortController()

		docxProtection.open(bytesSource(bytes), { signal: controller.signal }).then(
			opened => {
				if (opened.kind === "aborted" || controller.signal.aborted) {
					if (opened.kind === "needsPassword") {
						opened.locked.dispose()
					}

					return
				}

				const next = protectionOf(bytes, opened)

				setState(previous => (previous.bytes === bytes ? { bytes, protection: next } : previous))
			},
			(error: unknown) => {
				// Not one of the file's own verdicts (those resolve): the engine reads the file and says what it is.
				log.error("docx", "encryption probe failed", error)

				if (!controller.signal.aborted) {
					setState(previous =>
						previous.bytes === bytes
							? { bytes, protection: { kind: "open", content: fileContent(bytes), key: null } }
							: previous
					)
				}
			}
		)

		return () => {
			controller.abort()
		}
	}, [bytes, probing])

	// A locked file the prompt left without unlocking is dropped where the worker holds it, a turn later (a
	// remount in the same turn, React's Strict Mode check, keeps it).
	useEffect(() => {
		if (locked === null) {
			return
		}

		clearTimeout(disposals.get(locked))

		return () => {
			disposals.set(
				locked,
				setTimeout(() => {
					if (!unlockedLocks.has(locked)) {
						locked.dispose()
					}
				}, 0)
			)
		}
	}, [locked])

	useEffect(() => {
		if (key === null) {
			return
		}

		keys.add(key)
		key.hold()

		return () => {
			keys.delete(key)
			key.drop()
		}
	}, [key, keys])

	useEffect(
		() => () => {
			for (const held of keys) {
				held.drop()
			}
		},
		[keys]
	)

	switch (protection.kind) {
		case "probing":
			return <PreviewLoading />
		case "refused":
			return (
				<PreviewErrorState
					message={t(docxRefusalKey(protection.reason))}
					action={<PreviewDownloadButton />}
				/>
			)
		case "locked": {
			const lockedDocument = protection.locked
			const verify = async (password: string, signal: AbortSignal): Promise<PasswordVerdict> => {
				const result = await lockedDocument.unlock(password, { signal })

				if (!result.ok) {
					const { reason } = result

					if (reason === "wrongPassword") {
						return "wrong"
					}

					// The password verified, then the package failed its integrity check: no prompt helps.
					if (reason !== "aborted") {
						setState(previous => (previous.bytes === bytes ? { bytes, protection: { kind: "refused", reason } } : previous))
					}

					return "aborted"
				}

				unlockedLocks.add(lockedDocument)

				// The prompt went (closed, a new version, the preview unmounted): nothing holds the document.
				if (signal.aborted) {
					result.document.release()

					return "aborted"
				}

				const unlocked = docxSessionKey(result.document)
				const content: DocxContent = { host: result.document.host, bytes: result.document.bytes, file: null }

				keys.add(unlocked)
				setState(previous =>
					previous.bytes === bytes ? { bytes, protection: { kind: "open", content, key: unlocked } } : previous
				)

				return "accepted"
			}

			return (
				<DocxPasswordGate
					title={t("previewDocxPasswordTitle")}
					body={t("previewDocxPasswordBody", { name })}
					wrongMessage={t("previewDocxPasswordWrong")}
					label={t("previewDocxPasswordLabel")}
					submitLabel={t("previewDocxPasswordSubmit")}
					progressLabel={t("previewDocxPasswordUnlocking")}
					message={t("previewDocxNeedsPassword")}
					enterLabel={t("previewDocxEnterPassword")}
					verify={verify}
				/>
			)
		}
		case "open":
			return children(protection.content, key)
	}
}

// Locked files whose unlock succeeded: their dispose would release the unlocked document (the same hold).
const unlockedLocks = new WeakSet<DocxLockedDocument>()
// Locked files about to be dropped (their preview let go of them).
const disposals = new WeakMap<DocxLockedDocument, ReturnType<typeof setTimeout>>()

interface Opened {
	content: DocxContent
	mode: DocxMarkupMode
	source: DocxSource
}

function DocxDocumentView({
	content,
	alt,
	encrypted,
	context
}: {
	content: DocxContent
	alt: string
	// Opened with its password, and whether its integrity was verified (the badge says so).
	encrypted: PasswordIntegrity | undefined
	context: DocxOpenContext
}) {
	const { t } = useTranslation("preview")
	// A retry, or a markup mode laid out differently, opens the same content again under a fresh source (a new
	// document on the shared worker).
	const [opened, setOpened] = useState<Opened>(() => ({
		content,
		mode: DEFAULT_MARKUP_MODE,
		source: sourceFor(content, DEFAULT_MARKUP_MODE)
	}))
	let current = opened

	if (opened.content !== content) {
		current = { content, mode: DEFAULT_MARKUP_MODE, source: sourceFor(content, DEFAULT_MARKUP_MODE) }
		setOpened(current)
	}

	const document = useDocxDocument(current.source)
	const failure = docxFailure(document, content.bytes.size)
	// A decrypted package is not on the page: docx-preview has nothing to read.
	const file = failure?.fallback === true ? content.file : null
	const detail = failure?.detail

	useEffect(() => {
		if (detail !== undefined) {
			log.warn(
				"docx",
				file !== null
					? "Engine failed, showing docx-preview instead:"
					: content.file === null
						? "Engine failed on a decrypted package:"
						: "Engine failed, too large for docx-preview:",
				detail
			)
		}
	}, [detail, file, content.file])

	if (file !== null) {
		return (
			<DocxPreviewFallback
				bytes={file}
				alt={alt}
			/>
		)
	}

	if (document.status === "error") {
		return (
			<PreviewErrorState
				message={t(docxErrorKey(document.error?.reason))}
				onRetry={() => {
					setOpened({ content, mode: current.mode, source: sourceFor(content, current.mode) })
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
						: { content, mode, source: sourceFor(content, mode) }
				)
			}}
			encrypted={encrypted}
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
	encrypted,
	context
}: {
	document: DocxDocument
	alt: string
	mode: DocxMarkupMode
	onModeChange: (mode: DocxMarkupMode) => void
	encrypted: PasswordIntegrity | undefined
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
				encrypted={encrypted}
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
