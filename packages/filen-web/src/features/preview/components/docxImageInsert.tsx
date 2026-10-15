import { useEffect, useLayoutEffect, useRef, type DOMAttributes, type ReactNode, type RefObject } from "react"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import type { DocxEditorCommand, DocxEditorSnapshot } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxEditingHost } from "@filen/office/dom/docx/editing"
import type { DocxEditorHandle } from "@filen/office/react/docx/editor"
import type { DocxInputChannelProps, DocxSurface } from "@/features/preview/components/docxEditor.logic"
import { commandNeedsAuthor } from "@/features/preview/components/docxAuthorGate.logic"
import { INSERT_DIALOG_INPUTS } from "@/features/preview/components/docxInsert.logic"
import { DocxInsertDialogs } from "@/features/preview/components/docxInsertDialogs"
import type { ClipboardContent } from "@/features/preview/lib/docxClipboard"
import { BROWSER_PICTURE_CODECS, driveSource, fileSource } from "@/features/preview/lib/docxImages"
import {
	htmlHoldsOnlyImages,
	insertPicture,
	isPictureFile,
	pictureRefusalKey,
	preparePicture,
	type DocxPictureRefusal,
	type PictureInsert,
	type PictureRun,
	type PictureSource,
	type PreparedPicture
} from "@/features/preview/lib/docxImages.logic"
import type { DriveItem } from "@/features/drive/lib/item"
import { DriveFilePickerDialog } from "@/features/drive/components/driveFilePickerDialog"
import { log } from "@/lib/log"
import { useLatestRef } from "@/lib/useLatestRef"

export interface DocxImageInsertOptions extends DocxInputChannelProps {
	// The document's editing host (its media store) and editor, undefined while it opens.
	host: DocxEditingHost
	editor: PagedEditor | undefined
	handleRef: RefObject<DocxEditorHandle | null>
	// The toolbar's snapshot (what is enabled).
	snapshot: DocxEditorSnapshot | undefined
	readOnly: boolean
	execute: (command: DocxEditorCommand) => void
	surface: DocxSurface
}

// Put on the element around the document: pictures pasted, dragged over it and dropped.
export type DocxImageDropHandlers = Pick<DOMAttributes<HTMLElement>, "onPasteCapture" | "onDragOverCapture" | "onDropCapture">

export interface DocxImageInsertControls {
	handlers: DocxImageDropHandlers
	// Takes clipboard content the host's own Paste read when it holds a picture: true when it took it.
	paste: (content: ClipboardContent) => boolean
	// The dialogs it opens (a picture from the drive, and those of the Insert menu and the Picture group: they
	// outlive the menus that ask for them), rendered by the host.
	dialogs: ReactNode
}

// What the file chooser offers: pictures Word shows, and WebP (converted).
const PICTURE_ACCEPT = "image/*,.svg,.emf,.wmf,.tif,.tiff,.webp"

// Clipboard content goes in as its pictures when it holds no text, or only the pictures (a browser's Copy
// Image writes an <img> as HTML beside them); copied content with text keeps its own pictures (the editor pastes it).
function pastesPictures(text: string, html: string | undefined): boolean {
	return text === "" || htmlHoldsOnlyImages(html)
}

// Media waiting on the author's name: given back on the next edit (the insert ran, or the user declined and
// typed on) or when the editor closes.
interface Waiting {
	settle: () => void
	unsubscribe: () => void
}

interface TakeDeps {
	// Controllers of the pictures under way (aborted when the editor closes).
	running: Set<AbortController>
	refuse: (refusal: DocxPictureRefusal) => void
	insert: (picture: PreparedPicture) => PictureInsert
	// An animated picture went in as its first frame.
	firstFrame: () => void
}

// One picture through the pipeline; true when it went in. `outer` aborts it (the picker closed).
async function takePicture(source: PictureSource, outer: AbortSignal | undefined, deps: TakeDeps): Promise<boolean> {
	const controller = new AbortController()
	const abort = (): void => {
		controller.abort()
	}

	outer?.addEventListener("abort", abort, { once: true })
	deps.running.add(controller)

	try {
		const prepared = await preparePicture(source, BROWSER_PICTURE_CODECS, controller.signal)

		if (!prepared.ok) {
			deps.refuse(prepared.refusal)

			return false
		}

		const result = deps.insert(prepared.value)

		if (result.status === "refused") {
			deps.refuse(result.refusal)

			return false
		}

		if (prepared.value.animated === true) {
			deps.firstFrame()
		}

		return true
	} catch (error) {
		if (!controller.signal.aborted) {
			log.warn("docx", "picture not inserted", { error })
			deps.refuse("failed")
		}

		return false
	} finally {
		outer?.removeEventListener("abort", abort)
		deps.running.delete(controller)
	}
}

// Pictures into the document: from this device (Insert ▸ Picture, the file chooser), the drive (its picker),
// the clipboard (paste, the context menu's Paste) or a drop (at the point dropped on). Requests come through
// the input channel as image.insert with detail "device" or "drive".
export function useDocxImageInsert(options: DocxImageInsertOptions): DocxImageInsertControls {
	const { t } = useTranslation("preview")
	const { inputRequest, onInputDone } = options
	const latest = useLatestRef(options)
	const fileInputRef = useRef<HTMLInputElement | null>(null)
	const runningRef = useRef(new Set<AbortController>())
	const waitingRef = useRef(new Set<Waiting>())
	const picture = inputRequest?.type === "image.insert" ? inputRequest : null
	const driveNonce = picture?.detail === "drive" ? picture.nonce : null
	const deviceNonce = picture !== null && picture.detail !== "drive" ? picture.nonce : null
	const dialogRequest = inputRequest !== null && INSERT_DIALOG_INPUTS.has(inputRequest.type) ? inputRequest : null

	// The chooser opens in the commit of the click that asked for it (still a user activation).
	useLayoutEffect(() => {
		if (deviceNonce !== null) {
			onInputDone(deviceNonce)
			fileInputRef.current?.click()
		}
	}, [deviceNonce, onInputDone])

	useEffect(() => {
		const running = runningRef.current
		const waiting = waitingRef.current

		return () => {
			for (const controller of running) {
				controller.abort()
			}

			for (const entry of waiting) {
				entry.unsubscribe()
				entry.settle()
			}

			waiting.clear()
		}
	}, [])

	const editable = (): boolean => {
		const { snapshot, readOnly } = latest.current

		return !readOnly && snapshot !== undefined && !snapshot.readOnly && snapshot.enabled["image.insert"]
	}

	const run = (command: DocxEditorCommand): PictureRun => {
		const { editor, snapshot, execute } = latest.current

		if (editor === undefined || snapshot === undefined || !editable()) {
			return "refused"
		}

		const before = editor.state.doc

		execute(command)

		if (editor.state.doc !== before) {
			return "inserted"
		}

		// The author's name is asked for first: the insert runs once given, or never.
		return snapshot.review.authorNeeded || commandNeedsAuthor(command, snapshot.review) ? "waiting" : "refused"
	}

	const settleLater = (editor: PagedEditor, settle: () => void): void => {
		const doc = editor.state.doc
		const entry: Waiting = {
			settle,
			unsubscribe: editor.subscribe(() => {
				if (editor.state.doc !== doc) {
					entry.unsubscribe()
					waitingRef.current.delete(entry)
					settle()
				}
			})
		}

		waitingRef.current.add(entry)
	}

	const take = (source: PictureSource, outer?: AbortSignal): Promise<boolean> =>
		takePicture(source, outer, {
			running: runningRef.current,
			refuse: refusal => {
				toast.error(t(pictureRefusalKey(refusal)))
			},
			insert: picture => {
				const { host, editor } = latest.current
				const result = insertPicture(picture, host, run)

				if (result.status === "waiting" && editor !== undefined) {
					settleLater(editor, result.settle)
				}

				return result
			},
			firstFrame: () => {
				toast.info(t("previewDocxPictureAnimated"))
			}
		})

	const takeFiles = async (files: readonly File[]): Promise<void> => {
		for (const file of files) {
			await take(fileSource(file))
		}
	}

	const takeDriveFile = async (item: DriveItem, signal: AbortSignal): Promise<void> => {
		const taken = await take(driveSource(item), signal)

		if (taken && driveNonce !== null) {
			onInputDone(driveNonce)
		}
	}

	const handlers: DocxImageDropHandlers = {
		onDragOverCapture: event => {
			if (event.dataTransfer.types.includes("Files")) {
				event.preventDefault()
				event.dataTransfer.dropEffect = editable() ? "copy" : "none"
			}
		},
		onDropCapture: event => {
			const files = [...event.dataTransfer.files]

			if (files.length === 0) {
				return
			}

			// Never left to the browser (it would open the file in place of the app).
			event.preventDefault()
			event.stopPropagation()

			const editor = latest.current.editor

			if (editor === undefined || !editable()) {
				return
			}

			const hit = latest.current.handleRef.current?.hitTest(event.clientX, event.clientY)

			if (hit !== undefined) {
				editor.select(hit)
			}

			void takeFiles(files)
		},
		onPasteCapture: event => {
			const data = event.clipboardData
			const files = [...data.files]

			if (files.length === 0 || !pastesPictures(data.getData("text/plain"), data.getData("text/html"))) {
				return
			}

			event.preventDefault()
			event.stopPropagation()

			if (editable()) {
				void takeFiles(files)
			}
		}
	}

	return {
		handlers,
		paste: content => {
			const files = content.files ?? []

			if (files.length === 0 || !pastesPictures(content.text ?? "", content.html)) {
				return false
			}

			if (editable()) {
				void takeFiles(files)
			}

			return true
		},
		dialogs: (
			<>
				<input
					ref={fileInputRef}
					type="file"
					accept={PICTURE_ACCEPT}
					multiple
					hidden
					data-docx-picture-input
					onChange={event => {
						const files = [...(event.target.files ?? [])]

						event.target.value = ""
						void takeFiles(files)
					}}
				/>
				{dialogRequest === null || options.editor === undefined ? null : (
					<DocxInsertDialogs
						key={dialogRequest.nonce}
						request={dialogRequest}
						editor={options.editor}
						snapshot={options.snapshot}
						readOnly={options.readOnly}
						execute={options.execute}
						onDone={onInputDone}
					/>
				)}
				{driveNonce === null ? null : (
					<DriveFilePickerDialog
						title={t("previewDocxPictureDriveTitle")}
						description={t("previewDocxPictureDriveDescription")}
						accept={isPictureFile}
						onPick={takeDriveFile}
						onClose={() => {
							onInputDone(driveNonce)
						}}
					/>
				)}
			</>
		)
	}
}
