import { useEffect, useState, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import type { ByteSource } from "@filen/office/core/opc"
import type { DocxHost } from "@filen/office/dom/docx"
import type { DocxRtfDocument } from "@filen/office/dom/docx/rtf"
import type { DocxEditingBackend } from "@/features/preview/lib/docxEditing"
import { docxRtf } from "@/features/preview/lib/docxEngine"
import { rtfEditingBackend } from "@/features/preview/lib/rtfEditing"
import { PreviewDownloadButton, PreviewErrorState, PreviewLoading } from "@/features/preview/components/previewErrorState"
import { log } from "@/lib/log"

// An RTF file in the docx viewer: read in the docx worker into the package the viewer opens where it is held (none of
// it on the page), then `children` with it. To be edited the editor takes its own copy and opens on it while the
// document is proved there (Gate 1: blocks the RTF writer cannot write stay locked; editing starts once it is known),
// and saves are written back as RTF there (rtfEditing.ts). Released when the preview closes.

export interface RtfOpened {
	// The viewer's source: the package where the worker holds it.
	readonly host: DocxHost
	readonly bytes: ByteSource
	// Set when the document is opened for editing: the editor's copy of the package, its proof and writer.
	readonly editing: { readonly docx: Uint8Array; readonly backend: DocxEditingBackend } | undefined
}

type State = { readonly kind: "reading" } | { readonly kind: "ready"; readonly opened: RtfOpened } | { readonly kind: "failed" }

export function RtfDocument({ bytes, edit, children }: { bytes: Uint8Array; edit: boolean; children: (opened: RtfOpened) => ReactNode }) {
	const { t } = useTranslation("preview")
	const [state, setState] = useState<{ bytes: Uint8Array; state: State }>({ bytes, state: { kind: "reading" } })
	let current = state

	if (state.bytes !== bytes) {
		current = { bytes, state: { kind: "reading" } }
		setState(current)
	}

	useEffect(() => {
		const controller = new AbortController()
		let held: DocxRtfDocument | undefined
		let proofFailed = false

		const read = async (): Promise<RtfOpened | undefined> => {
			const opened = await docxRtf.open(bytes, { signal: controller.signal })

			if (opened.kind !== "opened") {
				return undefined
			}

			const document = opened.document

			held = document

			if (controller.signal.aborted) {
				return undefined
			}

			if (!edit) {
				return { host: document.host, bytes: document.bytes, editing: undefined }
			}

			// Asked in this order: the worker hands out the copy before the proof takes it over.
			const taking = document.take()
			const locked = document.prove()

			// A failed proof fails the preview, the editor open or not (the release on close rejects it too: ignored).
			void locked.catch((error: unknown) => {
				if (controller.signal.aborted) {
					return
				}

				log.error("docx", "proving an RTF document failed", error)
				proofFailed = true
				setState(previous => (previous.bytes === bytes ? { bytes, state: { kind: "failed" } } : previous))
			})

			const docx = await taking

			return { host: document.host, bytes: document.bytes, editing: { docx, backend: rtfEditingBackend(document, locked) } }
		}

		read().then(
			opened => {
				if (controller.signal.aborted) {
					return
				}

				setState(previous =>
					previous.bytes === bytes
						? { bytes, state: opened === undefined || proofFailed ? { kind: "failed" } : { kind: "ready", opened } }
						: previous
				)
			},
			(error: unknown) => {
				log.error("docx", "reading an RTF document failed", error)

				if (!controller.signal.aborted) {
					setState(previous => (previous.bytes === bytes ? { bytes, state: { kind: "failed" } } : previous))
				}
			}
		)

		return () => {
			controller.abort()
			held?.release()
		}
	}, [bytes, edit])

	switch (current.state.kind) {
		case "reading":
			return <PreviewLoading />
		case "failed":
			return (
				<PreviewErrorState
					message={t("previewDocxLoadFailed")}
					action={<PreviewDownloadButton />}
				/>
			)
		case "ready":
			return children(current.state.opened)
	}
}
