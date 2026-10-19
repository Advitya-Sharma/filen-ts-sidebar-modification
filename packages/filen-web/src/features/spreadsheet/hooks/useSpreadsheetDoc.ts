import { useEffect, useState } from "react"
import * as Comlink from "comlink"
import { driveItemName } from "@filen/shared"
import { type DriveItem } from "@/features/drive/lib/item"
import { extensionOf, itemTypeExtension } from "@/features/drive/lib/preview.logic"
import { usePreviewBytes } from "@/features/preview/hooks/usePreviewBytes"
import { gridDoc, type GridDoc } from "@/features/spreadsheet/lib/cellStore.logic"
import type { SpreadsheetDoc } from "@/features/spreadsheet/lib/model"
import { closeSpreadsheet, openSpreadsheet, sniffSpreadsheetKind, type SpreadsheetLock } from "@/features/spreadsheet/lib/spreadsheetClient"
import type { WorkbookRefusal } from "@/features/spreadsheet/lib/xlsxEncryption"
import type { PasswordVerdict } from "@/features/preview/components/docxPasswordPrompt"
import { spreadsheetFileKind, spreadsheetSaveFormat } from "@/features/spreadsheet/lib/fileKind"
import { type ErrorDTO } from "@/lib/sdk/errors"

export type SpreadsheetDocState =
	| { status: "pending" }
	| { status: "error"; dto: ErrorDTO; retry: () => void }
	| { status: "unreadable" }
	// A password-protected workbook, refused without asking (rights-managed, an encryption Filen can't open,
	// or damaged).
	| { status: "refused"; reason: WorkbookRefusal }
	// A password-protected workbook waiting for its password: `verify` tries one (passed on to the worker and
	// kept nowhere), for the password prompt.
	| { status: "locked"; verify: (password: string, signal: AbortSignal) => Promise<PasswordVerdict> }
	// `unnamed`: the file's name does not say it is a spreadsheet, so its kind was read from its bytes and
	// it opens read-only (saving could rewrite it as the wrong format). `renamed`: the file has since been
	// renamed to another format's extension, which a save would mislabel, so it is read-only too.
	// `bytes`: the file as opened (the preview cache's own buffer: read it, never transfer it). `doc`: the
	// view as it now stands, edits folded in through `updateDoc` (the one copy held: the view as opened is
	// let go once edits move past it). `encrypted`: a password-protected workbook, opened decrypted in the
	// worker (`bytes` is the sealed file: the worker proves it from its plaintext, and seals its saves).
	| {
			status: "ready"
			id: number
			doc: GridDoc
			updateDoc: (update: (doc: GridDoc) => GridDoc) => void
			unnamed: boolean
			renamed: boolean
			bytes: Uint8Array
			encrypted: "verified" | "unchecked" | undefined
	  }

type Opened =
	| { status: "unreadable" }
	| { status: "refused"; reason: WorkbookRefusal }
	| { status: "locked"; verify: (password: string, signal: AbortSignal) => Promise<PasswordVerdict> }
	| {
			status: "ready"
			id: number
			doc: GridDoc
			unnamed: boolean
			format: string | null
			encrypted: "verified" | "unchecked" | undefined
	  }

function extension(item: DriveItem): string {
	return extensionOf(driveItemName(item))
}

function nameFormat(item: DriveItem): string | null {
	return spreadsheetSaveFormat(extension(item))
}

// Downloads the file (the preview's shared byte load and cache) and opens it in the spreadsheet worker,
// closing it there again when the viewer goes. The worker gets a copy: the preview cache keeps the bytes.
// Opens once per `documentKey`: an item that changes under the same key (an own save rotating its uuid)
// keeps the open document, with its undo history.
export function useSpreadsheetDoc(item: DriveItem, documentKey: string): SpreadsheetDocState {
	const [pinned, setPinned] = useState({ key: documentKey, item })

	if (pinned.key !== documentKey) {
		setPinned({ key: documentKey, item })
	}

	const opening = pinned.key === documentKey ? pinned.item : item
	const bytes = usePreviewBytes(opening)
	const named = spreadsheetFileKind(extension(opening))
	// An unnamed file's bytes tell xlsx, xls and delimited text apart; only its stored type can say tab.
	const unnamedTab = spreadsheetFileKind(itemTypeExtension(opening)) === "tsv"
	const format = nameFormat(opening)
	const [opened, setOpened] = useState<{ bytes: Uint8Array; state: Opened } | null>(null)
	const source = bytes.status === "success" ? bytes.bytes : null

	useEffect(() => {
		if (source === null) {
			return undefined
		}

		let live = true
		let id: number | null = null
		let lock: SpreadsheetLock | null = null
		const sniffed = sniffSpreadsheetKind(source)
		const kind = named ?? (unnamedTab && sniffed === "csv" ? "tsv" : sniffed)
		const copy = source.slice()

		const file = source

		function show(state: Opened): void {
			setOpened({ bytes: file, state })
		}

		function ready(result: { id: number; doc: SpreadsheetDoc; encrypted?: "verified" | "unchecked" }): void {
			id = result.id
			show({
				status: "ready",
				id: result.id,
				doc: gridDoc(result.doc),
				unnamed: named === null,
				format,
				encrypted: result.encrypted
			})
		}

		// The password prompt's check (DocxPasswordPrompt's `verify`): `signal` stops the try (the user's
		// Cancel), leaving the workbook locked for another. A password that turns out right as the try is
		// cancelled still opens it.
		async function verify(password: string, signal: AbortSignal): Promise<PasswordVerdict> {
			const held = lock

			if (held === null) {
				return "aborted"
			}

			const cancel = () => {
				held.cancel()
			}

			signal.addEventListener("abort", cancel, { once: true })

			// A try that rejects (the decrypted workbook does not open, the worker died) gives the lock up.
			const answer = await held.unlock(password).catch(() => null)

			signal.removeEventListener("abort", cancel)

			// Gone meanwhile: the lock was discarded, which closes whatever the try opened.
			if (!live) {
				return "aborted"
			}

			if (answer === null) {
				held.discard()
				lock = null
				show({ status: "unreadable" })

				return "aborted"
			}

			if ("id" in answer) {
				lock = null
				ready(answer)

				return "accepted"
			}

			if ("wrongPassword" in answer) {
				return "wrong"
			}

			if ("refused" in answer) {
				lock = null
				show({ status: "refused", reason: answer.refused })
			}

			return "aborted"
		}

		// Stops a protected workbook's default-password try if the viewer goes first.
		const leaving = new AbortController()

		openSpreadsheet(Comlink.transfer(copy, [copy.buffer]), kind, leaving.signal)
			.then(result => {
				if (!live) {
					if ("id" in result) {
						closeSpreadsheet(result.id)
					} else if ("locked" in result) {
						result.locked.discard()
					}

					return
				}

				if ("id" in result) {
					ready(result)
				} else if ("locked" in result) {
					lock = result.locked
					show({ status: "locked", verify })
				} else {
					show({ status: "refused", reason: result.refused })
				}
			})
			.catch(() => {
				if (live) {
					show({ status: "unreadable" })
				}
			})

		return () => {
			live = false
			leaving.abort()
			lock?.discard()

			if (id !== null) {
				closeSpreadsheet(id)
			}
		}
	}, [source, named, unnamedTab, format])

	if (bytes.status === "error") {
		return { status: "error", dto: bytes.dto, retry: bytes.refetch }
	}

	// A state from an earlier buffer is not this one's.
	if (opened?.bytes !== source) {
		return { status: "pending" }
	}

	if (opened.state.status !== "ready") {
		return opened.state
	}

	const { id, doc, unnamed, format: openedAs, encrypted } = opened.state

	function updateDoc(update: (doc: GridDoc) => GridDoc): void {
		setOpened(prev => (prev?.state.status === "ready" ? { ...prev, state: { ...prev.state, doc: update(prev.state.doc) } } : prev))
	}

	return {
		status: "ready",
		id,
		doc,
		updateDoc,
		unnamed,
		renamed: !unnamed && nameFormat(item) !== openedAs,
		bytes: opened.bytes,
		encrypted
	}
}
