import * as Comlink from "comlink"
import { driveItemName } from "@filen/shared"
import { currentRootUuid } from "@/features/drive/lib/actions"
import { asDirectoryOrFile, type DriveItem } from "@/features/drive/lib/item"
import { runPreviewSave } from "@/features/drive/lib/previewSave.logic"
import { driveListingQueryUpdate, normalizeParentUuid } from "@/features/drive/queries/drive"
import { withSpreadsheetWorker } from "@/features/spreadsheet/lib/spreadsheetClient"
import { xlsxCopyName } from "@/features/spreadsheet/lib/xlsConvert"
import { sdkApi } from "@/lib/sdk/client"
import type { ErrorDTO } from "@/lib/sdk/errors"

export type SaveAsXlsxOutcome = { status: "saved"; item: DriveItem; name: string } | { status: "error"; dto: ErrorDTO | null }

// An .xlsx uploaded beside `item`, under its name with the new extension (or " (2)"… when that is taken). `item`
// itself is left as it is.
export async function uploadXlsxCopy(item: DriveItem, converted: Uint8Array): Promise<SaveAsXlsxOutcome> {
	const base = asDirectoryOrFile(item)

	if (base.type !== "file") {
		return { status: "error", dto: null }
	}

	const rootUuid = currentRootUuid()
	const parent = normalizeParentUuid(base.data.parent, rootUuid)
	const name = await xlsxCopyName(driveItemName(base), candidate => sdkApi.nameExistsInDirectory(parent, candidate))
	const outcome = await runPreviewSave(
		{
			// The source's own type is not the copy's: an empty one lets the SDK take it from the .xlsx name.
			uploadFileBytes: (parentUuid, data, uploadName) => sdkApi.uploadFileBytes(parentUuid, data, uploadName, ""),
			patchListing: driveListingQueryUpdate,
			rootUuid
		},
		{ item, content: converted, asNewFile: name }
	)

	return outcome.status === "error" ? { status: "error", dto: outcome.dto } : { status: "saved", item: outcome.item, name }
}

// An .xls converted to an .xlsx and uploaded beside it (uploadXlsxCopy). `bytes` is the preview cache's buffer:
// the worker gets a copy.
export async function saveAsXlsx(item: DriveItem, bytes: Uint8Array): Promise<SaveAsXlsxOutcome> {
	if (asDirectoryOrFile(item).type !== "file") {
		return { status: "error", dto: null }
	}

	const copy = bytes.slice()
	const converted = await withSpreadsheetWorker(remote => remote.xlsToXlsx(Comlink.transfer(copy, [copy.buffer])))

	return await uploadXlsxCopy(item, converted)
}
