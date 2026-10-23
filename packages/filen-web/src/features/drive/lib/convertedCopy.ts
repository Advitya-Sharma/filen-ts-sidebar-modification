import { driveItemName } from "@filen/shared"
import { currentRootUuid } from "@/features/drive/lib/actions"
import { convertedCopyName } from "@/features/drive/lib/copyName"
import { asDirectoryOrFile, type DriveItem } from "@/features/drive/lib/item"
import { runPreviewSave } from "@/features/drive/lib/previewSave.logic"
import { driveListingQueryUpdate, normalizeParentUuid } from "@/features/drive/queries/drive"
import { sdkApi } from "@/lib/sdk/client"
import type { ErrorDTO } from "@/lib/sdk/errors"

// A converted copy saved beside the file it came from (Save as .xlsx for a legacy workbook or a text file, Save as
// .docx for a Word 97–2003 document): the file's name with the new extension, or " (2)"… while that is taken. The
// original is left as it is.

export type ConvertedCopyOutcome = { status: "saved"; item: DriveItem; name: string } | { status: "error"; dto: ErrorDTO | null }

export async function uploadConvertedCopy(item: DriveItem, converted: Uint8Array, extension: string): Promise<ConvertedCopyOutcome> {
	const base = asDirectoryOrFile(item)

	if (base.type !== "file") {
		return { status: "error", dto: null }
	}

	const rootUuid = currentRootUuid()
	const parent = normalizeParentUuid(base.data.parent, rootUuid)
	const name = await convertedCopyName(driveItemName(base), extension, candidate => sdkApi.nameExistsInDirectory(parent, candidate))
	const outcome = await runPreviewSave(
		{
			// The source's own type is not the copy's: an empty one lets the SDK take it from the new name.
			uploadFileBytes: (parentUuid, data, uploadName) => sdkApi.uploadFileBytes(parentUuid, data, uploadName, ""),
			patchListing: driveListingQueryUpdate,
			rootUuid
		},
		{ item, content: converted, asNewFile: name }
	)

	return outcome.status === "error" ? { status: "error", dto: outcome.dto } : { status: "saved", item: outcome.item, name }
}
