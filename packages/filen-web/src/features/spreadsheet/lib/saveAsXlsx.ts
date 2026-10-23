import type { DriveItem } from "@/features/drive/lib/item"
import { uploadConvertedCopy, type ConvertedCopyOutcome } from "@/features/drive/lib/convertedCopy"

export type SaveAsXlsxOutcome = ConvertedCopyOutcome

// An .xlsx uploaded beside `item`, under its name with the new extension (or " (2)"… when that is taken). `item`
// itself is left as it is.
export function uploadXlsxCopy(item: DriveItem, converted: Uint8Array): Promise<SaveAsXlsxOutcome> {
	return uploadConvertedCopy(item, converted, "xlsx")
}
