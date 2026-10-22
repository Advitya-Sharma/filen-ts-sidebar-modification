import { isSealedOfficeFile } from "@filen/office/dom/encryption"

// What leaves a protected workbook's save for upload is a compound file ([MS-OFFCRYPTO] §2.3.4), never its zip.
export function assertSealedUpload(bytes: Uint8Array): void {
	if (!isSealedOfficeFile(bytes)) {
		throw new Error("spreadsheet: a protected workbook's save is not sealed")
	}
}
