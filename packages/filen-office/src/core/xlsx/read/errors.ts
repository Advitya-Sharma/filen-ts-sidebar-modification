export type XlsxReadErrorCode =
	// Markup that does not read as the part it should be.
	| "malformed"
	// Past a part, count or cell limit (XlsxLimits).
	| "tooLarge"
	// A part the workbook points to is not in the package.
	| "missingPart"
	| "unsupported"
	// The caller's signal was aborted.
	| "aborted"
	// cellXfs is full (Excel's limit): the edit refuses.
	| "styleLimit"
	// A relationship was added to a part whose relationships were never read.
	| "relationshipsNotRead"

// Every failure the reader and the model raise: anything else escaping them is a bug.
export class XlsxReadError extends Error {
	readonly code: XlsxReadErrorCode
	readonly part: string
	// Byte offset into the part, or -1.
	readonly offset: number

	constructor(code: XlsxReadErrorCode, message: string, part: string = "", offset: number = -1) {
		super(offset >= 0 ? `${message} (${part} byte ${offset})` : part === "" ? message : `${message} (${part})`)

		this.name = "XlsxReadError"
		this.code = code
		this.part = part
		this.offset = offset
	}
}
