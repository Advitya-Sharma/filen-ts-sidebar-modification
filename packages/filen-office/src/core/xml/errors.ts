export type XmlErrorCode =
	| "doctype"
	| "encoding"
	| "unterminated"
	| "invalidName"
	| "invalidAttribute"
	| "invalidReference"
	| "mismatchedTag"
	| "unexpectedCloseTag"
	| "unclosedElement"
	| "contentOutsideRoot"
	| "invalidCharacter"
	| "invalidContent"
	| "unboundPrefix"
	| "editConflict"

export class XmlError extends Error {
	readonly code: XmlErrorCode
	// Byte offset into the part, or -1 when the error is not tied to one position.
	readonly offset: number

	constructor(code: XmlErrorCode, message: string, offset: number = -1) {
		super(offset >= 0 ? `${message} (byte ${offset})` : message)

		this.name = "XmlError"
		this.code = code
		this.offset = offset
	}
}
