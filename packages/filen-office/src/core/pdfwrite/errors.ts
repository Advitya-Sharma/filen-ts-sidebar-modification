// Typed refusals of the PDF writer: the caller decides what to show (and never gets a partial file).
export type PdfWriteErrorReason =
	// An image the writer cannot embed: unknown or damaged encoding, or the platform cannot decode it.
	| "imageUnsupported"
	// An image over the pixel budget for decoding.
	| "imageTooLarge"
	// Marked content, artifacts or saved graphics states left open at the end of a page or the document.
	| "unbalanced"
	// A password the standard security handler cannot take (SASLprep prohibits it, ISO 32000-2 §7.6.4.3.3).
	| "password"
	// The platform lacks something the request needs (WebCrypto for encryption).
	| "unsupported"
	// A call out of order: drawing on a finished page, or after the document was finished.
	| "state"

export class PdfWriteError extends Error {
	readonly reason: PdfWriteErrorReason

	constructor(reason: PdfWriteErrorReason, message: string) {
		super(message)
		this.reason = reason
		this.name = "PdfWriteError"
	}
}
