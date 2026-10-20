// pdf.js failures as a closed set. Matched on the exception's name, not instanceof: the classes belong to
// whichever pdf.js build the host injected. No pdf.js message is ever passed on, as a document controls parts
// of those strings.

export type PdfErrorKind =
	// Not a PDF, or damaged beyond what pdf.js recovers.
	| "invalidDocument"
	// The bytes could not be read.
	| "transportFailed"
	// The in-thread worker was asked for but the host did not load pdf.js's worker module.
	| "workerUnavailable"
	| "unknown"

export type PdfPasswordReason = "required" | "incorrect"

export type PdfFailure =
	| { readonly type: "password"; readonly reason: PdfPasswordReason }
	| { readonly type: "error"; readonly kind: PdfErrorKind }
	| { readonly type: "aborted" }

// pdf.js PasswordResponses.INCORRECT_PASSWORD (NEED_PASSWORD is 1).
const INCORRECT_PASSWORD = 2

export function passwordReason(code: unknown): PdfPasswordReason {
	return code === INCORRECT_PASSWORD ? "incorrect" : "required"
}

// Thrown by openPdf when the in-thread worker is missing.
export class PdfWorkerUnavailableError extends Error {
	override readonly name = "PdfWorkerUnavailableError"

	constructor() {
		super("pdf.js worker module not loaded")
	}
}

export function pdfErrorName(error: unknown): string | undefined {
	if (typeof error !== "object" || error === null) return undefined

	const name = (error as { readonly name?: unknown }).name

	return typeof name === "string" ? name : undefined
}

export function classifyPdfError(error: unknown): PdfFailure {
	switch (pdfErrorName(error)) {
		case "PasswordException":
			return { type: "password", reason: passwordReason((error as { readonly code?: unknown }).code) }
		case "InvalidPDFException":
			return { type: "error", kind: "invalidDocument" }
		// 6.x folded MissingPDFException and UnexpectedResponseException into this one.
		case "ResponseException":
		case "PdfReadError":
			return { type: "error", kind: "transportFailed" }
		case "PdfWorkerUnavailableError":
			return { type: "error", kind: "workerUnavailable" }
		// A cancelled render or a destroyed task is the expected end of scrolling away or closing.
		case "RenderingCancelledException":
		case "AbortException":
			return { type: "aborted" }
		default:
			return { type: "error", kind: "unknown" }
	}
}
