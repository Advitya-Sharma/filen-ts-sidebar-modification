// Why a save was refused, in a module of its own: the zip half of a save runs where the editor does not (a
// worker) and must not pull the writer in.

export type VerifyFailure =
	| "invalidPackage"
	| "untouchedChanged"
	| "malformedPart"
	| "contentType"
	| "relationship"
	| "duplicateId"
	| "danglingReference"
	| "model"
	| "content"

export class DocxVerifyError extends Error {
	readonly reason: VerifyFailure

	constructor(reason: VerifyFailure, message: string) {
		super(message)
		this.name = "DocxVerifyError"
		this.reason = reason
	}
}

export type DocxSaveFailure = "readOnly" | "write" | VerifyFailure

const FAILURES: ReadonlySet<string> = new Set<DocxSaveFailure>([
	"readOnly",
	"write",
	"invalidPackage",
	"untouchedChanged",
	"malformedPart",
	"contentType",
	"relationship",
	"duplicateId",
	"danglingReference",
	"model",
	"content"
])

export function isDocxSaveFailure(value: unknown): value is DocxSaveFailure {
	return typeof value === "string" && FAILURES.has(value)
}

export class DocxSaveError extends Error {
	readonly reason: DocxSaveFailure

	constructor(reason: DocxSaveFailure, message: string) {
		super(message)
		this.name = "DocxSaveError"
		this.reason = reason
	}
}
