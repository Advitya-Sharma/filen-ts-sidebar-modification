// Why content.xml (or another stream of a spreadsheet package) could not be read.
export type OdsReadFailure = "malformed" | "tooLarge" | "aborted" | "unsupported"

export class OdsReadError extends Error {
	readonly code: OdsReadFailure

	constructor(code: OdsReadFailure, message: string) {
		super(message)

		this.name = "OdsReadError"
		this.code = code
	}
}
