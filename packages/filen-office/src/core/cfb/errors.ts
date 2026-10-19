export type CfbErrorCode =
	| "notCompoundFile"
	| "unsupported"
	| "truncated"
	| "corrupt"
	| "tooLarge"
	| "tooManyEntries"
	| "notFound"
	| "notStream"
	| "invalidName"
	| "duplicateName"
	| "invalidClsid"
	| "lengthMismatch"

export class CfbError extends Error {
	readonly code: CfbErrorCode

	constructor(code: CfbErrorCode, message: string) {
		super(message)

		this.name = "CfbError"
		this.code = code
	}
}
