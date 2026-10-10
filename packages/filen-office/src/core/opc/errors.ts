export type ZipErrorCode =
	"notZip" | "compoundFile" | "truncated" | "corrupt" | "encrypted" | "unsupported" | "duplicateName" | "tooManyEntries" | "tooLarge"

export class ZipError extends Error {
	readonly code: ZipErrorCode

	constructor(code: ZipErrorCode, message: string) {
		super(message)

		this.name = "ZipError"
		this.code = code
	}
}

export type OpcErrorCode =
	"invalidPartName" | "duplicatePart" | "missingPart" | "missingContentType" | "invalidContentTypes" | "invalidRelationships"

export class OpcError extends Error {
	readonly code: OpcErrorCode

	constructor(code: OpcErrorCode, message: string) {
		super(message)

		this.name = "OpcError"
		this.code = code
	}
}
