export type FontErrorCode = "notSfnt" | "unsupported" | "truncated" | "corrupt" | "tooLarge" | "invalidKey"

export class FontError extends Error {
	readonly code: FontErrorCode

	constructor(code: FontErrorCode, message: string) {
		super(message)

		this.name = "FontError"
		this.code = code
	}
}
