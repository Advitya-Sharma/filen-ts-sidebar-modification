// Why a file could not be opened as an OpenDocument package.
export type OdfOpenFailure =
	// Entries encrypted under the manifest (ODF 1.3 Part 2 §4.8 encryption-data, §4.15 keyinfo/OpenPGP) or the
	// zip's own encryption flag. Refused in this version; OdfEncryption says which algorithms were used.
	| "encrypted"
	| "notZip"
	// An OLE2 compound file: a legacy binary Office file or an encrypted OOXML package under an ODF name.
	| "compoundFile"
	| "tooLarge"
	// A zip that is not an ODF package: no ODF media type in mimetype or the manifest, or a StarOffice 1.x file.
	| "notOdf"
	// An ODF package that is damaged beyond reading: an unreadable manifest, a missing content stream, a corrupt entry.
	| "invalidPackage"
	// A valid package of another kind than the caller expects (a spreadsheet opened as a text document).
	| "wrongKind"

export class OdfError extends Error {
	readonly reason: OdfOpenFailure

	constructor(reason: OdfOpenFailure, message: string) {
		super(message)

		this.name = "OdfError"
		this.reason = reason
	}
}
