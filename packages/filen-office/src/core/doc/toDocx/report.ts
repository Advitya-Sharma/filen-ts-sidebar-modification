// What a conversion could not carry over as it was, by kind, so a viewer can say so. Nothing is dropped
// silently: each item is either converted, kept in another form (a placeholder picture), or counted here.

export type DocNotConvertedKind =
	// A floating shape other than a picture, text box or simple preset shape: shown as a placeholder picture.
	| "shape"
	// A group of shapes: shown as a placeholder picture, the text its shapes hold in text boxes where it is.
	| "group"
	// WordArt: its text in a text box of the shape's size (not drawn as WordArt).
	| "wordArt"
	// The data of an embedded object (only its picture is shown).
	| "embeddedObject"
	// A picture whose data could not be read or is in a format the document cannot show: a placeholder picture.
	| "picture"
	// A page or paragraph art border: drawn as a single line.
	| "artBorder"
	// A Word binary special character with no WordprocessingML equivalent.
	| "specialCharacter"
	// A password on the document's editing restrictions that could not be carried over (no hash function where the
	// conversion ran): the restriction stays enforced, without its password.
	| "passwordProtection"
	// Content the reader could not read (damaged text, a structure it could not parse, unbalanced fields or tables).
	| "damagedContent"
	// A VBA project (macros), which a .docx cannot hold.
	| "macros"
	// A template's AutoText document, which the package does not carry.
	| "autoText"

export interface DocNotConverted {
	readonly kind: DocNotConvertedKind
	readonly count: number
}

export class ConversionReport {
	private readonly counts = new Map<DocNotConvertedKind, number>()

	note(kind: DocNotConvertedKind, count = 1): void {
		this.counts.set(kind, (this.counts.get(kind) ?? 0) + count)
	}

	list(): DocNotConverted[] {
		return [...this.counts].map(([kind, count]) => ({ kind, count }))
	}
}
