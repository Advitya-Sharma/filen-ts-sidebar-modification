// What the font manager needs from the platform. Web and the mobile WebView implement FontHost with
// FontFace on the document or worker scope (see @filen/office/dom); each realm that measures or draws
// text needs its own host, since a FontFaceSet belongs to one realm.

export interface FontFaceDescriptor {
	// The internal alias: never a name taken from a document.
	readonly family: string
	// A [min, max] range for variable fonts.
	readonly weight: number | readonly [number, number]
	readonly italic: boolean
}

export interface FontHost {
	// Registers font file bytes (TrueType, OpenType or WOFF2). False when the platform rejects the data.
	// The host must not keep `data` beyond the call: the platform copies it.
	addBytes(face: FontFaceDescriptor, data: Uint8Array): Promise<boolean>
	// Registers an installed font by its full or PostScript name. False when it is not installed or the
	// platform does not expose installed fonts.
	addLocal(face: FontFaceDescriptor, name: string): Promise<boolean>
	// Unregisters every face of an alias.
	remove(family: string): void
}

// The bundled font files by their path under assets/fonts: web fetches them from its own origin, mobile
// reads them from the app bundle.
export interface FontAssets {
	read(file: string): Promise<Uint8Array>
}
