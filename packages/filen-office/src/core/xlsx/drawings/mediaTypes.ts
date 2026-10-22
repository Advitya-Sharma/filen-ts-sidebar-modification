// Content types of the media parts pictures are stored in, by extension (and image format name).
export const CONTENT_TYPES: Readonly<Record<string, string>> = {
	png: "image/png",
	jpeg: "image/jpeg",
	jpg: "image/jpeg",
	gif: "image/gif",
	bmp: "image/bmp",
	tif: "image/tiff",
	tiff: "image/tiff",
	svg: "image/svg+xml",
	emf: "image/x-emf",
	wmf: "image/x-wmf"
}

// The content type of a media part by its extension ("application/octet-stream" for one not known).
export function mediaContentType(name: string): string {
	const dot = name.lastIndexOf(".")

	return CONTENT_TYPES[dot < 0 ? "" : name.slice(dot + 1).toLowerCase()] ?? "application/octet-stream"
}
