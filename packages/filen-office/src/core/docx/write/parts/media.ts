import { partKey, relativeTarget } from "../../../opc/index"
import { sniffImage } from "../../objects/images"
import { DocxWriteError } from "../values"
import type { PartContributor, SaveParts } from "./index"

// Media parts of new pictures: each picture's bytes (DocxMediaStore) become one part word/media/imageN.ext,
// typed by a Default for its extension as Word types media (an Override where the package types the
// extension otherwise), and each story part that shows it gets one image relationship to it. Written as the
// pictures are, so a picture nothing writes adds nothing.

export const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

interface MediaState {
	// Handle -> part name.
	readonly names: Map<number, string>
	// "part\nhandle" -> relationship id.
	readonly relationships: Map<string, string>
}

const MEDIA_STATE = {}

function mediaState(parts: SaveParts): MediaState {
	return parts.state(MEDIA_STATE, () => ({ names: new Map(), relationships: new Map() }))
}

// The part a picture's bytes are written to (added on first use).
function mediaPart(parts: SaveParts, handle: number): string {
	const state = mediaState(parts)
	const known = state.names.get(handle)

	if (known !== undefined) return known

	const bytes = parts.context.media.bytes(handle)
	const image = bytes === undefined ? undefined : sniffImage(bytes)

	if (bytes === undefined || image === undefined || typeof image === "string")
		throw new DocxWriteError("missingSource", `The picture of media ${String(handle)}`)

	let name = ""

	for (let n = 1; name === ""; n++) {
		const candidate = `word/media/image${String(n)}.${image.extension}`

		if (parts.freeName(candidate) === candidate) name = candidate
	}

	if (!parts.pkg.contentTypes.defaults.has(partKey(image.extension))) parts.addDefault(image.extension, image.contentType)

	parts.add(name, bytes, image.contentType)
	state.names.set(handle, name)

	return name
}

// The relationship id from story part `part` to the picture of `handle`, made on first use.
export function mediaRelationship(parts: SaveParts, part: string, handle: number): string {
	const state = mediaState(parts)
	const key = `${part}\n${String(handle)}`
	const known = state.relationships.get(key)

	if (known !== undefined) return known

	const id = parts.relationship(part, { type: IMAGE_RELATIONSHIP, target: relativeTarget(part, mediaPart(parts, handle)) })

	state.relationships.set(key, id)

	return id
}

// Parts are added as the pictures are written (mediaRelationship).
export const mediaParts: PartContributor = {}
