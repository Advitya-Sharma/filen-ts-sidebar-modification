import type { FontFaceDescriptor, FontHost } from "../core/fonts/index"

// FontHost over the CSS Font Loading API, for a window (`{ FontFace, fonts: document.fonts }`) or a worker
// (`self`). Fonts load from bytes or local(): the apps' `font-src 'self'` blocks blob: and data: URLs, and
// bytes need no fetch from a WebView's file:// origin. The shapes are declared here so the package needs
// no DOM typings; the platform objects satisfy them.

interface FontFaceLike {
	load(): Promise<unknown>
}

export interface FontFaceScope {
	readonly FontFace: new (family: string, source: string | ArrayBuffer, descriptors: { weight: string; style: string }) => FontFaceLike
	readonly fonts: {
		add(face: FontFaceLike): unknown
		delete(face: FontFaceLike): unknown
	}
}

function descriptors(face: FontFaceDescriptor): { weight: string; style: string } {
	return {
		weight: typeof face.weight === "number" ? String(face.weight) : `${face.weight[0]} ${face.weight[1]}`,
		style: face.italic ? "italic" : "normal"
	}
}

// The platform copies font data, so a view over a larger buffer is sliced only when it has to be.
function arrayBufferOf(data: Uint8Array): ArrayBuffer {
	const buffer = data.buffer

	if (buffer instanceof ArrayBuffer && data.byteOffset === 0 && data.byteLength === buffer.byteLength) return buffer

	return data.slice().buffer
}

// A CSS string for local(): the name comes from a document, so quotes, backslashes and control
// characters cannot be allowed to end the string or the descriptor.
export function localSource(name: string): string {
	let escaped = ""

	for (const char of name) {
		const code = char.charCodeAt(0)

		if (code < 0x20 || code === 0x7f) continue

		escaped += char === '"' || char === "\\" ? `\\${char}` : char
	}

	return `local("${escaped}")`
}

export function fontFaceHost(scope: FontFaceScope): FontHost {
	const faces = new Map<string, FontFaceLike[]>()
	// Families removed while one of their faces was still loading: that face must not be added after.
	const removed = new Set<string>()

	const add = async (face: FontFaceDescriptor, source: string | ArrayBuffer): Promise<boolean> => {
		let fontFace: FontFaceLike

		try {
			fontFace = new scope.FontFace(face.family, source, descriptors(face))
			await fontFace.load()
		} catch {
			return false
		}

		if (removed.has(face.family)) return false

		scope.fonts.add(fontFace)

		const list = faces.get(face.family)

		if (list === undefined) faces.set(face.family, [fontFace])
		else list.push(fontFace)

		return true
	}

	return {
		addBytes: (face, data) => add(face, arrayBufferOf(data)),
		addLocal: (face, name) => add(face, localSource(name)),
		remove(family) {
			for (const face of faces.get(family) ?? []) scope.fonts.delete(face)

			faces.delete(family)
			// Aliases are never reused, so the set only grows by the families a document used.
			removed.add(family)
		}
	}
}
