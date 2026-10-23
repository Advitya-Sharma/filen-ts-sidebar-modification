import { createReadStream } from "node:fs"
import { Readable } from "node:stream"
import type { PdfByteStream, PdfFontAssets } from "../core/pdfwrite/fonts/assets"

// The PDF font assets from the package's own files, and a PDF's parts as one buffer (tests write no Blob parts).

const ASSETS = new URL("../../assets/fonts/", import.meta.url)

export const nodeFontAssets: PdfFontAssets = {
	open: file => Promise.resolve(Readable.toWeb(createReadStream(new URL(file, ASSETS))) as unknown as PdfByteStream)
}

export function pdfBytes(parts: readonly (Uint8Array | { readonly size: number })[]): Uint8Array {
	const chunks = parts.filter((part): part is Uint8Array => part instanceof Uint8Array)

	if (chunks.length !== parts.length) throw new Error("A part is not bytes")

	return Uint8Array.from(Buffer.concat(chunks))
}
