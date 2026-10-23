import { encodeUtf8, escapeAttribute, escapeText } from "../xml/index"

// Markup helpers for the WordprocessingML the translator writes.

export const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
export const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
export const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

// Every namespace a part may use, declared on its root.
export const PART_NAMESPACES = [
	`xmlns:w="${NS_W}"`,
	`xmlns:r="${NS_R}"`,
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
	'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
	'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
	'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
].join(" ")

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

// An attribute value.
export function esc(text: string): string {
	return escapeAttribute(text)
}

// Character data: ">" escaped too, so "]]>" never reaches content (XML 1.0 §2.4).
export function escText(text: string): string {
	return escapeText(text)
}

// Points to twentieths of a point, eighths (border widths) and EMUs.
export function twips(points: number): number {
	return Math.round(points * 20)
}

export function eighths(points: number): number {
	return Math.round(points * 8)
}

export function emu(points: number): number {
	return Math.round(points * 12700)
}

// A part's markup gathered as strings and encoded in steps, so a long body never sits as one huge string.
export class PartWriter {
	private readonly chunks: Uint8Array[] = []
	private pending: string[] = []
	private pendingLength = 0
	private size = 0

	add(markup: string): void {
		this.pending.push(markup)
		this.pendingLength += markup.length

		if (this.pendingLength >= 64 * 1024) this.encode()
	}

	get byteLength(): number {
		return this.size + this.pendingLength
	}

	bytes(): Uint8Array {
		this.encode()

		if (this.chunks.length === 1 && this.chunks[0] !== undefined) return this.chunks[0]

		const out = new Uint8Array(this.size)
		let offset = 0

		for (const chunk of this.chunks) {
			out.set(chunk, offset)
			offset += chunk.length
		}

		this.chunks.length = 0
		this.chunks.push(out)

		return out
	}

	private encode(): void {
		if (this.pending.length === 0) return

		const bytes = encodeUtf8(this.pending.join(""))

		this.chunks.push(bytes)
		this.size += bytes.length
		this.pending = []
		this.pendingLength = 0
	}
}
