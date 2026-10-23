import { Inflate } from "fflate"
import { METHOD_DEFLATE, METHOD_STORED } from "../opc/zipReader"
import { ZipError, type ZipArchive, type ZipEntry } from "../opc/index"
import { locateBody } from "./prologue"

// Compressed bytes read per step while looking for office:body.
const STEP = 64 * 1024
const BODY = new TextEncoder().encode("body")

// The start of content.xml up to and a little past office:body's start tag name, inflated only as far as
// that: a 300 MB spreadsheet body is never inflated to read its styles. Returns the whole stream when it has
// no office:body. A prefix cannot be CRC-checked; the body reader checks the whole stream when it reads it.
export async function readUntilBody(archive: ZipArchive, entry: ZipEntry): Promise<Uint8Array> {
	if (entry.method !== METHOD_STORED && entry.method !== METHOD_DEFLATE) return archive.read(entry)

	const { dataOffset } = await archive.layout(entry)
	let out = new Uint8Array(Math.min(entry.uncompressedSize, 256 * 1024))
	let length = 0
	let found = false
	let checked = 0
	const accept = (chunk: Uint8Array): void => {
		if (length + chunk.length > entry.uncompressedSize) throw new ZipError("corrupt", `${entry.name} inflates past its declared size`)

		if (length + chunk.length > out.length) {
			const grown = new Uint8Array(Math.min(entry.uncompressedSize, Math.max(out.length * 2, length + chunk.length)))

			grown.set(out.subarray(0, length))
			out = grown
		}

		out.set(chunk, length)
		length += chunk.length
	}
	// Looks again only when the new bytes (with an overlap for a split name) hold "body".
	const look = (): boolean => {
		const from = Math.max(0, checked - 64)
		const view = out.subarray(0, length)

		checked = length

		for (let at = view.indexOf(BODY[0] ?? 0, from); at >= 0 && at + BODY.length <= length; at = view.indexOf(BODY[0] ?? 0, at + 1)) {
			if (view[at + 1] === BODY[1] && view[at + 2] === BODY[2] && view[at + 3] === BODY[3]) return locateBody(view) !== undefined
		}

		return false
	}

	if (entry.method === METHOD_STORED) {
		for (let read = 0; read < entry.compressedSize && !found; read += STEP) {
			accept(await archive.source.read(dataOffset + read, Math.min(STEP, entry.compressedSize - read)))
			found = look()
		}
	} else {
		const inflater = new Inflate(accept)

		try {
			for (let read = 0; read < entry.compressedSize && !found; read += STEP) {
				const step = Math.min(STEP, entry.compressedSize - read)

				inflater.push(await archive.source.read(dataOffset + read, step), read + step >= entry.compressedSize)
				found = look()
			}
		} catch (error) {
			if (error instanceof ZipError) throw error

			throw new ZipError("corrupt", `Damaged deflate data in ${entry.name}`)
		}
	}

	return out.subarray(0, length)
}
