// zlib deflate (RFC 1950, what /FlateDecode reads, ISO 32000-1 §7.4.4) through the platform's
// CompressionStream("deflate"): native, fast and off the JS heap. Where there is none (Hermes) streams are
// written uncompressed, which is still a valid file.

export type PdfDeflate = (data: Uint8Array) => Promise<Uint8Array>

interface DeflateStream {
	readonly readable: unknown
	readonly writable: {
		getWriter(): { write(chunk: Uint8Array): Promise<void>; close(): Promise<void>; abort(reason: unknown): Promise<void> }
	}
}

export function nativeZlibDeflate(): PdfDeflate | undefined {
	const Compression = Reflect.get(globalThis, "CompressionStream") as (new (format: string) => DeflateStream) | undefined
	const Body = Reflect.get(globalThis, "Response") as (new (body: unknown) => { arrayBuffer(): Promise<ArrayBuffer> }) | undefined

	if (typeof Compression !== "function" || typeof Body !== "function") return undefined

	try {
		new Compression("deflate")
	} catch {
		return undefined
	}

	return async data => {
		const stream = new Compression("deflate")
		const writer = stream.writable.getWriter()
		const result = new Body(stream.readable).arrayBuffer()

		result.catch(() => undefined)

		try {
			await Promise.all([writer.write(data), writer.close()])
		} catch (error) {
			await writer.abort(error).catch(() => undefined)

			throw error
		}

		return new Uint8Array(await result)
	}
}
