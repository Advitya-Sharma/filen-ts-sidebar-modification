import { rasterOpaque, rasterPixels, rasterSize } from "../../core/docx/objects/images"
import type { DocxSessionApi } from "../../dom/docx/protocol"

// Object URLs for a document's images, read from the package only when a mounted page shows them. Entries
// in use by a mounted page are never revoked; unused ones stay (for scrolling back) until the cache is over
// its budget, oldest first. Everything is revoked on dispose.

export interface MediaCacheLimits {
	readonly maxBytes: number
	readonly maxEntries: number
}

export const DEFAULT_MEDIA_LIMITS: MediaCacheLimits = { maxBytes: 96 * 1024 * 1024, maxEntries: 128 }

interface Entry {
	readonly key: string
	url: string | undefined
	bytes: number
	// The picture as the URL serves it, where the URL factory makes URLs of blobs.
	blob: Blob | undefined
	// Pixels a raster picture decodes to (its header's), undefined for others.
	pixels: number | undefined
	// A PNG, JPEG or BMP's size in pixels (its header's).
	size: { readonly width: number; readonly height: number } | undefined
	// No transparency (rasterOpaque).
	opaque: boolean
	refs: number
	load: Promise<string | undefined>
}

export interface UrlFactory {
	create(bytes: Uint8Array, type: string): string
	revoke(url: string): void
	// An object URL for a blob already made (the blob is kept beside it, for a decoder that takes one).
	createFromBlob?(blob: Blob): string
}

const browserUrls: UrlFactory = {
	create: (bytes, type) => URL.createObjectURL(new Blob([bytes as Uint8Array<ArrayBuffer>], { type })),
	createFromBlob: blob => URL.createObjectURL(blob),
	revoke: url => URL.revokeObjectURL(url)
}

export class DocxMediaCache {
	private readonly session: Pick<DocxSessionApi, "media">
	private readonly limits: MediaCacheLimits
	private readonly urls: UrlFactory
	// Insertion order is recency order: a hit moves the entry to the end.
	private readonly entries = new Map<string, Entry>()
	private bytes = 0
	private disposed = false
	// Object URLs of bytes drawings carry, until revokeBytesUrl() or dispose.
	private readonly bytesUrls = new Set<string>()
	// Settles when the font faces the document registered so far have loaded (text a drawing lays out is
	// measured with them); the document's store sets it.
	fontsReady: () => Promise<unknown> = () => Promise.resolve()

	constructor(session: Pick<DocxSessionApi, "media">, limits: MediaCacheLimits = DEFAULT_MEDIA_LIMITS, urls: UrlFactory = browserUrls) {
		this.session = session
		this.limits = limits
		this.urls = urls
	}

	// An object URL for `partName` shown as `type`, held until release(); undefined when the part is missing,
	// too large or not an image the platform shows. Every acquire needs one release.
	acquire(partName: string, type: string): Promise<string | undefined> {
		const key = `${type}\u0000${partName}`
		let entry = this.entries.get(key)

		if (entry !== undefined) {
			this.entries.delete(key)
			this.entries.set(key, entry)
			entry.refs++

			return entry.load
		}

		const created: Entry = {
			key,
			url: undefined,
			bytes: 0,
			blob: undefined,
			pixels: undefined,
			size: undefined,
			opaque: false,
			refs: 1,
			load: Promise.resolve(undefined)
		}

		entry = created
		this.entries.set(key, created)
		created.load = (async () => {
			if (type.length === 0) {
				this.entries.delete(key)

				return undefined
			}

			const media = await this.session.media(partName).catch(() => null)

			if (this.disposed || this.entries.get(key) !== created) return undefined

			// Not remembered: a page mounted later asks again (a read can fail transiently).
			if (media === null) {
				this.entries.delete(key)

				return undefined
			}

			if (this.urls.createFromBlob === undefined) {
				created.url = this.urls.create(media.bytes, type)
			} else {
				created.blob = new Blob([media.bytes as Uint8Array<ArrayBuffer>], { type })
				created.url = this.urls.createFromBlob(created.blob)
			}

			created.bytes = media.bytes.byteLength
			created.pixels = rasterPixels(media.bytes)
			created.size = rasterSize(media.bytes)
			created.opaque = rasterOpaque(media.bytes)
			this.bytes += created.bytes
			this.trim()

			return created.url
		})()

		return created.load
	}

	// An acquired picture as a blob (what its URL serves); undefined before it loaded, or where URLs are not of blobs.
	blob(partName: string, type: string): Blob | undefined {
		return this.entries.get(`${type}\u0000${partName}`)?.blob
	}

	// An acquired PNG, JPEG or BMP's size in pixels; undefined before it loaded, or for other formats.
	rasterSize(partName: string, type: string): { readonly width: number; readonly height: number } | undefined {
		return this.entries.get(`${type}\u0000${partName}`)?.size
	}

	// Whether an acquired picture is a PNG, JPEG or BMP without transparency.
	opaque(partName: string, type: string): boolean {
		return this.entries.get(`${type}\u0000${partName}`)?.opaque === true
	}

	// The pixels an acquired raster picture decodes to; undefined before it loaded, or for a vector picture.
	pixels(partName: string, type: string): number | undefined {
		return this.entries.get(`${type}\u0000${partName}`)?.pixels
	}

	release(partName: string, type: string): void {
		const entry = this.entries.get(`${type}\u0000${partName}`)

		if (entry === undefined) return

		entry.refs = Math.max(0, entry.refs - 1)
		this.trim()
	}

	private trim(): void {
		for (const entry of this.entries.values()) {
			if (this.bytes <= this.limits.maxBytes && this.entries.size <= this.limits.maxEntries) return
			if (entry.refs > 0 || entry.url === undefined) continue

			this.drop(entry)
		}
	}

	private drop(entry: Entry): void {
		this.entries.delete(entry.key)
		this.bytes -= entry.bytes

		if (entry.url !== undefined) this.urls.revoke(entry.url)

		entry.url = undefined
		entry.blob = undefined
	}

	// An object URL for image bytes a drawing carries (a bitmap inside a metafile); undefined after dispose.
	// Every URL needs one revokeBytesUrl().
	createBytesUrl(bytes: Uint8Array, type: string): string | undefined {
		if (this.disposed || type.length === 0) return undefined

		const url = this.urls.create(bytes, type)

		this.bytesUrls.add(url)

		return url
	}

	revokeBytesUrl(url: string): void {
		if (!this.bytesUrls.delete(url)) return

		this.urls.revoke(url)
	}

	get size(): number {
		return this.entries.size
	}

	get byteSize(): number {
		return this.bytes
	}

	dispose(): void {
		if (this.disposed) return

		this.disposed = true

		for (const entry of [...this.entries.values()]) this.drop(entry)
		for (const url of this.bytesUrls) this.urls.revoke(url)

		this.bytesUrls.clear()
	}
}
