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
	refs: number
	load: Promise<string | undefined>
}

export interface UrlFactory {
	create(bytes: Uint8Array, type: string): string
	revoke(url: string): void
}

const browserUrls: UrlFactory = {
	create: (bytes, type) => URL.createObjectURL(new Blob([bytes as Uint8Array<ArrayBuffer>], { type })),
	revoke: url => URL.revokeObjectURL(url)
}

export class DocxMediaCache {
	private readonly session: DocxSessionApi
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

	constructor(session: DocxSessionApi, limits: MediaCacheLimits = DEFAULT_MEDIA_LIMITS, urls: UrlFactory = browserUrls) {
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

		const created: Entry = { key, url: undefined, bytes: 0, refs: 1, load: Promise.resolve(undefined) }

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

			created.url = this.urls.create(media.bytes, type)
			created.bytes = media.bytes.byteLength
			this.bytes += created.bytes
			this.trim()

			return created.url
		})()

		return created.load
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
