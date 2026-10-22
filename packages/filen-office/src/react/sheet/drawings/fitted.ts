// Pictures far larger than the box they are shown in, decoded once at the box's size: the page shows a copy of that
// size, so the platform never holds the full-size decode for them. Copies are shared by part and size and kept while
// shown; unused ones stay (for scrolling back) within a byte budget, oldest first. Everything is revoked on dispose.

// A picture is fitted when it decodes to more than this many times the pixels its box shows.
export const FIT_RATIO = 4

// Box sizes are rounded up to this many device pixels: a small zoom or resize reuses the copy.
const FIT_STEP = 64

const FITTED_BYTES = 32 * 1024 * 1024

interface FittedEntry {
	readonly key: string
	url: string | undefined
	bytes: number
	refs: number
	load: Promise<string | undefined>
}

export interface FittedSize {
	readonly width: number
	readonly height: number
}

// The device pixels a copy is made at for a box of `width` x `height` CSS px showing the part `crop` leaves (a:srcRect
// fractions; the copy is of the whole picture, so its shown part comes out at the box's size).
export function fittedSize(
	width: number,
	height: number,
	ratio: number,
	crop: { readonly left: number; readonly top: number; readonly right: number; readonly bottom: number }
): FittedSize {
	const shown = (a: number, b: number): number => Math.min(1, Math.max(0.05, 1 - Math.max(0, a) - Math.max(0, b)))
	const step = (value: number): number => Math.max(FIT_STEP, Math.ceil((value * ratio) / FIT_STEP) * FIT_STEP)

	return { width: step(width / shown(crop.left, crop.right)), height: step(height / shown(crop.top, crop.bottom)) }
}

// Whether a picture of `pixels` is worth a copy at `size`.
export function wantsFit(pixels: number | undefined, size: FittedSize): boolean {
	return pixels !== undefined && pixels > FIT_RATIO * size.width * size.height
}

// Draws `source` (an object URL) at `size` and encodes it as PNG; undefined where the platform cannot.
async function scaledCopy(source: string, size: FittedSize): Promise<Blob | undefined> {
	if (typeof createImageBitmap !== "function" || typeof document === "undefined") return undefined

	const blob = await (await fetch(source)).blob()
	const bitmap = await createImageBitmap(blob, { resizeWidth: size.width, resizeHeight: size.height, resizeQuality: "high" })

	try {
		const canvas = document.createElement("canvas")

		canvas.width = size.width
		canvas.height = size.height

		const context = canvas.getContext("2d")

		if (context === null) return undefined

		context.drawImage(bitmap, 0, 0, size.width, size.height)

		const encoded = await new Promise<Blob | null>(resolve => {
			canvas.toBlob(resolve, "image/png")
		})

		// The canvas's backing store goes with it.
		canvas.width = 0
		canvas.height = 0

		return encoded ?? undefined
	} finally {
		bitmap.close()
	}
}

export class FittedPictures {
	// Insertion order is recency order.
	private readonly entries = new Map<string, FittedEntry>()
	private bytes = 0
	private disposed = false

	// A copy of `partName` (shown from `source`) at `size`, held until release(key); undefined when none could be
	// made (the full picture is shown then).
	acquire(partName: string, source: string, size: FittedSize): { readonly key: string; readonly url: Promise<string | undefined> } {
		const key = `${partName}\u0000${String(size.width)}x${String(size.height)}`
		let entry = this.entries.get(key)

		if (entry !== undefined) {
			this.entries.delete(key)
			this.entries.set(key, entry)
			entry.refs++

			return { key, url: entry.load }
		}

		const created: FittedEntry = { key, url: undefined, bytes: 0, refs: 1, load: Promise.resolve(undefined) }

		entry = created
		this.entries.set(key, created)
		created.load = scaledCopy(source, size).then(
			blob => {
				if (this.disposed || this.entries.get(key) !== created) return undefined

				// Not remembered: the full picture shows, and a later mount may try again.
				if (blob === undefined) {
					this.entries.delete(key)

					return undefined
				}

				created.url = URL.createObjectURL(blob)
				created.bytes = blob.size
				this.bytes += blob.size
				this.trim()

				return created.url
			},
			() => {
				if (this.entries.get(key) === created) this.entries.delete(key)

				return undefined
			}
		)

		return { key, url: created.load }
	}

	release(key: string): void {
		const entry = this.entries.get(key)

		if (entry === undefined) return

		entry.refs = Math.max(0, entry.refs - 1)
		this.trim()
	}

	private trim(): void {
		for (const entry of this.entries.values()) {
			if (this.bytes <= FITTED_BYTES) return
			if (entry.refs > 0 || entry.url === undefined) continue

			this.entries.delete(entry.key)
			this.bytes -= entry.bytes
			URL.revokeObjectURL(entry.url)
			entry.url = undefined
		}
	}

	dispose(): void {
		if (this.disposed) return

		this.disposed = true

		for (const entry of this.entries.values()) if (entry.url !== undefined) URL.revokeObjectURL(entry.url)

		this.entries.clear()
	}
}
