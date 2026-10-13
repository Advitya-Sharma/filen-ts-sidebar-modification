import {
	bytesSource,
	crc32,
	memorySink,
	partKey,
	writeZip,
	OpcPackage,
	type ByteSink,
	type ByteSource,
	type PartBytes,
	type PartChange,
	type WriteZipOptions,
	type ZipEntry,
	type ZipPlan,
	OpcError,
	ZipError
} from "../../opc/index"
import { equalBytes } from "./bytes"
import { DocxSaveError, DocxVerifyError } from "./failure"

// The zip half of a save: the checked parts written into the original archive (every untouched entry copied
// as its compressed bytes), then the zip half of gate 2 on the result. It needs only the original archive
// and the changed parts, so it runs wherever the bytes are (a worker on web, next to the file on mobile),
// without the editor or the model.

const CHUNK = 4 * 1024 * 1024

// The parts a save writes: replaced and new, checked by the first half of gate 2 (prepareDocxSave).
export interface DocxPartChanges {
	readonly replace: readonly PartChange[]
	readonly add: readonly PartChange[]
	// Parts taken out (nothing relates to them any more).
	readonly remove?: readonly string[]
}

export interface DocxZipOptions {
	readonly zip?: WriteZipOptions
	// Where the file goes (default: memory, returned as `bytes`).
	readonly sink?: ByteSink
	// Random access to what `sink` received, for the zip half of gate 2; required with a sink.
	readonly readBack?: () => Promise<ByteSource>
	// Refuse a changed part that comes without PartChange.crc32 (where the part checks ran elsewhere).
	readonly requireChecksums?: boolean
	// Called as the zip half of gate 2 goes.
	readonly progress?: () => void
}

export interface DocxZipResult {
	// The file, when no sink was given.
	readonly bytes: Uint8Array | undefined
	// Parts rewritten or added; every other entry was copied as it was, compressed bytes included.
	readonly changedParts: readonly string[]
	readonly zipMs: number
	readonly verifyZipMs: number
}

interface DeflateStream {
	readonly readable: unknown
	readonly writable: {
		getWriter(): { write(chunk: Uint8Array): Promise<void>; close(): Promise<void>; abort(reason: unknown): Promise<void> }
	}
}

// The platform's own raw deflate (CompressionStream "deflate-raw", in browsers, workers and Node), several
// times faster than fflate on large parts; undefined where there is none (Hermes).
export function nativeDeflate(): ((data: Uint8Array) => Promise<Uint8Array>) | undefined {
	const Compression = Reflect.get(globalThis, "CompressionStream") as (new (format: string) => DeflateStream) | undefined
	const Body = Reflect.get(globalThis, "Response") as (new (body: unknown) => { arrayBuffer(): Promise<ArrayBuffer> }) | undefined

	if (typeof Compression !== "function" || typeof Body !== "function") return undefined

	try {
		new Compression("deflate-raw")
	} catch {
		return undefined
	}

	return async data => {
		const stream = new Compression("deflate-raw")
		const writer = stream.writable.getWriter()
		const written = Promise.all([writer.write(data), writer.close()])
		const result = new Uint8Array(await new Body(stream.readable).arrayBuffer())

		await written

		return result
	}
}

function fail(reason: "invalidPackage" | "untouchedChanged", message: string): never {
	throw new DocxVerifyError(reason, message)
}

async function sameRaw(
	a: OpcPackage,
	entryA: ZipEntry,
	b: OpcPackage,
	entryB: ZipEntry,
	progress: (() => void) | undefined
): Promise<boolean> {
	if (
		entryA.method !== entryB.method ||
		entryA.crc32 !== entryB.crc32 ||
		entryA.compressedSize !== entryB.compressedSize ||
		entryA.uncompressedSize !== entryB.uncompressedSize
	) {
		return false
	}

	const layoutA = await a.zip.layout(entryA)
	const layoutB = await b.zip.layout(entryB)

	for (let offset = 0; offset < entryA.compressedSize; offset += CHUNK) {
		const length = Math.min(CHUNK, entryA.compressedSize - offset)
		const left = await a.zip.source.read(layoutA.dataOffset + offset, length)
		const right = await b.zip.source.read(layoutB.dataOffset + offset, length)

		if (!equalBytes(left, 0, right, 0, length)) return false

		progress?.()
	}

	return true
}

// A copied entry: its directory record as the original's, and its raw bytes in the written file (local header,
// data and descriptor) with the CRC its source computed of them where they live (WriteZipReport.copies).
async function sameAsRead(
	entry: ZipEntry,
	output: OpcPackage,
	copy: ZipEntry,
	read: { readonly crc: number; readonly length: number },
	progress: (() => void) | undefined
): Promise<boolean> {
	if (
		entry.method !== copy.method ||
		entry.crc32 !== copy.crc32 ||
		entry.compressedSize !== copy.compressedSize ||
		entry.uncompressedSize !== copy.uncompressedSize ||
		copy.localHeaderOffset + read.length > output.zip.source.size
	) {
		return false
	}

	let crc = 0

	for (let offset = 0; offset < read.length; offset += CHUNK) {
		crc = crc32(await output.zip.source.read(copy.localHeaderOffset + offset, Math.min(CHUNK, read.length - offset)), crc)
		progress?.()
	}

	return crc === read.crc
}

// The zip half of gate 2: `written` (random access to the produced file) holds every original entry as its
// original compressed bytes, every changed part with the CRC of its planned bytes (`crcs`, by partKey), and
// nothing else; a removed entry (`removed`, by partKey) is not there.
export async function verifyDocxZip(
	original: { readonly opc: OpcPackage },
	written: ByteSource,
	save: {
		readonly crcs: ReadonlyMap<string, number>
		readonly removed?: ReadonlySet<string>
		// What the writer read of each copied entry (WriteZipReport.copies): the copy is checked against it in
		// the written file, with the original's directory record, instead of reading the original again.
		readonly copies?: ReadonlyMap<string, { readonly crc: number; readonly length: number }>
		// Called as the check goes (a worker tells its page it is still at work).
		readonly progress?: () => void
	}
): Promise<void> {
	let output: OpcPackage

	try {
		output = await OpcPackage.open(written)
	} catch (error) {
		return fail("invalidPackage", `The written file does not open: ${error instanceof Error ? error.message : String(error)}`)
	}

	const entries = new Map(output.zip.entries.map(entry => [entry.name, entry]))
	const seen = new Set<string>()

	for (const entry of original.opc.zip.entries) {
		const key = partKey(entry.name)
		const copy = entries.get(entry.name)

		seen.add(entry.name)

		if (save.removed?.has(key) === true) {
			if (copy !== undefined) fail("untouchedChanged", `${entry.name} is still there`)

			continue
		}

		if (copy === undefined) fail("untouchedChanged", `${entry.name} is missing`)

		const crc = save.crcs.get(key)

		if (crc !== undefined) {
			if (copy.crc32 !== crc) fail("untouchedChanged", `${entry.name} is not what was written`)

			continue
		}

		const read = save.copies?.get(entry.name)
		const same =
			read === undefined
				? await sameRaw(original.opc, entry, output, copy, save.progress)
				: await sameAsRead(entry, output, copy, read, save.progress)

		if (!same) fail("untouchedChanged", `${entry.name} changed`)
	}

	for (const entry of output.zip.entries) {
		if (seen.has(entry.name)) continue

		const crc = save.crcs.get(partKey(entry.name))

		if (crc === undefined) fail("untouchedChanged", `Unexpected entry ${entry.name}`)
		if (entry.crc32 !== crc) fail("untouchedChanged", `${entry.name} is not what was written`)
	}

	const names = new Set(output.zip.entries.map(entry => partKey(entry.name)))

	for (const key of save.crcs.keys()) if (!names.has(key)) fail("untouchedChanged", `${key} was not written`)
}

async function crcOf(bytes: PartBytes): Promise<number> {
	if (bytes instanceof Uint8Array) return crc32(bytes)

	let crc = 0

	for await (const chunk of bytes.chunks()) crc = crc32(chunk, crc)

	return crc
}

// partKey -> CRC-32 of each changed part's new bytes: what the zip half of gate 2 expects in the archive.
export async function changedCrcs(changes: DocxPartChanges): Promise<Map<string, number>> {
	const crcs = new Map<string, number>()

	for (const change of [...changes.replace, ...changes.add]) crcs.set(partKey(change.name), change.crc32 ?? (await crcOf(change.bytes)))

	return crcs
}

// Raw deflate of pieces through the platform's CompressionStream, never joining them; undefined where there
// is none.
export function nativeDeflateChunks(): ((chunks: AsyncIterable<Uint8Array>) => Promise<Uint8Array>) | undefined {
	const Compression = Reflect.get(globalThis, "CompressionStream") as (new (format: string) => DeflateStream) | undefined
	const Body = Reflect.get(globalThis, "Response") as (new (body: unknown) => { arrayBuffer(): Promise<ArrayBuffer> }) | undefined

	if (typeof Compression !== "function" || typeof Body !== "function" || nativeDeflate() === undefined) return undefined

	return async chunks => {
		const stream = new Compression("deflate-raw")
		const writer = stream.writable.getWriter()
		const result = new Body(stream.readable).arrayBuffer()

		// Read whatever happens below, so a failed write leaves no unhandled rejection.
		result.catch(() => undefined)

		try {
			for await (const chunk of chunks) await writer.write(chunk)

			await writer.close()
		} catch (error) {
			await writer.abort(error).catch(() => undefined)

			throw error
		}

		return new Uint8Array(await result)
	}
}

// The zip plan of `changes` (untouched entries copied raw); a change the archive cannot take refuses the save.
export async function planChanges(original: OpcPackage, changes: DocxPartChanges): Promise<ZipPlan> {
	try {
		return await original.plan({ replace: changes.replace, add: changes.add, remove: changes.remove ?? [] })
	} catch (error) {
		if (error instanceof OpcError || error instanceof ZipError) throw new DocxSaveError("invalidPackage", error.message)

		throw error
	}
}

// Writes `changes` into the original archive and checks the result (the zip half of gate 2). Refusals are
// DocxSaveErrors with gate 2's reasons.
export async function zipDocxSave(original: OpcPackage, changes: DocxPartChanges, options: DocxZipOptions = {}): Promise<DocxZipResult> {
	let started = Date.now()
	const plan = await planChanges(original, changes)
	const memory = options.sink === undefined ? memorySink() : undefined
	const sink = options.sink ?? memory

	if (sink === undefined) throw new DocxSaveError("invalidPackage", "No sink")

	const deflate = options.zip?.deflate ?? nativeDeflate()
	const deflateChunks = options.zip?.deflateChunks ?? nativeDeflateChunks()
	const report = await writeZip(plan, original.zip, sink, {
		...options.zip,
		...(deflate !== undefined ? { deflate } : {}),
		...(deflateChunks !== undefined ? { deflateChunks } : {})
	})
	// What the archive must hold: each changed part's CRC as checked (PartChange.crc32, computed where the
	// part checks ran), else as the writer was fed it. What the writer was fed must be what was checked.
	const crcs = new Map<string, number>()
	const removed = new Set((changes.remove ?? []).map(partKey))

	for (const change of [...changes.replace, ...changes.add]) {
		const entry = plan.items.find(item => item.kind === "data" && partKey(item.name) === partKey(change.name))
		const fed = entry?.kind === "data" ? report.crcs.get(entry.name) : undefined
		if (change.crc32 === undefined && options.requireChecksums === true)
			throw new DocxSaveError("untouchedChanged", `${change.name} came without the CRC it was checked with`)

		const expected = change.crc32 ?? fed

		if (expected === undefined || fed !== expected)
			throw new DocxSaveError("untouchedChanged", `${change.name} is not what was checked`)

		crcs.set(partKey(change.name), expected)
	}

	const zipMs = Date.now() - started
	const bytes = memory?.bytes()

	started = Date.now()

	const source = bytes !== undefined ? bytesSource(bytes) : await options.readBack?.()

	if (source === undefined) throw new DocxSaveError("invalidPackage", "The written file cannot be read back")

	try {
		await verifyDocxZip(
			{ opc: original },
			source,
			options.progress === undefined
				? { crcs, removed, copies: report.copies }
				: { crcs, removed, copies: report.copies, progress: options.progress }
		)
	} catch (error) {
		if (error instanceof DocxVerifyError) throw new DocxSaveError(error.reason, error.message)

		throw error
	}

	return {
		bytes,
		changedParts: [...changes.replace, ...changes.add].map(change => change.name),
		zipMs,
		verifyZipMs: Date.now() - started
	}
}
