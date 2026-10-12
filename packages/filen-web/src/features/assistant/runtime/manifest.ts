import * as z from "zod"

// zod probes `new Function` when it builds its first object schema; the app's CSP blocks that and reports a
// violation on every page load. Under that CSP the interpreted parser runs anyway. Must run before any schema exists.
z.config({ jitless: true })

// Import-free of the app (no "@/" imports): scripts/assistant-model-manifest.ts and the dev server plugin
// read it under Node too.

// The production model host; covered by connect-src's https://*.filen.io. A build can point elsewhere with
// VITE_ASSISTANT_MODEL_BASE_URL (vite.config.ts then adds that origin to connect-src).
export const DEFAULT_MODEL_BASE_URL = "https://cdn.filen.io/gguf/"

const SHA256 = /^[0-9a-f]{64}$/

const SPLIT_NAME = /-(\d{5})-of-(\d{5})\.gguf$/

const fileSchema = z.strictObject({
	file: z.string().regex(/^[\w.-]+\.gguf$/),
	size: z.int().positive(),
	sha256: z.string().regex(SHA256)
})

const manifestSchema = z.strictObject({
	id: z.string().regex(/^[\w.-]+$/),
	label: z.string(),
	license: z.strictObject({
		name: z.string(),
		url: z.string()
	}),
	source: z.strictObject({
		repo: z.string(),
		revision: z.string()
	}),
	// The unsplit GGUF, for runtimes without wllama's 2 GB per-file limit (mobile).
	whole: fileSchema.optional(),
	// llama-gguf-split pieces in order; the web runtime downloads and loads these.
	shards: z.array(fileSchema).min(1)
})

export type AssistantModelFile = z.infer<typeof fileSchema>

export type AssistantModelManifest = z.infer<typeof manifestSchema>

// wllama copies a file into one ArrayBuffer when it cannot stream it (no JSPI), which caps a file at 2 GB.
export const MAX_SHARD_BYTES = 2 * 1024 * 1024 * 1024 - 1

// Throws on anything malformed: a split model's pieces must be named and listed in llama-gguf-split order,
// since llama.cpp finds the next piece by that name.
export function parseManifest(value: unknown): AssistantModelManifest {
	const manifest = manifestSchema.parse(value)
	const { shards } = manifest

	if (new Set(shards.map(shard => shard.file)).size !== shards.length) {
		throw new Error(`Model manifest ${manifest.id}: duplicate shard names`)
	}

	if (shards.some(shard => shard.size > MAX_SHARD_BYTES)) {
		throw new Error(`Model manifest ${manifest.id}: a shard exceeds 2 GB`)
	}

	if (shards.length > 1) {
		shards.forEach((shard, index) => {
			const match = SPLIT_NAME.exec(shard.file)

			if (match === null || Number(match[1]) !== index + 1 || Number(match[2]) !== shards.length) {
				throw new Error(`Model manifest ${manifest.id}: ${shard.file} is out of llama-gguf-split order`)
			}
		})
	}

	return manifest
}

export function manifestBytes(manifest: AssistantModelManifest): number {
	return manifest.shards.reduce((sum, shard) => sum + shard.size, 0)
}

export function isAllowedModelOrigin(url: URL): boolean {
	return url.protocol === "https:" || (url.protocol === "http:" && (url.hostname === "localhost" || url.hostname === "127.0.0.1"))
}

// The base is the directory the files sit in; only https (or a loopback dev server) is accepted, so a
// misconfigured build cannot fetch a model in the clear.
export function modelFileUrl({ baseUrl, file }: { baseUrl: string; file: AssistantModelFile }): string {
	const base = new URL(baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`)

	if (!isAllowedModelOrigin(base)) {
		throw new Error(`Model host ${base.origin} is not https`)
	}

	return new URL(encodeURIComponent(file.file), base).href
}
