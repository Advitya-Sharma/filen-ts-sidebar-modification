import { createHash } from "node:crypto"
import { createReadStream, readdirSync, statSync } from "node:fs"
import { basename, dirname, join } from "node:path"
import { pipeline } from "node:stream/promises"
import type { Plugin } from "vite"
import type { AssistantModelFile, AssistantModelManifest } from "../src/features/assistant/runtime/manifest"

const SPLIT_NAME = /^(.*)-(\d{5})-of-(\d{5})\.gguf$/

// The files of the model `path` names: every llama-gguf-split piece when it names one, else the file alone.
export function modelFiles(path: string): string[] {
	const match = SPLIT_NAME.exec(basename(path))

	if (match === null) {
		return [path]
	}

	const [, stem, , total] = match
	const directory = dirname(path)

	return readdirSync(directory)
		.filter(name => name.startsWith(`${stem ?? ""}-`) && name.endsWith(`-of-${total ?? ""}.gguf`))
		.sort()
		.map(name => join(directory, name))
}

async function describeFile(path: string): Promise<AssistantModelFile> {
	const hash = createHash("sha256")

	await pipeline(createReadStream(path), hash)

	return { file: basename(path), size: statSync(path).size, sha256: hash.digest("hex") }
}

// The `shards` list of a manifest (src/features/assistant/runtime/models.ts) for the model at `path`.
export async function describeModelShards(path: string): Promise<AssistantModelFile[]> {
	const shards: AssistantModelFile[] = []

	for (const file of modelFiles(path)) {
		shards.push(await describeFile(file))
	}

	return shards
}

const DEV_PREFIX = "/__assistant-model/"

// Dev only: FILEN_ASSISTANT_DEV_GGUF=/abs/path/model.gguf (or the first piece of a split model) serves that
// model under /__assistant-model/ with a manifest hashed on first request, so the runtime can be tried
// against any local GGUF. The client asks here only in dev builds without VITE_ASSISTANT_MODEL_BASE_URL.
export function assistantModelDev(): Plugin {
	const path = process.env["FILEN_ASSISTANT_DEV_GGUF"]
	let manifest: Promise<AssistantModelManifest> | null = null

	return {
		name: "filen:assistant-model-dev",
		apply: "serve",
		configureServer(server) {
			if (path === undefined || path === "") {
				return
			}

			const files = new Map(modelFiles(path).map(file => [basename(file), file]))

			server.middlewares.use((request, response, next) => {
				const url = (request.url ?? "").split("?")[0] ?? ""

				if (!url.startsWith(DEV_PREFIX)) {
					next()

					return
				}

				const name = decodeURIComponent(url.slice(DEV_PREFIX.length))

				if (name === "manifest.json") {
					manifest ??= describeModelShards(path).then(shards => ({
						id: `dev-${shards[0]?.sha256.slice(0, 12) ?? "model"}`,
						label: basename(path),
						license: { name: "local", url: "" },
						source: { repo: "local", revision: "" },
						shards
					}))

					void manifest.then(
						current => {
							response.setHeader("Content-Type", "application/json")
							response.end(JSON.stringify(current))
						},
						(e: unknown) => {
							server.config.logger.error(`assistant model: ${String(e)}`)
							response.statusCode = 500
							response.end()
						}
					)

					return
				}

				const file = files.get(name)

				if (file === undefined) {
					response.statusCode = 404
					response.end()

					return
				}

				// The runtime resumes with a single "bytes=N-" range, as it does against the real host.
				const size = statSync(file).size
				const start = Number(/^bytes=(\d+)-$/.exec(request.headers.range ?? "")?.[1] ?? 0)

				response.statusCode = start > 0 ? 206 : 200
				response.setHeader("Content-Type", "application/octet-stream")
				response.setHeader("Content-Length", String(size - start))
				createReadStream(file, { start })
					.on("error", () => {
						response.statusCode = 500
						response.end()
					})
					.pipe(response)
			})
		}
	}
}
