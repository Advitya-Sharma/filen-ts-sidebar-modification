import { createHash } from "node:crypto"
import { createReadStream, existsSync } from "node:fs"
import { copyFile, mkdir, open, readFile, readdir, rename, rm, stat } from "node:fs/promises"
import { dirname, join, resolve } from "node:path"
import { fileURLToPath } from "node:url"
import { parseArgs } from "node:util"
import { FAMILIES, familyOf, unavailableAllowed, type CorpusFamily, type CorpusFile, type CorpusManifest } from "./corpusManifest"

// Downloads the real-world corpus listed in corpus/manifest.json into a cache directory (default .corpus/).
// Resumable: files already present with the right SHA-256 are kept. Every download is verified before it is
// renamed into place, so the cache never holds an unverified file. Files no longer in the manifest are removed.
// Exits non-zero when more than a few files are unavailable (any with --strict). --check-remote also asks every
// source of an already cached file whether it is still there, without downloading it. --family limits the
// download to some families (see FAMILIES); files of other families already cached stay.
//
//   pnpm run corpus:fetch [--strict] [--check-remote] [--dir <directory>] [--family pptx,ppt]

const PACKAGE = fileURLToPath(new URL("../../", import.meta.url))
const MANIFEST = join(PACKAGE, "corpus", "manifest.json")
const README = join(PACKAGE, "corpus", "README.md")
// At most CONCURRENCY requests in flight, fewer to a host that rate-limits: GOV.UK blocks an address that asks too fast.
const CONCURRENCY = 8
const HOST_CONCURRENCY: Readonly<Record<string, number>> = { "assets.publishing.service.gov.uk": 3 }
const ATTEMPTS = 4
const TIMEOUT_MS = 120_000
const RETRYABLE = new Set([408, 425, 429, 500, 502, 503, 504])

type Failure = { kind: "missing" | "changed" | "failed"; detail: string }

class Changed extends Error {}

function sleep(ms: number): Promise<void> {
	return new Promise(done => setTimeout(done, ms))
}

function backoff(attempt: number, retryAfter: string | null): number {
	const seconds = retryAfter === null ? NaN : Number(retryAfter)

	if (Number.isFinite(seconds) && seconds >= 0) return Math.min(seconds, 60) * 1000

	return 1000 * 2 ** (attempt - 1) * (1 + Math.random() * 0.3)
}

async function sha256File(path: string): Promise<string> {
	const hash = createHash("sha256")

	for await (const chunk of createReadStream(path)) hash.update(chunk as Buffer)

	return hash.digest("hex")
}

async function isVerified(path: string, file: CorpusFile): Promise<boolean> {
	try {
		return (await stat(path)).size === file.size && (await sha256File(path)) === file.sha256
	} catch {
		return false
	}
}

// Streams one response body to `part`, hashing on the way; throws Changed as soon as it cannot match.
async function receive(body: ReadableStream<Uint8Array>, part: string, file: CorpusFile): Promise<void> {
	const hash = createHash("sha256")
	const handle = await open(part, "w")
	let received = 0

	try {
		for await (const chunk of body) {
			received += chunk.byteLength

			if (received > file.size) throw new Changed(`more than ${file.size} bytes`)

			hash.update(chunk)
			await handle.write(chunk)
		}
	} finally {
		await handle.close()
	}

	if (received !== file.size) throw new Changed(`${received} bytes instead of ${file.size}`)

	const digest = hash.digest("hex")

	if (digest !== file.sha256) throw new Changed(`SHA-256 ${digest}`)
}

// null once the file is verified in place.
async function download(file: CorpusFile, path: string): Promise<Failure | null> {
	const part = `${path}.part`
	let detail = ""

	for (let attempt = 1; ; attempt++) {
		let retryAfter: string | null = null

		try {
			const response = await fetch(file.url, { signal: AbortSignal.timeout(TIMEOUT_MS) })

			if (response.ok && response.body !== null) {
				await receive(response.body, part, file)
				await rename(part, path)

				return null
			}

			await response.body?.cancel()

			detail = `HTTP ${response.status}`

			if (response.status === 404 || response.status === 410) return { kind: "missing", detail }
			if (!RETRYABLE.has(response.status)) return { kind: "failed", detail }

			retryAfter = response.headers.get("retry-after")
		} catch (error) {
			if (error instanceof Changed) return { kind: "changed", detail: error.message }

			detail =
				error instanceof Error
					? error.cause instanceof Error
						? `${error.message}: ${error.cause.message}`
						: error.message
					: String(error)
		} finally {
			await rm(part, { force: true })
		}

		if (attempt === ATTEMPTS) return { kind: "failed", detail }

		await sleep(backoff(attempt, retryAfter))
	}
}

// Whether the source still serves the file, by a HEAD request (no body); null when it does.
async function probe(file: CorpusFile): Promise<Failure | null> {
	let detail = ""

	for (let attempt = 1; ; attempt++) {
		let retryAfter: string | null = null

		try {
			const response = await fetch(file.url, { method: "HEAD", signal: AbortSignal.timeout(30_000) })
			const length = response.headers.get("content-length")

			if (response.ok) {
				// A compressed transfer reports the compressed length.
				if (length !== null && response.headers.get("content-encoding") === null && Number(length) !== file.size) {
					return { kind: "changed", detail: `${length} bytes instead of ${file.size}` }
				}

				return null
			}

			detail = `HTTP ${response.status}`

			if (response.status === 404 || response.status === 410) return { kind: "missing", detail }
			if (!RETRYABLE.has(response.status)) return { kind: "failed", detail }

			retryAfter = response.headers.get("retry-after")
		} catch (error) {
			detail = error instanceof Error ? error.message : String(error)
		}

		if (attempt === ATTEMPTS) return { kind: "failed", detail }

		await sleep(backoff(attempt, retryAfter))
	}
}

// Runs `work` over the files, in order per host, with at most CONCURRENCY in flight and HOST_CONCURRENCY per host.
async function pool(files: readonly CorpusFile[], work: (file: CorpusFile) => Promise<void>): Promise<void> {
	const byHost = Map.groupBy(files, file => new URL(file.url).host)
	let free = CONCURRENCY
	const waiting: (() => void)[] = []
	const acquire = async (): Promise<void> => {
		if (free > 0) {
			free--

			return
		}

		await new Promise<void>(granted => waiting.push(granted))
	}
	const release = (): void => {
		const next = waiting.shift()

		if (next === undefined) free++
		else next()
	}

	await Promise.all(
		[...byHost].flatMap(([host, queue]) => {
			let next = 0

			return Array.from({ length: Math.min(HOST_CONCURRENCY[host] ?? CONCURRENCY, queue.length) }, async () => {
				while (next < queue.length) {
					const file = queue[next++] as CorpusFile

					await acquire()

					try {
						await work(file)
					} finally {
						release()
					}
				}
			})
		})
	)
}

function families(option: string | undefined): ReadonlySet<CorpusFamily> | undefined {
	if (option === undefined) return undefined

	const names = option.split(",").map(name => name.trim())
	const unknown = names.filter(name => !Object.hasOwn(FAMILIES, name))

	if (unknown.length > 0) {
		annotate("error", `unknown --family ${unknown.join(", ")}; families: ${Object.keys(FAMILIES).join(", ")}`)
		process.exit(2)
	}

	return new Set(names as CorpusFamily[])
}

// Removes what the manifest no longer lists, only inside a directory this script populated before.
async function prune(directory: string, expected: ReadonlySet<string>): Promise<number> {
	if (!existsSync(join(directory, "manifest.json"))) return 0

	let removed = 0

	for (const format of await readdir(directory, { withFileTypes: true })) {
		if (!format.isDirectory()) continue

		for (const entry of await readdir(join(directory, format.name), { withFileTypes: true })) {
			const relative = `${format.name}/${entry.name}`

			if (!entry.isFile() || expected.has(relative)) continue

			await rm(join(directory, relative), { force: true })
			removed++
		}
	}

	return removed
}

function annotate(level: "warning" | "error", message: string): void {
	console.log(process.env["GITHUB_ACTIONS"] === "true" ? `::${level}::${message}` : `${level}: ${message}`)
}

async function main(): Promise<void> {
	const { values } = parseArgs({
		options: {
			dir: { type: "string", default: join(PACKAGE, ".corpus") },
			strict: { type: "boolean", default: false },
			"check-remote": { type: "boolean", default: false },
			family: { type: "string" }
		}
	})
	const only = families(values.family)
	const directory = resolve(values.dir)
	const manifest = JSON.parse(await readFile(MANIFEST, "utf8")) as CorpusManifest
	const started = performance.now()
	const unavailable = new Map<CorpusFile, Failure>()
	const remote = new Map<CorpusFile, Failure>()
	const cached: CorpusFile[] = []
	let downloaded = 0
	let bytes = 0

	await mkdir(directory, { recursive: true })

	const removed = await prune(directory, new Set(manifest.files.map(file => file.path)))
	const wanted = manifest.files.filter(file => {
		const family = familyOf(file.path)

		if (family === undefined) throw new Error(`${file.path}: no corpus family for this extension`)

		return only === undefined || only.has(family)
	})

	// Largest first, so one big file does not trail at the end.
	await pool(
		[...wanted].sort((a, b) => b.size - a.size),
		async file => {
			const path = join(directory, file.path)

			if (await isVerified(path, file)) {
				cached.push(file)

				return
			}

			await mkdir(dirname(path), { recursive: true })
			await rm(path, { force: true })

			const failure = await download(file, path)

			if (failure === null) {
				downloaded++
				bytes += file.size
			} else {
				unavailable.set(file, failure)
			}
		}
	)

	if (values["check-remote"]) {
		await pool(cached, async file => {
			const failure = await probe(file)

			if (failure !== null) remote.set(file, failure)
		})
	}

	await copyFile(MANIFEST, join(directory, "manifest.json"))
	await copyFile(README, join(directory, "README.md"))

	const seconds = ((performance.now() - started) / 1000).toFixed(1)

	console.log(
		`corpus: ${wanted.length - unavailable.size}/${wanted.length} files verified in ${directory} ` +
			`(${cached.length} already cached, ${downloaded} downloaded, ${(bytes / 1e6).toFixed(1)} MB, ${removed} stale removed) in ${seconds} s`
	)

	for (const [file, failure] of unavailable) {
		annotate("warning", `${file.path} unavailable (${failure.kind}: ${failure.detail}) from ${file.url}`)
	}

	for (const [file, failure] of remote) {
		annotate("warning", `${file.path} cached, but its source is ${failure.kind} (${failure.detail}): ${file.url}`)
	}

	const allowed = values.strict ? 0 : unavailableAllowed(wanted.length)
	const problems = unavailable.size + (values.strict ? remote.size : 0)

	if (problems > allowed) {
		annotate("error", `${problems} corpus files unavailable or changed at their source (${allowed} allowed)`)
		process.exitCode = 1
	}
}

if (process.argv[1] === fileURLToPath(import.meta.url)) await main()
