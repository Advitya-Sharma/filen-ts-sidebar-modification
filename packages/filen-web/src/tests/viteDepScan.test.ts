import { describe, expect, it } from "vitest"
import { readFileSync, readdirSync } from "node:fs"
import { join, matchesGlob, relative, resolve } from "node:path"
import type { UserConfig } from "vite"

// The dev server's dependency scan follows imports from its entries but treats a `?worker` import as external.
// A worker (or the service worker) whose module graph is not an entry has its dependencies found only when
// it first runs, and that late re-optimization reloads the page under it (the assistant's model download).

const ROOT = resolve(import.meta.dirname, "../..")

function sourceFiles(directory: string): string[] {
	return readdirSync(directory, { withFileTypes: true, recursive: true })
		.filter(entry => entry.isFile() && /\.tsx?$/.test(entry.name) && !entry.name.includes(".test."))
		.map(entry => join(entry.parentPath, entry.name))
}

// Every module the app starts as a worker, as a path from the package root.
function workerEntries(): string[] {
	const found = new Set<string>()

	for (const file of sourceFiles(join(ROOT, "src"))) {
		for (const match of readFileSync(file, "utf8").matchAll(/from\s+"([^"]+)\?worker"/g)) {
			const specifier = match[1] ?? ""

			found.add(specifier.startsWith("@/") ? `src/${specifier.slice(2)}` : relative(ROOT, resolve(file, "..", specifier)))
		}
	}

	// lib/sw/register.ts registers the service worker's source module in development.
	const register = readFileSync(join(ROOT, "src/lib/sw/register.ts"), "utf8")
	const serviceWorker = /PROD \? "[^"]+" : "\/([^"]+)"/.exec(register)?.[1]

	if (serviceWorker !== undefined) {
		found.add(serviceWorker)
	}

	return [...found]
}

async function devConfig(): Promise<UserConfig> {
	const config = (await import("../../vite.config")).default

	return config({ mode: "development", command: "serve" })
}

describe("vite dependency scan", () => {
	it("scans every worker entry and the service worker, plus the page", async () => {
		const entries = (await devConfig()).optimizeDeps?.entries
		const patterns = typeof entries === "string" ? [entries] : (entries ?? [])
		const workers = workerEntries()

		expect(workers).toEqual(expect.arrayContaining(["src/workers/assistant.worker.ts", "src/sw/sw.ts"]))
		expect(patterns).toContain("index.html")

		for (const worker of workers) {
			expect(
				patterns.some(pattern => matchesGlob(worker, pattern)),
				`${worker} is not a dependency-scan entry`
			).toBe(true)
		}
	})
})
