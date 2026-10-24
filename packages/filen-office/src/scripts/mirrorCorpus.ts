import { execFileSync } from "node:child_process"
import { mkdtempSync, rmSync, symlinkSync } from "node:fs"
import { readFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join, resolve } from "node:path"
import { fileURLToPath } from "node:url"
import { parseArgs } from "node:util"
import type { CorpusManifest } from "./corpusManifest"
import { isVerified, MIRROR_TAG, MIRRORED_SOURCES } from "./fetchCorpus"

// Uploads every file of a mirrored source (fetchCorpus.ts's MIRRORED_SOURCES) that the mirror release lacks, as an
// asset named by its SHA-256, from a verified cache (`pnpm run corpus:fetch` first). A file the cache does not hold
// verified is listed and makes the run fail, after the rest are uploaded. Needs the GitHub CLI signed in
// with write access. Assets are never replaced: a name is its content.
//
//   pnpm run corpus:mirror [--dir <directory>]

const PACKAGE = fileURLToPath(new URL("../../", import.meta.url))
const REPO = "FilenCloudDienste/filen-ts"
// Files per `gh release upload` call.
const BATCH = 25

function gh(args: readonly string[]): string {
	return execFileSync("gh", ["--repo", REPO, ...args], { encoding: "utf8", stdio: ["ignore", "pipe", "inherit"] })
}

function existingAssets(): ReadonlySet<string> | null {
	try {
		return new Set(gh(["release", "view", MIRROR_TAG, "--json", "assets", "--jq", ".assets[].name"]).split("\n").filter(Boolean))
	} catch {
		return null
	}
}

async function main(): Promise<void> {
	const { values } = parseArgs({ options: { dir: { type: "string", default: join(PACKAGE, ".corpus") } } })
	const directory = resolve(values.dir)
	const manifest = JSON.parse(await readFile(join(PACKAGE, "corpus", "manifest.json"), "utf8")) as CorpusManifest
	const mirrored = []
	const unverified = []

	for (const file of manifest.files) {
		if (!MIRRORED_SOURCES.has(file.source)) continue

		if (await isVerified(join(directory, file.path), file)) mirrored.push(file)
		else unverified.push(file.path)
	}

	let assets = existingAssets()

	if (assets === null) {
		// Kept off "latest": the repository's releases are the apps'.
		gh([
			"release",
			"create",
			MIRROR_TAG,
			"--title",
			"Office test corpus mirror",
			"--latest=false",
			"--notes",
			"Pinned copies of the office test corpus files whose sources replace them in place, each named by its SHA-256. " +
				"Sources, licences and attribution: packages/filen-office/corpus/README.md and manifest.json."
		])
		assets = new Set()
	}

	const missing = [...new Map(mirrored.filter(file => !assets.has(file.sha256)).map(file => [file.sha256, file])).values()]
	// Uploaded under the asset name, through links named by the hash.
	const staging = mkdtempSync(join(tmpdir(), "filen-office-mirror-"))

	try {
		for (let start = 0; start < missing.length; start += BATCH) {
			const batch = missing.slice(start, start + BATCH).map(file => {
				const link = join(staging, file.sha256)

				symlinkSync(join(directory, file.path), link)

				return link
			})

			gh(["release", "upload", MIRROR_TAG, ...batch])
			console.log(`mirror: ${Math.min(start + BATCH, missing.length)}/${missing.length} uploaded`)
		}
	} finally {
		rmSync(staging, { recursive: true, force: true })
	}

	console.log(`mirror: ${mirrored.length} files, ${missing.length} uploaded, ${mirrored.length - missing.length} already mirrored`)

	if (unverified.length > 0) {
		console.error(`not verified in ${directory}, so not mirrored:\n${unverified.join("\n")}`)
		process.exitCode = 1
	}
}

if (process.argv[1] === fileURLToPath(import.meta.url)) await main()
