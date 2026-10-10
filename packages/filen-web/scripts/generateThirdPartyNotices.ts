import { execFileSync } from "node:child_process"
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs"
import { homedir } from "node:os"
import { join } from "node:path"
import {
	describeCargoCrate,
	describedNotice,
	finalizeNotices,
	listInstalledNpm,
	readJson,
	repositoryOf,
	spdxOf,
	vendoredNotice,
	type CollectedNotice,
	type InstalledNpmPackage
} from "@filen/shared/tooling"

/**
 * Emits the third-party attribution payload the Open source licenses dialog renders.
 *
 * MIT, BSD and Apache-2.0 all require the license text AND the copyright notice to accompany the
 * distribution, so listing package names and SPDX ids would not discharge the obligation — the text has
 * to ship inside the app.
 *
 * Shipping it verbatim per package would be megabytes, because an MIT file is identical everywhere
 * except its copyright line. So the copyright lines are lifted out and stored per package, and what
 * remains — the boilerplate — is deduplicated across every package that shares it. Nothing is
 * summarised or paraphrased; a rendered notice is copyright line plus verbatim boilerplate.
 *
 * A package whose license file cannot be found keeps its SPDX id and repository URL and points at no
 * text. Substituting another package's text of the same license would attribute the wrong copyright
 * holder, which is worse than an honest gap.
 *
 * Three sources reach the browser: the npm tree (minus dev/optional entries, which are build tooling),
 * the Rust crates the SDK wasm is compiled from, and components vendored into this package as built
 * output, described from the manifest their build script writes. @filen/sdk-rs is a registry package and
 * is described here; @filen/shared is a workspace member, so pnpm does not report it and it is correctly
 * absent — first-party code needs no attribution.
 *
 * The licensing core is shared with filen-mobile's generator through @filen/shared/tooling. Unlike
 * mobile, the npm dedup key here is `name@version` rather than the bare name (this tree ships 30
 * packages at two versions, and a bare-name key would silently describe only one of each pair).
 *
 * CANNOT RUN IN CI: it needs a filen-rs checkout and the local cargo caches. Re-run after any
 * @filen/sdk-rs bump; the payload's exported SDK version is checked against the installed tree by
 * src/tests/thirdPartyNotices.test.ts, so a stale payload fails there rather than shipping quietly.
 *
 * Run from the package root: node --experimental-strip-types scripts/generateThirdPartyNotices.ts
 *   --filen-rs=<path>   filen-rs checkout to read the crate lockfile from
 *   --allow-untagged    read the checkout's working tree instead of the release tag (local experiment
 *                       only — the payload is then stamped worktree@<sha> and fails the payload test)
 */

const OUTPUT = "src/features/settings/thirdPartyNotices.gen.ts"

type Ecosystem = "npm" | "rust" | "vendored"

type Collected = CollectedNotice<Ecosystem>

function requireJson(path: string): Record<string, unknown> {
	const json = readJson(path)

	if (json === null) {
		throw new Error(`could not read ${path}`)
	}

	return json
}

/**
 * The npm packages that reach a browser.
 *
 * Identity is `name@version`, not the bare name: several packages ship at two versions in this tree and
 * both are distributed, so a name-keyed dedup would describe only one of each pair.
 */
function collectNpm(installed: InstalledNpmPackage[]): Collected[] {
	const seen = new Set<string>()
	const entries: Collected[] = []

	for (const { name, version, dir } of installed) {
		const id = `${name}@${version}`

		if (seen.has(id) || dir === null) {
			continue
		}

		const manifest = readJson(join(dir, "package.json"))

		// Marked seen only once actually described: an id pnpm reports but no directory can supply stays
		// undescribed and fails the payload's drift guard loudly rather than silently.
		if (!manifest) {
			continue
		}

		const manifestVersion = manifest["version"]

		if (typeof manifestVersion === "string" && manifestVersion !== version) {
			throw new Error(`${id}: pnpm says ${version}, ${dir}/package.json says ${manifestVersion} — run pnpm install`)
		}

		seen.add(id)

		entries.push(
			describedNotice({
				name,
				version,
				license: spdxOf(manifest["license"] ?? manifest["licenses"]),
				ecosystem: "npm",
				dir,
				repository: repositoryOf(manifest["repository"])
			})
		)
	}

	return entries
}

function argValue(flag: string): string | undefined {
	return process.argv
		.slice(2)
		.find(arg => arg.startsWith(`${flag}=`))
		?.slice(flag.length + 1)
}

function git(checkout: string, args: string[]): string {
	return execFileSync("git", ["-C", checkout, ...args], { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 })
}

/**
 * Which filen-rs checkout the crate list comes from, and which of the three candidates supplied it.
 *
 * The monorepo's `packages/filen-mobile/filen-rs` submodule is deliberately NOT a candidate: it is
 * pinned by mobile's native build on mobile's schedule and routinely sits at a different SDK release,
 * so falling through to it would attribute one release's crates to another's wasm.
 */
function resolveCheckout(): { path: string; source: string } {
	const candidates: { path: string; source: string }[] = [
		{ path: argValue("--filen-rs") ?? "", source: "--filen-rs" },
		{ path: process.env["FILEN_RS_PATH"] ?? "", source: "FILEN_RS_PATH" },
		{ path: "../../../filen-rs", source: "sibling" }
	]

	for (const candidate of candidates) {
		if (candidate.path.length > 0 && existsSync(join(candidate.path, "Cargo.lock"))) {
			return candidate
		}
	}

	throw new Error(
		`no filen-rs checkout with a Cargo.lock — tried ${candidates.map(c => `${c.source}=${c.path.length > 0 ? c.path : "(unset)"}`).join(", ")}`
	)
}

/** The crate sources cargo keeps on this machine: registry crates plus git-dependency checkouts. */
function crateSourceRoots(): string[] {
	const roots: string[] = []

	for (const base of [join(homedir(), ".cargo", "registry", "src"), join(homedir(), ".cargo", "git", "checkouts")]) {
		if (!existsSync(base)) {
			continue
		}

		for (const entry of readdirSync(base)) {
			roots.push(join(base, entry))
		}
	}

	return roots
}

/**
 * The Rust crates the shipped wasm is compiled from, read at the RELEASE TAG matching the installed
 * @filen/sdk-rs version.
 *
 * Reading the checkout's working tree instead would attribute whatever the branch happens to be at —
 * measurably a different crate set — and a version equality check alone cannot catch that, because the
 * working tree still declares the same version.
 */
function collectRust(expected: string): { entries: Collected[]; ref: string; source: string; crates: number } {
	const checkout = resolveCheckout()
	const tag = `filen-js@${expected}`
	const allowUntagged = process.argv.includes("--allow-untagged")
	let ref = tag

	try {
		git(checkout.path, ["rev-parse", "-q", "--verify", `refs/tags/${tag}`])
	} catch {
		if (!allowUntagged) {
			throw new Error(`${checkout.path} has no tag ${tag} — fetch it (git -C ${checkout.path} fetch --tags) or pass --allow-untagged`)
		}

		ref = `worktree@${git(checkout.path, ["rev-parse", "--short", "HEAD"]).trim()}`
	}

	const atTag = ref === tag
	const cargoToml = atTag
		? git(checkout.path, ["show", `${tag}:filen-sdk-rs/Cargo.toml`])
		: readFileSync(join(checkout.path, "filen-sdk-rs", "Cargo.toml"), "utf8")
	const cargoLock = atTag ? git(checkout.path, ["show", `${tag}:Cargo.lock`]) : readFileSync(join(checkout.path, "Cargo.lock"), "utf8")
	const declared = /^\s*version\s*=\s*"([^"]+)"/m.exec(cargoToml.split("[dependencies]")[0] ?? cargoToml)?.[1]

	if (declared !== expected) {
		throw new Error(`${checkout.path} at ${ref} declares filen-sdk-rs ${declared ?? "(none)"}, expected ${expected}`)
	}

	const roots = crateSourceRoots()
	const entries: Collected[] = []
	const byId = new Map<string, Collected>()
	let crates = 0

	for (const match of cargoLock.matchAll(/\[\[package\]\]\nname = "([^"]+)"\nversion = "([^"]+)"(?:\nsource = "([^"]+)")?/g)) {
		const [, name, version, source] = match

		// No `source` means a workspace member — Filen's own code, not a third party.
		if (!source || !name || !version) {
			continue
		}

		crates++

		// A registry crate unpacks to <root>/<name>-<version>; a git dependency to a checkout directory
		// named after the abbreviated revision from the source URL's fragment.
		const rev = /#([0-9a-f]+)$/.exec(source)?.[1]
		const dir =
			roots.map(root => join(root, `${name}-${version}`)).find(existsSync) ??
			(rev === undefined
				? undefined
				: roots
						.flatMap(root => {
							try {
								return readdirSync(root).map(entry => ({ entry, path: join(root, entry) }))
							} catch {
								return []
							}
						})
						.find(({ entry }) => rev.startsWith(entry))?.path)

		if (dir === undefined) {
			throw new Error(`no local source for crate ${name}-${version} — run cargo fetch in ${checkout.path}`)
		}

		const entry = describeCargoCrate(name, version, dir)

		// The same crate version can be locked twice — once from crates.io and once from a git fork of it.
		// Identical attribution is one notice, not two (the payload is keyed by name@version). Attribution
		// that actually differs is a real ambiguity and throws rather than silently dropping one holder.
		const previous = byId.get(`${name}@${version}`)

		if (previous !== undefined) {
			if (JSON.stringify({ ...previous, texts: [] }) !== JSON.stringify(entry)) {
				throw new Error(
					`crate ${name}-${version} is locked twice with differing attribution (${previous.license} vs ${entry.license})`
				)
			}

			continue
		}

		byId.set(`${name}@${version}`, entry)
		entries.push(entry)
	}

	return { entries, ref, source: checkout.source, crates }
}

// The one definition of "the installed SDK": what pnpm reports for the installed tree, which is also
// what the payload test reads. The symlink under node_modules is the cross-check, so an edited
// package.json with no install cannot generate a payload that only fails later.
const installed = listInstalledNpm("@filen/web")
const expectedSdk = installed.find(entry => entry.name === "@filen/sdk-rs")?.version

if (typeof expectedSdk !== "string" || expectedSdk.length === 0) {
	throw new Error("the installed tree has no @filen/sdk-rs — run `pnpm install` at the repo root")
}

const installedSdk = requireJson("node_modules/@filen/sdk-rs/package.json")["version"]

if (installedSdk !== expectedSdk) {
	throw new Error(`pnpm says @filen/sdk-rs ${expectedSdk}, node_modules says ${String(installedSdk)} — run pnpm install`)
}

function stringField(record: Record<string, unknown>, key: string, path: string): string {
	const value = record[key]

	if (typeof value !== "string" || value.length === 0) {
		throw new Error(`${path}: "${key}" is missing`)
	}

	return value
}

function stringsField(record: Record<string, unknown>, key: string, path: string): string[] {
	const value = record[key]

	if (!Array.isArray(value) || value.length === 0 || !value.every(item => typeof item === "string")) {
		throw new Error(`${path}: "${key}" is missing`)
	}

	return value
}

function objectField(record: Record<string, unknown>, key: string, path: string): Record<string, unknown> {
	const value = record[key]

	if (typeof value !== "object" || value === null || Array.isArray(value)) {
		throw new Error(`${path}: "${key}" is missing`)
	}

	return Object.fromEntries(Object.entries(value))
}

/**
 * The FFmpeg decoders module, built to WebAssembly by scripts/codecs/build-ffmpeg-decoders.sh. What it
 * contains is read from the manifest the build writes, so the notices cannot drift from what is shipped:
 *
 * - FFmpeg configured without --enable-gpl or --enable-nonfree, LGPL-2.1-or-later
 *   (https://github.com/FFmpeg/FFmpeg/blob/master/LICENSE.md). The text is FFmpeg's own COPYING.LGPLv2.1,
 *   kept verbatim beside this script. The corresponding source is named by path and SHA-256: the commit
 *   that ships a build is not known when the build is made.
 * - Emscripten's runtime and emmalloc (MIT OR NCSA) and musl (MIT), statically linked in; their license
 *   files are copied next to the module by the build. compiler-rt's builtins are linked too, but under
 *   LLVM's exception to Apache-2.0 an object form embedding them owes no attribution.
 */
const FFMPEG_DIR = "src/vendor/ffmpeg-decoders"
const FFMPEG_MANIFEST = `${FFMPEG_DIR}/manifest.json`
const CODECS_SOURCE = "packages/filen-web/scripts/codecs"

function ffmpegModuleNotices(): Collected[] {
	const manifest = requireJson(FFMPEG_MANIFEST)
	const ffmpeg = objectField(manifest, "ffmpeg", FFMPEG_MANIFEST)
	const emscripten = objectField(manifest, "emscripten", FFMPEG_MANIFEST)
	const sources = objectField(manifest, "sources", FFMPEG_MANIFEST)
	const tag = stringField(ffmpeg, "tag", FFMPEG_MANIFEST)
	const emcc = stringField(manifest, "emcc", FFMPEG_MANIFEST)
	const linkedInto = `Statically linked into the FFmpeg decoders module (${FFMPEG_DIR}/decoders.worker.js), built with Emscripten ${emcc}.`
	const sourceList = ["build-ffmpeg-decoders.sh", "bridge.c", "decoder-worker.js"]
		.map(name => `${CODECS_SOURCE}/${name} (SHA-256 ${stringField(sources, name, FFMPEG_MANIFEST)})`)
		.join(", ")

	return [
		vendoredNotice(
			{
				name: "FFmpeg",
				version: tag.replace(/^n/, ""),
				license: "LGPL-2.1-or-later",
				copyright: ["Copyright (c) the FFmpeg developers"],
				repository: "https://git.ffmpeg.org/ffmpeg.git",
				terms: [
					[
						`FFmpeg's AC-3, E-AC-3, DTS, TrueHD/MLP, ALAC and FLAC decoders (libavcodec, libavutil), built to WebAssembly from the FFmpeg ${tag} release and licensed under the GNU Lesser General Public License, version 2.1 or later.`,
						`Source: ${stringField(ffmpeg, "tarball", FFMPEG_MANIFEST)} (SHA-256 ${stringField(ffmpeg, "sha256", FFMPEG_MANIFEST)}).`,
						`Configured with: ${stringsField(ffmpeg, "configure", FFMPEG_MANIFEST).join(" ")}.`,
						`Built with Emscripten ${emcc}. The complete corresponding source of the module, in the filen-ts repository (https://github.com/FilenCloudDienste/filen-ts): ${sourceList}.`
					].join(" "),
					readFileSync("scripts/licenses/ffmpeg-COPYING.LGPLv2.1", "utf8")
				]
			},
			"vendored"
		),
		vendoredNotice(
			{
				name: "Emscripten",
				version: emcc,
				license: "MIT OR NCSA",
				copyright: ["Copyright (c) 2010-2014 Emscripten authors"],
				repository: "https://github.com/emscripten-core/emscripten",
				terms: [
					`The Emscripten runtime and emmalloc. ${linkedInto}`,
					readFileSync(`${FFMPEG_DIR}/licenses/emscripten-LICENSE`, "utf8")
				]
			},
			"vendored"
		),
		vendoredNotice(
			{
				name: "musl",
				version: stringField(emscripten, "musl", FFMPEG_MANIFEST),
				license: "MIT",
				copyright: ["Copyright © 2005-2020 Rich Felker, et al."],
				repository: "https://git.musl-libc.org/cgit/musl",
				terms: [
					`The musl C library, as bundled with Emscripten. ${linkedInto}`,
					readFileSync(`${FFMPEG_DIR}/licenses/musl-COPYRIGHT`, "utf8")
				]
			},
			"vendored"
		)
	]
}

const npm = collectNpm(installed)
const rust = collectRust(expectedSdk)
const vendored = ffmpegModuleNotices()
const collected = [...npm, ...rust.entries, ...vendored]

const { texts, notices } = finalizeNotices(collected)

// No absolute path is ever written into the artifact: it is a committed source file, and a machine path
// would make regeneration produce a spurious diff on the next machine. The repository, the ref it was
// read at, and which candidate resolved carry the whole diagnostic value.
const output = `// AUTO-GENERATED by scripts/generateThirdPartyNotices.ts — do not edit.
//
// Attribution for everything bundled into the app: the npm tree minus dev/optional entries, the Rust
// crates the SDK wasm is built from, and the components vendored as built output. License terms are
// deduplicated across packages and the copyright lines kept per package, so a rendered notice is that
// package's copyright followed by the verbatim terms. An empty \`texts\` means no license file shipped
// with the package; its SPDX id and repository are given instead rather than borrowing another package's
// copyright. More than one entry means the declared license is a conjunction — every text applies.
//
// npm:  pnpm licenses list --json --prod --no-optional --filter @filen/web... (${String(npm.length)} name@version packages)
// rust: filen-rs @ ${rust.ref} (${String(rust.crates)} crates)  [resolved via: ${rust.source}]
// vendored: ${vendored.map(entry => `${entry.name} ${entry.version}`).join(", ")} (${FFMPEG_MANIFEST})
//
// Cannot run in CI: needs a filen-rs checkout and the local cargo caches. Re-run after any SDK bump.

export type ThirdPartyNotice = {
	name: string
	version: string
	license: string
	ecosystem: "npm" | "rust" | "vendored"
	copyright: string[]
	repository: string | null
	texts: number[]
}

export const THIRD_PARTY_NOTICES_SDK_VERSION = "${expectedSdk}"

export const THIRD_PARTY_NOTICES_FILEN_RS_REF = "${rust.ref}"

export const THIRD_PARTY_NOTICES_PAYLOAD = ${JSON.stringify(JSON.stringify({ texts, notices }))}
`

writeFileSync(OUTPUT, output, "utf8")

const withText = notices.filter(entry => entry.texts.length > 0).length

console.log(
	[
		`third-party notices -> ${OUTPUT}`,
		`  npm:            ${String(npm.length).padStart(4)}`,
		`  rust:           ${String(rust.entries.length).padStart(4)}  (${rust.ref}, via ${rust.source})`,
		`  vendored:       ${String(vendored.length).padStart(4)}`,
		`  total:          ${String(notices.length)}`,
		`  license texts:  ${String(texts.length)} unique (deduplicated)`,
		`  with text:      ${String(withText)}`,
		`  without text:   ${String(notices.length - withText)}`,
		`  payload:        ${(Buffer.byteLength(output, "utf8") / 1024).toFixed(0)} KB`
	].join("\n")
)
