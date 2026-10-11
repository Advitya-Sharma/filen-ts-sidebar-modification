import { createHash } from "node:crypto"
import { readdirSync, readFileSync } from "node:fs"
import { createRequire } from "node:module"
import { dirname, join } from "node:path"
import { describe, expect, it } from "vitest"
import manifest from "@/vendor/wllama/manifest.json"

// The assistant's engine is built by scripts/wllama/build-wllama.sh and committed with the manifest it
// writes. These checks tie the committed files to that record, so an edit to the script or the patch that
// skips the rebuild, or a binary swapped by hand, fails here; `build-wllama.sh --check` proves the bytes
// themselves come from the script.

const SCRIPTS = new URL("../../scripts/wllama/", import.meta.url)
const VENDOR = new URL("../vendor/wllama/", import.meta.url)
const PACKAGE = dirname(createRequire(import.meta.url).resolve("@wllama/wllama/package.json"))

function sha256(bytes: Buffer): string {
	return createHash("sha256").update(bytes).digest("hex")
}

interface GlueField {
	type: string
	name: string
	isNullable: boolean
}

// The load request's fields in the engine's own order (GLUE_FIELD lines of its struct).
function engineLoadFields(header: string): GlueField[] {
	const body = /^struct glue_msg_load_req\n\{\n([\s\S]*?)^\};/m.exec(header)?.[1] ?? ""

	return [...body.matchAll(/^\s*GLUE_FIELD(_NULLABLE)?\((\w+), (\w+)\)/gm)].map(([, nullable, type, name]) => ({
		type: type ?? "",
		name: name ?? "",
		isNullable: nullable !== undefined
	}))
}

describe("vendored wllama engine", () => {
	it("is the binary the manifest records", () => {
		const wasm = readFileSync(new URL("wllama.wasm", VENDOR))

		expect({ bytes: wasm.length, sha256: sha256(wasm) }).toEqual(manifest.output["wllama.wasm"])
		expect(manifest.output["wllama.wasm"].sha256).not.toBe(manifest.upstream["wllama.wasm"])
	})

	it("was built from the build script and patch as they are now", () => {
		for (const [name, hash] of Object.entries(manifest.sources)) {
			expect(sha256(readFileSync(new URL(name, SCRIPTS))), name).toBe(hash)
		}
	})

	it("carries the license files the build copied, and no others", () => {
		const files = readdirSync(new URL("licenses/", VENDOR)).sort()

		expect(files).toEqual(Object.keys(manifest.licenses).sort())

		for (const [name, hash] of Object.entries(manifest.licenses)) {
			expect(sha256(readFileSync(new URL(`licenses/${name}`, VENDOR))), name).toBe(hash)
		}
	})

	it("is built from the wllama release the JavaScript comes from", () => {
		const installed = JSON.parse(readFileSync(join(PACKAGE, "package.json"), "utf8")) as { version: string }

		expect(installed.version).toBe(manifest.wllama.version)
	})

	it("reads the load request in the field order the patched JavaScript writes it", () => {
		// The protocol is positional: a field the two sides list differently shifts every later one.
		const patch = readFileSync(new URL("glue.patch", SCRIPTS), "utf8")
		const added = [...patch.matchAll(/^\+\s*GLUE_FIELD(_NULLABLE)?\((\w+), (\w+)\)/gm)].map(([, nullable, type, name]) => ({
			type: type ?? "",
			name: name ?? "",
			isNullable: nullable !== undefined
		}))

		// Appended at the end of the struct, after the fields the release has.
		expect(added.length).toBeGreaterThan(0)
		expect(patch).toMatch(/^\+\s*GLUE_FIELD[^\n]*\n \};$/m)

		const engine = [...engineLoadFields(readFileSync(join(PACKAGE, "cpp/glue.hpp"), "utf8")), ...added]

		expect(engine.length).toBeGreaterThan(added.length)

		// Both readable builds are patched; the minified one is not, and nothing may load it (below).
		for (const bundle of ["esm/index.js", "esm/index.cjs"]) {
			const script = readFileSync(join(PACKAGE, bundle), "utf8")
			const prototype = /"load_req": \{[\s\S]*?"fields": (\[[\s\S]*?\n {4}\])/.exec(script)?.[1] ?? "[]"

			expect(JSON.parse(prototype), bundle).toEqual(engine)
		}
	})

	it("is driven only by the patched JavaScript build", () => {
		const imports = new Set<string>()
		const sources = readdirSync(new URL("..", import.meta.url), { recursive: true, encoding: "utf8" }).filter(
			file => /\.tsx?$/.test(file) && !file.startsWith("tests/")
		)

		for (const file of sources) {
			const text = readFileSync(new URL(`../${file}`, import.meta.url), "utf8")

			for (const [, specifier] of text.matchAll(/from "(@wllama\/wllama[^"]*)"/g)) {
				imports.add(specifier ?? "")
			}
		}

		expect([...imports]).toEqual(["@wllama/wllama/esm/index.js"])
	})
})
