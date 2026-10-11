import { createHash } from "node:crypto"
import { createReadStream, readFileSync, statSync } from "node:fs"
import { createRequire } from "node:module"
import { dirname, extname, join } from "node:path"
import type { Plugin } from "vite"
import { COI_HEADERS } from "./sdk-artifacts-plugin"

// @filen/office's bundled document fonts and their licence files, served from 'self' (CSP font-src). A
// build emits each face under /assets/ with its manifest SHA-256 in the name (cached as immutable, like
// every other asset there); `virtual:office-fonts` maps each manifest path to that URL for the viewer, whose
// worker fetches the bytes and registers them with `new FontFace`. The licence files keep their manifest
// paths under OFFICE_FONTS_BASE (unhashed, so revalidated). The dev server serves everything from
// OFFICE_FONTS_BASE.
export const OFFICE_FONTS_BASE = "/office/fonts/"

const VIRTUAL_ID = "virtual:office-fonts"
const RESOLVED_ID = `\0${VIRTUAL_ID}`

// A face's URL in a build: its manifest path with the first 8 hex digits of its SHA-256 before the extension.
function hashedPath(face: Face): string {
	const extension = extname(face.file)

	return `assets/office-fonts/${face.file.slice(0, -extension.length)}-${face.sha256.slice(0, 8)}${extension}`
}

const FONTS_DIR = dirname(createRequire(import.meta.url).resolve("@filen/office/assets/fonts/manifest.json"))
const MANIFEST = join(FONTS_DIR, "manifest.json")
const MIME: Record<string, string> = { ".woff2": "font/woff2", ".txt": "text/plain; charset=utf-8" }

interface Face {
	file: string
	bytes: number
	sha256: string
}

function record(value: unknown): Record<string, unknown> {
	return typeof value === "object" && value !== null && !Array.isArray(value) ? Object.fromEntries(Object.entries(value)) : {}
}

// Every face and each family's licence file, as paths under the fonts directory.
function readManifest(): { faces: Face[]; licences: string[] } {
	const families = record(JSON.parse(readFileSync(MANIFEST, "utf8")))["families"]
	const faces: Face[] = []
	const licences = new Set<string>()

	if (!Array.isArray(families) || families.length === 0) {
		throw new Error(`${MANIFEST} lists no families`)
	}

	for (const family of families.map(record)) {
		const licence = family["licenseFile"]
		const familyFaces = family["faces"]

		if (typeof licence !== "string" || !Array.isArray(familyFaces) || familyFaces.length === 0) {
			throw new Error(`${MANIFEST}: family ${String(family["id"])} is malformed`)
		}

		for (const face of familyFaces.map(record)) {
			const { file, bytes, sha256 } = face

			if (typeof file !== "string" || typeof bytes !== "number" || typeof sha256 !== "string") {
				throw new Error(`${MANIFEST}: a face of ${String(family["id"])} is malformed`)
			}

			faces.push({ file, bytes, sha256 })
		}

		licences.add(licence)
	}

	return { faces, licences: [...licences] }
}

export function officeFonts(): Plugin {
	const { faces, licences } = readManifest()
	// Exactly the manifest's files, so the dev middleware needs no path containment of its own.
	const served = new Set([...faces.map(face => face.file), ...licences])
	const outDir = OFFICE_FONTS_BASE.slice(1)
	let build = false

	return {
		name: "filen:office-fonts",
		configResolved(config) {
			build = config.command === "build"
		},
		resolveId(id) {
			return id === VIRTUAL_ID ? RESOLVED_ID : null
		},
		load(id) {
			if (id !== RESOLVED_ID) {
				return null
			}

			const urls = Object.fromEntries(faces.map(face => [face.file, build ? `/${hashedPath(face)}` : OFFICE_FONTS_BASE + face.file]))

			return `export const OFFICE_FONT_URLS = Object.freeze(${JSON.stringify(urls)})\n`
		},
		// The preview serves the build's copies from dist/ instead, under the preview header set.
		configureServer(server) {
			server.middlewares.use((req, res, next) => {
				const path = (req.url ?? "").split("?")[0] ?? ""
				const rel = path.startsWith(OFFICE_FONTS_BASE) ? path.slice(OFFICE_FONTS_BASE.length) : null

				if (rel === null || !served.has(rel)) {
					next()

					return
				}

				const file = join(FONTS_DIR, rel)

				res.setHeader("Content-Type", MIME[extname(rel)] ?? "application/octet-stream")
				res.setHeader("Content-Length", statSync(file).size)

				for (const [name, value] of Object.entries(COI_HEADERS)) {
					res.setHeader(name, value)
				}

				createReadStream(file)
					.on("error", () => {
						res.statusCode = 500
						res.end()
					})
					.pipe(res)
			})
		},
		generateBundle() {
			for (const face of faces) {
				const source = readFileSync(join(FONTS_DIR, face.file))
				const sha256 = createHash("sha256").update(source).digest("hex")

				if (sha256 !== face.sha256 || source.byteLength !== face.bytes) {
					this.error(
						`${face.file} is ${String(source.byteLength)} bytes with SHA-256 ${sha256}; manifest.json says ${String(face.bytes)} bytes, ${face.sha256}`
					)
				}

				this.emitFile({ type: "asset", fileName: hashedPath(face), source })
			}

			for (const licence of licences) {
				this.emitFile({ type: "asset", fileName: outDir + licence, source: readFileSync(join(FONTS_DIR, licence)) })
			}
		}
	}
}
