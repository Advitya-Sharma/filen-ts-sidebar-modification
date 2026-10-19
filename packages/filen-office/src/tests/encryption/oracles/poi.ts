import { spawnSync } from "node:child_process"
import { existsSync, mkdtempSync, readdirSync } from "node:fs"
import { tmpdir } from "node:os"
import { delimiter, join } from "node:path"
import type { OracleResult, ResolvedEntry } from "./outputs"

// The Apache POI oracle: a tiny Java runner (OfficeOracle.java, in poi/) decrypts our output and, for Agile
// files, verifies the data-integrity HMAC over the whole EncryptedPackage stream. Opt-in through
// FILEN_OFFICE_POI (a directory of POI jars, Apache-2.0, supplied not bundled); java comes from FILEN_OFFICE_JAVA
// or the PATH. Absent jars or java → the suite skips it.

const SOURCE = join(import.meta.dirname, "poi", "OfficeOracle.java")

export interface Poi {
	readonly java: string
	readonly classpath: string
}

export function poi(env: NodeJS.ProcessEnv = process.env): Poi | undefined {
	const jars = env["FILEN_OFFICE_POI"]

	if (jars === undefined || jars === "") return undefined
	if (!existsSync(jars)) throw new Error(`FILEN_OFFICE_POI does not exist: ${jars}`)

	const classpath = readdirSync(jars)
		.filter(name => name.endsWith(".jar"))
		.map(name => join(jars, name))
		.join(delimiter)

	if (classpath === "") return undefined

	const java = env["FILEN_OFFICE_JAVA"] ?? "java"

	if (spawnSync(java, ["-version"], { encoding: "utf8", timeout: 60_000 }).status !== 0) return undefined

	return { java, classpath }
}

let built: string | undefined

// Compiles OfficeOracle.java once per process into a temporary directory, returning that directory.
function build(poi: Poi): string {
	if (built !== undefined) return built

	const out = mkdtempSync(join(tmpdir(), "filen-office-poi-"))
	const compile = spawnSync(javac(poi.java), ["-cp", poi.classpath, "-d", out, SOURCE], { encoding: "utf8", timeout: 120_000 })

	if (compile.status !== 0) throw new Error(`POI oracle did not compile: ${(compile.stderr || "").slice(0, 500)}`)

	built = out

	return out
}

function javac(java: string): string {
	return java.endsWith("java") ? `${java.slice(0, -4)}javac` : "javac"
}

export function runPoi(poi: Poi, entry: ResolvedEntry): OracleResult {
	const classes = build(poi)
	const args = [
		"-cp",
		`${poi.classpath}${delimiter}${classes}`,
		"OfficeOracle",
		entry.file,
		entry.password,
		...(entry.plaintext !== undefined ? [entry.plaintext] : [])
	]
	const result = spawnSync(poi.java, args, { encoding: "utf8", timeout: 120_000 })

	// POI logs an absent-log4j-core warning to stderr; the JSON is on stdout.
	const line = result.stdout.trim().split("\n").at(-1) ?? ""

	if (line === "") return { ok: false, error: `POI oracle exited ${result.status}: ${(result.stderr || "").slice(0, 300)}` }

	try {
		return JSON.parse(line) as OracleResult
	} catch {
		return { ok: false, error: `POI oracle output not JSON: ${result.stdout.slice(0, 300)}` }
	}
}
