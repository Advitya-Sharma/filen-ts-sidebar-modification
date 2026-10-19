import { spawn, spawnSync } from "node:child_process"
import { copyFileSync, existsSync, mkdirSync, mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import { basename, dirname, join } from "node:path"
import type { OracleResult, ResolvedEntry } from "./outputs"

// The LibreOffice UNO oracle: a headless soffice opens our output with its password (open.py, in lo/), proving
// LibreOffice accepts both. Opt-in through FILEN_OFFICE_SOFFICE (the same env the interop suite uses; the first
// install is taken). Opens copies only — the originals are never handed to LibreOffice. Needs the Python bundled
// with that install.

const SCRIPT = join(import.meta.dirname, "lo", "open.py")

export interface LibreOffice {
	readonly soffice: string
	readonly python: string
}

function officeEnv(): NodeJS.ProcessEnv {
	return { ...process.env, SAL_ENABLE_FILE_LOCKING: "0", PYTHONDONTWRITEBYTECODE: "1" }
}

// The first FILEN_OFFICE_SOFFICE install plus its bundled Python (macOS: .../Contents/Resources/python).
export function libreOffice(env: NodeJS.ProcessEnv = process.env): LibreOffice | undefined {
	const soffice = (env["FILEN_OFFICE_SOFFICE"] ?? "").split(",").find(path => path !== "" && existsSync(path))

	if (soffice === undefined) return undefined

	const python = join(dirname(soffice), "..", "Resources", "python")

	if (!existsSync(python)) return undefined

	return { soffice, python }
}

// Opens every entry in one soffice session (startup is the slow part), returning a result per entry name.
export function runLibreOffice(lo: LibreOffice, entries: readonly ResolvedEntry[]): Map<string, OracleResult> {
	const work = mkdtempSync(join(tmpdir(), "filen-office-lo-"))
	const profile = join(work, "profile")
	const copies = join(work, "copies")

	mkdirSync(profile, { recursive: true })
	mkdirSync(copies, { recursive: true })

	const pipe = `filen-office-oracle-${process.pid}-${Date.now()}`
	const soffice = spawn(
		lo.soffice,
		[
			`-env:UserInstallation=file://${profile}`,
			"--headless",
			"--norestore",
			"--nologo",
			"--nodefault",
			"--nolockcheck",
			`--accept=pipe,name=${pipe};urp;`
		],
		{ detached: true, stdio: "ignore", env: officeEnv() }
	)
	const results = new Map<string, OracleResult>()

	try {
		entries.forEach((entry, index) => {
			const copy = join(copies, `${index}-${basename(entry.file)}`)

			copyFileSync(entry.file, copy)

			const out = join(copies, `${index}.txt`)
			// open.py retries the connection for 60 s, so it also covers soffice startup.
			const run = spawnSync(lo.python, [SCRIPT, pipe, copy, out, entry.password], { encoding: "utf8", timeout: 180_000 })
			const line = run.stdout.trim().split("\n").at(-1) ?? ""
			let parsed: { ok: boolean; textLength?: number; error?: string }

			try {
				parsed = JSON.parse(line) as typeof parsed
			} catch {
				results.set(entry.name, { ok: false, error: `LibreOffice oracle output not JSON: ${run.stdout.slice(0, 300)}` })

				return
			}

			results.set(
				entry.name,
				parsed.ok
					? { ok: true, ...(parsed.textLength !== undefined && { plaintextSize: parsed.textLength }) }
					: { ok: false, error: parsed.error ?? "not loaded" }
			)
		})
	} finally {
		try {
			if (soffice.pid !== undefined) process.kill(-soffice.pid, "SIGKILL")
		} catch {
			// Already gone.
		}
	}

	return results
}
