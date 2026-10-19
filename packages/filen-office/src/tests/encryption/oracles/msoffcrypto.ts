import { spawnSync } from "node:child_process"
import { existsSync } from "node:fs"
import { join } from "node:path"
import type { OracleResult, ResolvedEntry } from "./outputs"

// The msoffcrypto-tool oracle: a Python with msoffcrypto-tool installed decrypts our output and checks the
// plaintext. Opt-in through FILEN_OFFICE_MSOFFCRYPTO (the interpreter path); absent → the suite skips it.

const SCRIPT = join(import.meta.dirname, "msoffcrypto", "decrypt.py")

export function msoffcryptoPython(env: NodeJS.ProcessEnv = process.env): string | undefined {
	const python = env["FILEN_OFFICE_MSOFFCRYPTO"]

	if (python === undefined || python === "") return undefined
	if (!existsSync(python)) throw new Error(`FILEN_OFFICE_MSOFFCRYPTO does not exist: ${python}`)

	return python
}

export function runMsoffcrypto(python: string, entry: ResolvedEntry): OracleResult {
	// -I: an isolated interpreter, so no code is loaded from the file's directory.
	const args = ["-I", SCRIPT, entry.file, entry.password, ...(entry.plaintext !== undefined ? [entry.plaintext] : [])]
	const result = spawnSync(python, args, { encoding: "utf8", timeout: 120_000 })

	if (result.status !== 0 && result.stdout.trim() === "") {
		return { ok: false, error: `msoffcrypto exited ${result.status}: ${(result.stderr || "").slice(0, 300)}` }
	}

	try {
		return JSON.parse(result.stdout.trim().split("\n").at(-1) ?? "") as OracleResult
	} catch {
		return { ok: false, error: `msoffcrypto output not JSON: ${result.stdout.slice(0, 300)}` }
	}
}
