import { createHash } from "node:crypto"
import { readdirSync, readFileSync } from "node:fs"
import { EVAL_DATE_LINE } from "./prompts"

const SHARED_ASSISTANT = new URL("../../filen-shared/src/assistant/", import.meta.url)

// The eval sources that decide what a case measures: the fake app and its data, both loops, the chat
// wrappers, the prompts and the worker's context settings. Not the cases (each record hashes its own), the
// graders (a grading change needs --fresh or another --out), the report or the CLI.
const EVAL_SOURCES = ["agent", "chatWrappers", "decoys", "fakeApp", "fixture", "loopHarness", "prompts", "worker"]

// What a stored result was measured against: every source of the shared assistant (prompts, tools and
// their titles and guard metadata, the loop, guard, sanitizing, dates, budgets), the eval sources above, the
// case itself and the model (its file's sha256 and context options). A record whose fingerprint differs
// measured something else, so a resumed run does it again and the report leaves it out; otherwise a prompt
// edit would re-report the numbers from before it.
const PROMPTS_AND_TOOLS = JSON.stringify({
	date: EVAL_DATE_LINE,
	shared: readdirSync(SHARED_ASSISTANT)
		.filter(file => file.endsWith(".ts"))
		.sort()
		.map(file => [file, readFileSync(new URL(file, SHARED_ASSISTANT), "utf8")]),
	eval: EVAL_SOURCES.map(file => readFileSync(new URL(`./${file}.ts`, import.meta.url), "utf8"))
})

// A model as fingerprints name it: its id, the exact weights and the context it runs with.
export function modelKey(entry: { id: string; sha256: string; contextOptions?: unknown }): string {
	return `${entry.id}:${entry.sha256}:${JSON.stringify(entry.contextOptions ?? {})}`
}

export function caseFingerprint({ modelId, item }: { modelId: string; item: unknown }): string {
	return createHash("sha256")
		.update(PROMPTS_AND_TOOLS)
		.update("\0")
		.update(modelId)
		.update("\0")
		.update(JSON.stringify(item, (_key, value: unknown) => (typeof value === "function" ? String(value) : value)))
		.digest("hex")
		.slice(0, 16)
}

// One JSONL line, or null for a line a crash cut short.
export function parseRecordLine<T>(line: string): T | null {
	try {
		return JSON.parse(line) as T
	} catch {
		return null
	}
}
