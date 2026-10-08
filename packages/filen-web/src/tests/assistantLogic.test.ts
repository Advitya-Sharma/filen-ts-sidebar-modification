import { describe, expect, it } from "vitest"
import { ASSISTANT_TOOLS } from "@filen/shared/assistant"
import { EN_CATALOGS } from "@/lib/i18n/catalog"
import {
	canRemoveModel,
	downloadByteFigures,
	downloadPercent,
	proposalTargetKind,
	spokenText,
	toolActivityKey,
	transcriptRows
} from "@/features/assistant/components/assistant.logic"
import type { TranscriptEntry } from "@/features/assistant/store/transcript"

const MODEL = { id: "m", name: "M", sizeBytes: 100 }

describe("downloadPercent", () => {
	it("floors to a whole percentage and clamps to 0..100", () => {
		expect(downloadPercent(0, 100)).toBe(0)
		expect(downloadPercent(999, 1000)).toBe(99)
		expect(downloadPercent(1000, 1000)).toBe(100)
		expect(downloadPercent(2000, 1000)).toBe(100)
	})

	it("reads 0 for an unknown size instead of dividing by zero", () => {
		expect(downloadPercent(10, 0)).toBe(0)
	})
})

describe("downloadByteFigures", () => {
	const total = 2_740_938_752

	it("shows both figures at one precision, so the received part never reads above the total", () => {
		// One decimal against two used to read "2.6 GiB of 2.55 GiB" here.
		expect(downloadByteFigures(total - 1_000_000, total)).toEqual({ received: "2.6 GiB", total: "2.6 GiB" })
		expect(downloadByteFigures(total, total)).toEqual({ received: "2.6 GiB", total: "2.6 GiB" })
		expect(downloadByteFigures(512 * 1024 * 1024, total)).toEqual({ received: "512.0 MiB", total: "2.6 GiB" })
	})

	it("caps the received part at the total and at zero", () => {
		expect(downloadByteFigures(total + 100_000_000, total).received).toBe("2.6 GiB")
		expect(downloadByteFigures(-5, total).received).toBe("0 B")
	})
})

describe("canRemoveModel", () => {
	it("offers removal only while the files are on disk", () => {
		expect(canRemoveModel({ kind: "downloaded", model: MODEL })).toBe(true)
		expect(canRemoveModel({ kind: "ready", model: MODEL })).toBe(true)
		expect(canRemoveModel({ kind: "error", error: "load-failed", model: MODEL })).toBe(true)
		expect(canRemoveModel({ kind: "error", error: "out-of-memory", model: MODEL })).toBe(true)
		expect(canRemoveModel({ kind: "error", error: "download-failed", model: MODEL })).toBe(false)
		expect(canRemoveModel({ kind: "not-downloaded", model: MODEL })).toBe(false)
		expect(canRemoveModel({ kind: "downloading", model: MODEL, receivedBytes: 1, waiting: false })).toBe(false)
		expect(canRemoveModel({ kind: "loading", model: MODEL, verifying: false })).toBe(false)
		expect(canRemoveModel({ kind: "unavailable", reason: "in-use-elsewhere", model: MODEL })).toBe(false)
	})
})

describe("transcriptRows", () => {
	const user: TranscriptEntry = { kind: "user", id: "user:t1", turnId: "t1", text: "hi" }

	it("adds the waiting bubble while a turn shows nothing in progress", () => {
		expect(transcriptRows([user], true).at(-1)).toEqual({ kind: "thinking", id: "thinking" })
		expect(transcriptRows([user], false)).toEqual([user])
	})

	it("leaves it out while a reply streams or an activity runs", () => {
		const streaming: TranscriptEntry = { kind: "reply", id: "reply:r", turnId: "t1", text: "x", final: false }
		const running: TranscriptEntry = { kind: "activity", id: "activity:a", turnId: "t1", tool: "read_note", status: "running" }
		const done: TranscriptEntry = { ...running, status: "done" }

		expect(transcriptRows([user, streaming], true)).toHaveLength(2)
		expect(transcriptRows([user, running], true)).toHaveLength(2)
		expect(transcriptRows([user, done], true)).toHaveLength(3)
	})
})

describe("toolActivityKey", () => {
	// The runtime sends tool names only; a tool added to the shared registry without wording here would
	// fall back to the generic line.
	it("words every tool in the shared registry with a key the catalog ships", () => {
		for (const tool of ASSISTANT_TOOLS) {
			const key = toolActivityKey(tool.name)

			expect(key, tool.name).not.toBe("assistantToolGeneric")
			expect(EN_CATALOGS.assistant[key], tool.name).toBeTypeOf("string")
		}
	})

	it("falls back to a generic line for a name it does not know, including prototype keys", () => {
		expect(toolActivityKey("made_up")).toBe("assistantToolGeneric")
		expect(toolActivityKey("constructor")).toBe("assistantToolGeneric")
	})
})

describe("proposalTargetKind", () => {
	it("reads the kind from the handle's leading letter, defaulting to a file", () => {
		expect(proposalTargetKind("d3")).toBe("directory")
		expect(proposalTargetKind("n1")).toBe("note")
		expect(proposalTargetKind("c2")).toBe("chat")
		expect(proposalTargetKind("x9")).toBe("file")
		expect(proposalTargetKind("")).toBe("file")
	})
})

describe("spokenText", () => {
	it("drops emphasis, code and heading/quote markers but keeps the words", () => {
		expect(spokenText("# Title\n> quoted **bold** and `code`\nmy_file.txt")).toBe("Title\nquoted bold and code\nmy_file.txt")
	})
})
