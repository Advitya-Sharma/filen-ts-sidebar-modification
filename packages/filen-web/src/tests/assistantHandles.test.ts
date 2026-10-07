import { describe, expect, it } from "vitest"
import type { Dir, File, Note } from "@filen/sdk-rs"
import { narrowItem } from "@/features/drive/lib/item"
import { AssistantHandleRegistry, ROOT_HANDLE } from "@/features/assistant/tools/handles"
import { testUuid } from "@/tests/support/uuid"

const directory = narrowItem({
	uuid: testUuid("dir"),
	parent: testUuid("root"),
	color: "default",
	timestamp: 0n,
	favorited: false,
	meta: { type: "decoded", data: { name: "Docs" } }
} satisfies Dir)

const file = narrowItem({
	uuid: testUuid("file"),
	stableUUID: undefined,
	parent: testUuid("dir"),
	size: 1n,
	favorited: false,
	region: "de-1",
	bucket: "filen-1",
	timestamp: 0n,
	chunks: 1n,
	canMakeThumbnail: false,
	meta: { type: "decoded", data: { name: "a.txt", mime: "text/plain", modified: 0n, size: 1n, key: "k", version: 2 } }
} satisfies File)

const note = { uuid: testUuid("note"), title: "Plan" } as Note

describe("AssistantHandleRegistry", () => {
	it("starts with the drive root as d0", () => {
		const registry = new AssistantHandleRegistry()

		expect(registry.get(ROOT_HANDLE)).toEqual({ ref: { handle: "d0", name: "/" }, entity: { kind: "root" } })
	})

	it("hands out short per-kind handles and never a uuid", () => {
		const registry = new AssistantHandleRegistry()

		expect(registry.register({ kind: "drive", item: directory }, "Docs")).toEqual({ handle: "d1", name: "Docs" })
		expect(registry.register({ kind: "drive", item: file }, "a.txt")).toEqual({ handle: "f1", name: "a.txt" })
		expect(registry.register({ kind: "note", note }, "Plan")).toEqual({ handle: "n1", name: "Plan" })
		expect(registry.entries().every(entry => !entry.ref.handle.includes("-"))).toBe(true)
	})

	it("keeps one handle per item for the conversation, with the newest name and the known path", () => {
		const registry = new AssistantHandleRegistry()

		registry.register({ kind: "drive", item: file, path: "/Docs/a.txt" }, "a.txt")

		const again = registry.register({ kind: "drive", item: file }, "b.txt")

		expect(again).toEqual({ handle: "f1", name: "b.txt" })
		expect(registry.get("f1")?.entity).toMatchObject({ kind: "drive", path: "/Docs/a.txt" })
	})

	it("resolves a handle with stray whitespace and forgets everything on clear", () => {
		const registry = new AssistantHandleRegistry()

		registry.register({ kind: "note", note }, "Plan")

		expect(registry.get(" n1 ")?.ref.name).toBe("Plan")

		registry.clear()

		expect(registry.get("n1")).toBeUndefined()
		expect(registry.get(ROOT_HANDLE)).toBeDefined()
		expect(registry.register({ kind: "note", note }, "Plan").handle).toBe("n1")
	})
})
