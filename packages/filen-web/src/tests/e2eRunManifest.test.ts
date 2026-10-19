import { describe, expect, it } from "vitest"
import { fixtureManifestName } from "@/e2e-hooks/runManifest"

// Concurrent e2e runs keep their fixture manifests apart: one per run id, so a run's teardown removes only
// its own.
describe("fixtureManifestName", () => {
	it("is one file per run", () => {
		const a = fixtureManifestName("0b6f6a7e-1c1e-4d62-9d0a-1f7f4c2b9e11")
		const b = fixtureManifestName("5d0c2a9b-7e3f-4b8a-a1c4-2e9d8f7b6a50")

		expect(a).toBe("tree-0b6f6a7e-1c1e-4d62-9d0a-1f7f4c2b9e11.json")
		expect(b).not.toBe(a)
	})

	it("refuses a missing id or one that is not a plain file name", () => {
		expect(() => fixtureManifestName(undefined)).toThrow(/unset/)
		expect(() => fixtureManifestName("")).toThrow(/unset/)
		expect(() => fixtureManifestName("../tree")).toThrow(/file-name-safe/)
	})
})
