import { describe, expect, it } from "vitest"
import { layerKeyFor } from "../core/sheet/index"

describe("layerKeyFor", () => {
	it("keys by the stable id, and by the document for the session without one", () => {
		expect(layerKeyFor("stable-1", "doc-1")).toEqual({ kind: "stable", id: "stable-1" })
		expect(layerKeyFor(undefined, "doc-1")).toEqual({ kind: "session", id: "doc-1" })
	})
})
