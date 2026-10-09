// @vitest-environment jsdom

import { afterEach, describe, expect, it } from "vitest"
import { act, cleanup, render } from "@testing-library/react"
import { createElement } from "react"
import { ItemIcon } from "@/features/drive/components/itemIcon"
import { recordFileKindVerdict, resetFileKindVerdicts } from "@/features/drive/lib/fileKindVerdicts"
import { linkedFileItem } from "@/tests/fixtures/sdk"

// A drive row's icon follows a content verdict that lands while it is on screen: a `.ts` read as a transport
// stream turns from the code glyph to the video one, and only the row whose file it concerns renders again.

afterEach(() => {
	cleanup()
	resetFileKindVerdicts()
})

describe("ItemIcon", () => {
	it("follows its own file's verdict", () => {
		const clip = linkedFileItem("clip.ts", { size: 40_000_000n, uuid: "b1000000-0000-0000-0000-000000000000" })
		const other = linkedFileItem("other.ts", { size: 40_000_000n, uuid: "b2000000-0000-0000-0000-000000000000" })
		const view = render(createElement("div", null, createElement(ItemIcon, { item: clip }), createElement(ItemIcon, { item: other })))
		const sources = (): string[] => [...view.container.querySelectorAll("img")].map(image => image.getAttribute("src") ?? "")
		const [codeGlyph] = sources()

		act(() => {
			recordFileKindVerdict("b1000000-0000-0000-0000-000000000000", "transportStream")
		})

		const [clipGlyph, otherGlyph] = sources()

		expect(clipGlyph).not.toBe(codeGlyph)
		expect(otherGlyph).toBe(codeGlyph)
	})
})
