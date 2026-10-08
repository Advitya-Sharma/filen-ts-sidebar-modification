// @vitest-environment jsdom

import { afterEach, describe, expect, it } from "vitest"
import { cleanup, render } from "@testing-library/react"
import { AssistantReplyMarkdown } from "@/features/assistant/components/replyMarkdown"

afterEach(() => {
	cleanup()
})

describe("AssistantReplyMarkdown", () => {
	it("never renders a live link or an image, however the URL is spelled", () => {
		const replies = [
			"[Verify](https:evil.example/login)",
			"[Verify](h&#116;tps://evil.example/login)",
			"[Verify](https&#58;//evil.example/login)",
			"[Verify](https:/\\evil.example/login)",
			"[x](http:evil.com)",
			"[Mail](mailto:a@evil.example?body=secret)",
			"<http:evil.com>",
			"write to support@filen-verify.example",
			"[a][1]\n\n[1]: https://evil.example",
			"![logo](https://cdn.filen.io/a.png)"
		]

		for (const text of replies) {
			const { container, unmount } = render(
				<AssistantReplyMarkdown
					text={text}
					alt="Reply"
				/>
			)

			expect(container.querySelector("a, img, [href], [src]")).toBeNull()
			unmount()
		}
	})

	it("keeps the text of a link and an image's description", () => {
		const { container } = render(
			<AssistantReplyMarkdown
				text={"See [the docs](https://filen.io) and ![a chart](https://cdn.filen.io/a.png), **bold**"}
				alt="Reply"
			/>
		)

		expect(container.textContent).toContain("See the docs and a chart, bold")
		expect(container.querySelector("strong")).not.toBeNull()
	})
})
