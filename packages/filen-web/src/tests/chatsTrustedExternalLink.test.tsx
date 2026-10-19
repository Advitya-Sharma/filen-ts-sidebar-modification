// @vitest-environment jsdom

import { describe, expect, it } from "vitest"
import { render, screen } from "@testing-library/react"
import { TrustedExternalLink } from "@/features/chats/components/thread/trustedExternalLink"
import { createTestQueryClient, queryClientWrapper } from "@/tests/testQueryClient"
import "@/lib/i18n"

// A chat message's external link asks before a domain not trusted yet opens, however it is opened: a click or
// a middle-click (which opens a new tab without a click event).

describe("TrustedExternalLink", () => {
	it("asks before a middle-click opens a domain not trusted yet, and leaves other buttons alone", async () => {
		const view = render(<TrustedExternalLink href="https://example.com/x" />, {
			wrapper: queryClientWrapper(createTestQueryClient())
		})
		const anchor = screen.getByRole("link", { name: "https://example.com/x" })
		const right = new MouseEvent("auxclick", { bubbles: true, cancelable: true, button: 2 })

		anchor.dispatchEvent(right)
		expect(right.defaultPrevented).toBe(false)
		expect(screen.queryByRole("alertdialog")).toBeNull()

		const middle = new MouseEvent("auxclick", { bubbles: true, cancelable: true, button: 1 })

		anchor.dispatchEvent(middle)
		expect(middle.defaultPrevented).toBe(true)
		expect(await screen.findByRole("alertdialog", { name: "Open external link?" })).toBeTruthy()
		view.unmount()
	})
})
