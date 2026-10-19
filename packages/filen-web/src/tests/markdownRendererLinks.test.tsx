// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { cleanup, render, screen, waitFor } from "@testing-library/react"
import { createTestQueryClient, queryClientWrapper } from "@/tests/testQueryClient"
import "@/lib/i18n"

const { kvStore } = vi.hoisted(() => ({ kvStore: new Map<string, unknown>() }))

vi.mock("@/lib/storage/adapter", () => ({
	kvGetJson: (key: string) => Promise.resolve(kvStore.get(key) ?? null),
	kvSetJson: (key: string, value: unknown) => {
		kvStore.set(key, value)

		return Promise.resolve()
	}
}))

import { MarkdownRenderer } from "@/features/preview/components/markdownRenderer"

const TEXT = "[web](https://example.com/a) [mail](mailto:a@example.com) [tel](tel:+15551234) [bad](javascript:alert(1))"

function renderGated() {
	const client = createTestQueryClient()

	render(
		<MarkdownRenderer
			text={TEXT}
			alt="readme.md"
		/>,
		{ wrapper: queryClientWrapper(client) }
	)

	return client
}

function click(name: string, button = 0): MouseEvent {
	const event = new MouseEvent(button === 0 ? "click" : "auxclick", { bubbles: true, cancelable: true, button })

	screen.getByText(name).dispatchEvent(event)

	return event
}

beforeEach(() => {
	kvStore.clear()
})

afterEach(() => {
	cleanup()
})

describe("MarkdownRenderer links", () => {
	it("renders the shared policy's schemes as new-tab links and drops the rest", () => {
		renderGated()

		expect(screen.getByText("web").getAttribute("href")).toBe("https://example.com/a")
		expect(screen.getByText("mail").getAttribute("href")).toBe("mailto:a@example.com")
		expect(screen.getByText("tel").getAttribute("href")).toBe("tel:+15551234")
		expect(screen.getByText("bad").hasAttribute("href")).toBe(false)
		expect(screen.getByText("web").getAttribute("target")).toBe("_blank")
		expect(screen.getByText("web").getAttribute("rel")).toBe("noreferrer")
	})

	it("asks before a click or middle-click opens a domain not trusted yet", async () => {
		const client = renderGated()

		await waitFor(() => {
			expect(client.getQueryState(["chats", "trustedLinkDomains"])?.status).toBe("success")
		})

		expect(click("web", 2).defaultPrevented).toBe(false)
		expect(click("web", 1).defaultPrevented).toBe(true)

		const dialog = await screen.findByRole("alertdialog", { name: "Open external link?" })

		expect(dialog.textContent).toContain("example.com")
	})

	it("opens a trusted domain and the hostless schemes at once", async () => {
		kvStore.set("chats.trustedLinkDomains.v1", ["example.com"])

		const client = renderGated()

		await waitFor(() => {
			expect(client.getQueryState(["chats", "trustedLinkDomains"])?.status).toBe("success")
		})

		expect(click("web").defaultPrevented).toBe(false)
		expect(click("mail").defaultPrevented).toBe(false)
		expect(click("tel").defaultPrevented).toBe(false)
		expect(screen.queryByRole("alertdialog")).toBeNull()
	})

	it("leaves the links to a host that replaces the link slot", () => {
		render(
			<MarkdownRenderer
				text={TEXT}
				alt="reply"
				overrides={{ a: ({ children }) => <span>{children}</span> }}
			/>
		)

		expect(screen.getByText("web").tagName).toBe("SPAN")
		expect(click("web").defaultPrevented).toBe(false)
	})
})
