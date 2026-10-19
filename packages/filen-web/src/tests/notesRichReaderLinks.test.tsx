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

import { RichReader } from "@/features/notes/components/reader/richReader"

const CONTENT = '<p><a href="https://example.com/a">web</a> <a href="mailto:a@example.com">mail</a></p>'

async function renderReader(): Promise<void> {
	const client = createTestQueryClient()

	render(<RichReader content={CONTENT} />, { wrapper: queryClientWrapper(client) })

	await waitFor(() => {
		expect(client.getQueryState(["chats", "trustedLinkDomains"])?.status).toBe("success")
	})
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

describe("RichReader links", () => {
	it("asks before a click or middle-click opens a domain not trusted yet", async () => {
		await renderReader()

		expect(screen.getByText("web").getAttribute("target")).toBe("_blank")
		expect(click("web", 2).defaultPrevented).toBe(false)
		expect(click("web", 1).defaultPrevented).toBe(true)

		const dialog = await screen.findByRole("alertdialog", { name: "Open external link?" })

		expect(dialog.textContent).toContain("example.com")
	})

	it("opens a trusted domain and a mailto: link at once", async () => {
		kvStore.set("chats.trustedLinkDomains.v1", ["example.com"])
		await renderReader()

		expect(click("web").defaultPrevented).toBe(false)
		expect(click("mail").defaultPrevented).toBe(false)
		expect(screen.queryByRole("alertdialog")).toBeNull()
	})
})
