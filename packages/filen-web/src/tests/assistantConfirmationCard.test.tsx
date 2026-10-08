// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { cleanup, fireEvent, render, screen, within } from "@testing-library/react"
import "@/lib/i18n"

// Only the browser can dispatch a trusted click, so the trust check is the one thing stood in for here.
const { trusted } = vi.hoisted(() => ({ trusted: { value: false } }))

vi.mock("@/features/assistant/components/userActivation", () => ({ isTrustedActivation: () => trusted.value }))

import { AssistantConfirmationCard } from "@/features/assistant/components/assistantConfirmationCard"
import { registerAssistantSession } from "@/features/assistant/store/assistantSession"
import { createMockAssistantSession, MOCK_ASSISTANT_MODEL } from "@/features/assistant/store/mockAssistantSession"
import type { ProposalStatus } from "@/features/assistant/store/transcript"
import type { AssistantEvent, AssistantProposal, AssistantSession } from "@/features/assistant/types"

const MOVE: AssistantProposal = {
	id: "p1",
	action: "move",
	targets: [
		{ handle: "f1", name: "Taxes 2025.pdf" },
		{ handle: "d4", name: "Scans" }
	],
	destination: "Documents/Archive",
	text: {}
}

let session: AssistantSession
let unregister: () => void
let deliver: (event: AssistantEvent) => void = () => undefined

// As the live turn would: only a card the session itself proposed can be confirmed.
function propose(proposal: AssistantProposal): void {
	deliver({ type: "proposal", turnId: "turn-live", proposal })
}

beforeEach(() => {
	trusted.value = false

	const mock = createMockAssistantSession({ initialState: { kind: "ready", model: MOCK_ASSISTANT_MODEL } })

	session = {
		...mock,
		subscribe: listener => {
			deliver = listener

			return mock.subscribe(listener)
		}
	}
	unregister = registerAssistantSession(session)
})

afterEach(() => {
	unregister()
	cleanup()
})

function renderCard(proposal: AssistantProposal, status: ProposalStatus = "pending"): void {
	render(
		<AssistantConfirmationCard
			proposal={proposal}
			status={status}
		/>
	)
}

describe("AssistantConfirmationCard", () => {
	it("spells out the count, every name and the destination, never a handle", () => {
		renderCard(MOVE)

		const card = screen.getByRole("group", { name: "Move 2 items" })
		const items = within(card).getByRole("list", { name: "Items" })

		expect(
			within(items)
				.getAllByRole("listitem")
				.map(item => item.textContent)
		).toEqual(["Taxes 2025.pdf", "Scans"])
		expect(card.textContent).toContain("Documents/Archive")
		expect(card.textContent).not.toMatch(/\bf1\b|\bd4\b/)
	})

	it("ignores a click the browser did not dispatch itself", () => {
		const confirm = vi.spyOn(session, "confirm")

		renderCard(MOVE)
		fireEvent.click(screen.getByRole("button", { name: "Confirm" }))

		expect(confirm).not.toHaveBeenCalled()
	})

	// Asserted straight after the dispatch returns, with no await in between: a confirm that only ran after
	// a microtask would have lost the user activation a download needs.
	it("runs a trusted confirm synchronously, inside the click", () => {
		const confirm = vi.spyOn(session, "confirm")

		trusted.value = true
		propose(MOVE)
		renderCard(MOVE)
		fireEvent.click(screen.getByRole("button", { name: "Confirm" }))

		expect(confirm).toHaveBeenCalledWith("p1")
	})

	// A card restored from a saved conversation, or left from another one, never reaches the session.
	it("never confirms a card this page's session did not propose", () => {
		const confirm = vi.spyOn(session, "confirm")

		trusted.value = true
		renderCard(MOVE)
		fireEvent.click(screen.getByRole("button", { name: "Confirm" }))

		expect(confirm).not.toHaveBeenCalled()
	})

	it("confirms a proposed card once only", () => {
		const confirm = vi.spyOn(session, "confirm")

		trusted.value = true
		propose(MOVE)
		renderCard(MOVE)
		fireEvent.click(screen.getByRole("button", { name: "Confirm" }))
		fireEvent.click(screen.getByRole("button", { name: "Confirm" }))

		expect(confirm).toHaveBeenCalledTimes(1)
	})

	it("declines through the session", () => {
		const reject = vi.spyOn(session, "reject")

		renderCard(MOVE)
		fireEvent.click(screen.getByRole("button", { name: "Cancel" }))

		expect(reject).toHaveBeenCalledWith("p1")
	})

	it("labels a trash confirm as what it does", () => {
		renderCard({ ...MOVE, action: "trash", destination: null })

		expect(screen.getByRole("group", { name: "Move 2 items to the trash" })).toBeDefined()
		expect(screen.getByRole("button", { name: "Move to trash" })).toBeDefined()
	})

	it("shows a drafted message verbatim with the chat it goes into", () => {
		const text = "Line one\n  indented <b>not html</b>"

		renderCard({ id: "p2", action: "draft-message", targets: [{ handle: "c1", name: "Family" }], destination: null, text: { text } })

		const card = screen.getByRole("group", { name: "Put this draft into Family?" })

		expect(within(card).getByLabelText("Draft").textContent).toBe(text)
		expect(card.textContent).toContain("You still send it yourself.")
		expect(within(card).getByRole("button", { name: "Add draft" })).toBeDefined()
	})

	it("replaces the buttons with the outcome once answered", () => {
		renderCard(MOVE, "expired")

		expect(screen.queryByRole("button")).toBeNull()
		expect(screen.getByText("Expired. Ask again to do this.")).toBeDefined()
	})

	it("shows a saved card that was still running as interrupted, with no buttons", () => {
		renderCard(MOVE, "interrupted")

		expect(screen.queryByRole("button")).toBeNull()
		expect(screen.getByText("This may have finished after you left. Check before asking again.")).toBeDefined()
	})
})
