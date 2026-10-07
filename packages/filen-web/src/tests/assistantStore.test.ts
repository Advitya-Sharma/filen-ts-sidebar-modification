import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

// Same kv boundary as sidebarWidth.test.ts: the adapter itself, backed by a Map reset per test.
const { kvStore } = vi.hoisted(() => ({ kvStore: new Map<string, unknown>() }))

vi.mock("@/lib/storage/adapter", () => ({
	kvGetJson: (key: string) => Promise.resolve(kvStore.get(key) ?? null),
	kvSetJson: (key: string, value: unknown) => {
		kvStore.set(key, value)

		return Promise.resolve()
	}
}))

import { useAssistantStore, toggleAssistant, hydrateAssistantPanelOpen } from "@/features/assistant/store/useAssistantStore"
import {
	cancelAssistantTurn,
	confirmAssistantProposal,
	downloadAssistantModel,
	registerAssistantSession,
	rejectAssistantProposal,
	removeAssistantModel,
	sendAssistantMessage,
	startNewAssistantConversation,
	trackAssistantPanelShown
} from "@/features/assistant/store/assistantSession"
import { createMockAssistantSession, MOCK_ASSISTANT_MODEL } from "@/features/assistant/store/mockAssistantSession"
import type { AssistantSession } from "@/features/assistant/types"

const STEP_MS = 10

let unregister: () => void = () => undefined

function withMock(options: Parameters<typeof createMockAssistantSession>[0] = {}): AssistantSession {
	const session = createMockAssistantSession({ stepMs: STEP_MS, downloadSteps: 2, ...options })

	unregister = registerAssistantSession(session)

	return session
}

function ready(): AssistantSession {
	return withMock({ initialState: { kind: "ready", model: MOCK_ASSISTANT_MODEL } })
}

function state(): ReturnType<typeof useAssistantStore.getState> {
	return useAssistantStore.getState()
}

beforeEach(() => {
	vi.useFakeTimers()
	kvStore.clear()
	useAssistantStore.setState({ panelOpen: false, drawerOpen: false, draft: "", focusRequested: false })
})

afterEach(() => {
	unregister()
	vi.useRealTimers()
})

describe("panel open state", () => {
	it("toggles the inline panel at desktop widths, persists it, and asks for focus only when opening", async () => {
		toggleAssistant(false)

		expect(state().panelOpen).toBe(true)
		expect(state().takeFocusRequest()).toBe(true)
		expect(state().takeFocusRequest()).toBe(false)

		await vi.waitFor(() => {
			expect(kvStore.get("assistant.v1.panelOpen")).toBe(true)
		})

		toggleAssistant(false)

		expect(state().panelOpen).toBe(false)
		expect(state().focusRequested).toBe(false)
	})

	it("toggles the drawer below the breakpoint without touching the persisted panel", () => {
		toggleAssistant(true)

		expect(state().drawerOpen).toBe(true)
		expect(state().panelOpen).toBe(false)
		expect(kvStore.has("assistant.v1.panelOpen")).toBe(false)
	})

	// The loader is load-once per module, so this is the file's only hydration case: a toggle has already
	// happened in the tests above, and that choice must win over the stored value.
	it("never lets a late stored value override the user's own toggle", async () => {
		kvStore.set("assistant.v1.panelOpen", true)
		useAssistantStore.setState({ panelOpen: false })

		await hydrateAssistantPanelOpen()

		expect(state().panelOpen).toBe(false)
	})
})

describe("session wiring", () => {
	it("tells a newly registered session whether the panel already shows, and every change after", () => {
		const calls: boolean[] = []
		const mock = createMockAssistantSession()
		const release = trackAssistantPanelShown()

		unregister = registerAssistantSession({
			...mock,
			setPanelOpen: open => {
				calls.push(open)
			}
		})
		release()

		expect(calls).toEqual([true, false])
	})

	it("downloads, then loads a downloaded model only once the panel shows", () => {
		withMock()
		downloadAssistantModel()

		expect(state().modelState.kind).toBe("downloading")

		vi.advanceTimersByTime(STEP_MS * 3)

		expect(state().modelState.kind).toBe("downloaded")

		const release = trackAssistantPanelShown()

		expect(state().modelState.kind).toBe("loading")

		vi.advanceTimersByTime(STEP_MS * 2)

		expect(state().modelState.kind).toBe("ready")

		release()
	})

	// The mock emits the turn's first activity synchronously inside send(); it must still land after the
	// user's message.
	it("streams a turn into the transcript below the message that started it, and frees the composer when it ends", () => {
		ready()

		expect(sendAssistantMessage("  find my taxes  ")).toBe(true)
		expect(state().entries[0]).toMatchObject({ kind: "user", text: "find my taxes" })
		expect(state().entries[1]).toMatchObject({ kind: "activity", tool: "search_items", status: "running" })
		expect(state().activeTurnId).not.toBeNull()
		// A second message while one is running is refused.
		expect(sendAssistantMessage("again")).toBe(false)

		vi.runAllTimers()

		expect(state().activeTurnId).toBeNull()
		expect(state().entries.some(entry => entry.kind === "reply" && entry.final)).toBe(true)
		expect(state().entries.some(entry => entry.kind === "activity" && entry.status === "done")).toBe(true)
	})

	it("refuses to send before the model is ready or with nothing typed", () => {
		withMock()

		expect(sendAssistantMessage("hello")).toBe(false)

		unregister()
		ready()

		expect(sendAssistantMessage("   ")).toBe(false)
	})

	it("stops a running turn", () => {
		ready()
		sendAssistantMessage("hello")
		cancelAssistantTurn()

		expect(state().activeTurnId).toBeNull()
		expect(state().entries.at(-1)).toMatchObject({ kind: "turnEnd", outcome: "cancelled" })
	})

	it("confirms a card synchronously into the session, then settles it from the runtime's outcome", () => {
		const mock = ready()
		const confirm = vi.spyOn(mock, "confirm")

		sendAssistantMessage("move my receipts")
		vi.runAllTimers()

		const card = state().entries.find(entry => entry.kind === "proposal")

		if (card?.kind !== "proposal") {
			throw new Error("no card")
		}

		confirmAssistantProposal(card.proposal.id)

		expect(confirm).toHaveBeenCalledWith(card.proposal.id)
		expect(state().entries.find(entry => entry.id === card.id)).toMatchObject({ status: "running" })

		vi.runAllTimers()

		expect(state().entries.find(entry => entry.id === card.id)).toMatchObject({ status: "done" })
	})

	it("rejects a declined card, and every unanswered card when a new message goes out", () => {
		const mock = ready()
		const reject = vi.spyOn(mock, "reject")

		sendAssistantMessage("trash these")
		vi.runAllTimers()
		sendAssistantMessage("copy those")
		vi.runAllTimers()

		expect(reject).toHaveBeenCalledWith("turn-1-proposal")
		expect(state().entries.find(entry => entry.id === "proposal:turn-1-proposal")).toMatchObject({ status: "expired" })

		rejectAssistantProposal("turn-2-proposal")

		expect(reject).toHaveBeenCalledWith("turn-2-proposal")
		expect(state().entries.find(entry => entry.id === "proposal:turn-2-proposal")).toMatchObject({ status: "cancelled" })
	})

	it("starts a new conversation with an empty transcript and a reset session", () => {
		const mock = ready()
		const reset = vi.spyOn(mock, "reset")

		sendAssistantMessage("hello")
		vi.runAllTimers()
		startNewAssistantConversation()

		expect(reset).toHaveBeenCalled()
		expect(state().entries).toEqual([])
		expect(state().modelState.kind).toBe("ready")
	})

	it("removes the downloaded model back to the setup state", async () => {
		ready()

		await removeAssistantModel()

		expect(state().modelState.kind).toBe("not-downloaded")
	})
})
