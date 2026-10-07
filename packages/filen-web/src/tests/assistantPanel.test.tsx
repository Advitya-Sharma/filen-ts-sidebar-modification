// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react"
import "@/lib/i18n"

const { kvStore, viewport } = vi.hoisted(() => ({ kvStore: new Map<string, unknown>(), viewport: { wide: true } }))

vi.mock("@/lib/storage/adapter", () => ({
	kvGetJson: (key: string) => Promise.resolve(kvStore.get(key) ?? null),
	kvSetJson: (key: string, value: unknown) => {
		kvStore.set(key, value)

		return Promise.resolve()
	}
}))

// The width query only feeds the resize handle; no query client needed.
vi.mock("@/features/shell/queries/sidebarWidth", () => ({
	useSidebarWidthQuery: () => ({ data: 360, refetch: () => Promise.resolve() })
}))

// One fake MediaQueryList for the file: breakpoints.ts caches the first one it gets.
vi.mock("@/features/shell/lib/breakpoints", () => ({
	isNarrowViewport: () => !viewport.wide,
	subscribeToLayoutBreakpoint: () => () => undefined
}))

import { registerAllActions } from "@/lib/keymap/actions"
import { AssistantSidePanel } from "@/features/assistant/components/assistantPanel"
import { AssistantToggle } from "@/features/assistant/components/assistantToggle"
import { useAssistantStore } from "@/features/assistant/store/useAssistantStore"
import { registerAssistantSession } from "@/features/assistant/store/assistantSession"
import { createMockAssistantSession, MOCK_ASSISTANT_MODEL } from "@/features/assistant/store/mockAssistantSession"
import type { AssistantModelState, AssistantSession } from "@/features/assistant/types"

let session: AssistantSession
let unregister: () => void = () => undefined

// The toggle binds app.toggleAssistant, which must be registered like main.tsx does before the first render.
registerAllActions()

function withSession(initialState: AssistantModelState): AssistantSession {
	session = createMockAssistantSession({ initialState, stepMs: 10, downloadSteps: 2 })
	unregister = registerAssistantSession(session)

	return session
}

beforeEach(() => {
	viewport.wide = true
	kvStore.clear()
	useAssistantStore.setState({ panelOpen: false, drawerOpen: false, draft: "", focusRequested: false })
})

afterEach(() => {
	unregister()
	cleanup()
	vi.useRealTimers()
})

describe("AssistantToggle", () => {
	function toggleButton(): HTMLElement {
		return screen.getByRole("button", { name: "Assistant" })
	}

	it("opens and closes the inline panel at desktop widths, reporting it as pressed", () => {
		render(<AssistantToggle itemClass={active => (active ? "active" : "idle")} />)

		expect(toggleButton().getAttribute("aria-pressed")).toBe("false")

		fireEvent.click(toggleButton())

		expect(useAssistantStore.getState().panelOpen).toBe(true)
		expect(toggleButton().getAttribute("aria-pressed")).toBe("true")
		expect(toggleButton().className).toBe("active")

		fireEvent.click(toggleButton())

		expect(useAssistantStore.getState().panelOpen).toBe(false)
	})

	it("opens the drawer below the breakpoint instead", () => {
		viewport.wide = false
		render(<AssistantToggle itemClass={() => ""} />)

		fireEvent.click(toggleButton())

		expect(useAssistantStore.getState().drawerOpen).toBe(true)
		expect(useAssistantStore.getState().panelOpen).toBe(false)
	})
})

describe("AssistantSidePanel", () => {
	it("is a labelled complementary landmark, never a dialog that would silence every shortcut", () => {
		withSession({ kind: "not-downloaded", model: MOCK_ASSISTANT_MODEL })
		render(<AssistantSidePanel />)

		expect(screen.getByRole("complementary", { name: "Assistant" })).toBeDefined()
		expect(screen.queryByRole("dialog")).toBeNull()
	})

	it("tells the session the panel shows while it is mounted", () => {
		const shown: boolean[] = []

		withSession({ kind: "not-downloaded", model: MOCK_ASSISTANT_MODEL })
		vi.spyOn(session, "setPanelOpen").mockImplementation(open => {
			shown.push(open)
		})

		const { unmount } = render(<AssistantSidePanel />)

		unmount()

		expect(shown).toEqual([true, false])
	})

	it("explains the opt-in with the model and its size, then downloads on click", () => {
		vi.useFakeTimers()
		withSession({ kind: "not-downloaded", model: MOCK_ASSISTANT_MODEL })
		render(<AssistantSidePanel />)

		expect(screen.getByText("Qwen3.5 4B")).toBeDefined()
		expect(screen.getByText(/runs entirely on this device/)).toBeDefined()

		fireEvent.click(screen.getByRole("button", { name: "Download model" }))

		expect(screen.getByRole("progressbar", { name: "Model download" })).toBeDefined()
		expect(screen.getByRole("button", { name: "Cancel download" })).toBeDefined()

		// The mock loads a downloaded model by itself because the panel is showing.
		act(() => {
			vi.runAllTimers()
		})

		expect(screen.getByRole("textbox", { name: "Ask the assistant" })).toBeDefined()
	})

	it("disables the download on a device that cannot run the model, saying why", () => {
		withSession({ kind: "unavailable", reason: "insufficient-memory", model: MOCK_ASSISTANT_MODEL })
		render(<AssistantSidePanel />)

		const button = screen.getByRole("button", { name: "Download model" })

		expect(button.hasAttribute("disabled")).toBe(true)
		expect(button.getAttribute("aria-describedby")).toBe("assistant-unavailable-reason")
		expect(screen.getByText(/doesn't have enough memory/)).toBeDefined()
	})

	it("offers no action while another tab holds the model", () => {
		withSession({ kind: "unavailable", reason: "in-use-elsewhere", model: MOCK_ASSISTANT_MODEL })
		render(<AssistantSidePanel />)

		expect(screen.getByText("Open in another Filen tab")).toBeDefined()
		expect(screen.queryByRole("button", { name: "Download model" })).toBeNull()
	})

	it("focuses the composer when the user opened the panel, and not when it was restored", () => {
		withSession({ kind: "ready", model: MOCK_ASSISTANT_MODEL })

		const restored = render(<AssistantSidePanel />)

		expect(document.activeElement).toBe(document.body)

		restored.unmount()
		useAssistantStore.setState({ focusRequested: true })
		render(<AssistantSidePanel />)

		expect(document.activeElement).toBe(screen.getByRole("textbox", { name: "Ask the assistant" }))
	})

	it("sends on Enter, keeps Shift+Enter and IME confirmation as typing, and offers Stop while replying", () => {
		vi.useFakeTimers()
		withSession({ kind: "ready", model: MOCK_ASSISTANT_MODEL })
		render(<AssistantSidePanel />)

		const composer = screen.getByRole("textbox", { name: "Ask the assistant" })

		fireEvent.change(composer, { target: { value: "find my taxes" } })
		fireEvent.keyDown(composer, { key: "Enter", shiftKey: true })
		fireEvent.keyDown(composer, { key: "Enter", isComposing: true })

		expect(useAssistantStore.getState().entries).toEqual([])

		fireEvent.keyDown(composer, { key: "Enter" })

		expect(useAssistantStore.getState().entries[0]).toMatchObject({ kind: "user", text: "find my taxes" })
		expect(useAssistantStore.getState().draft).toBe("")

		const live = document.querySelector('[aria-live="polite"]')

		expect(live?.textContent).toBe("Assistant is responding…")

		fireEvent.click(screen.getByRole("button", { name: "Stop" }))

		expect(useAssistantStore.getState().activeTurnId).toBeNull()
		expect(screen.getByRole("button", { name: "Send" })).toBeDefined()
		expect(live?.textContent).toContain("Stopped")
	})

	it("offers example questions in an empty conversation and sends one on click", () => {
		withSession({ kind: "ready", model: MOCK_ASSISTANT_MODEL })
		render(<AssistantSidePanel />)

		fireEvent.click(screen.getByRole("button", { name: "Summarize my most recent note" }))

		expect(useAssistantStore.getState().entries[0]).toMatchObject({ kind: "user", text: "Summarize my most recent note" })
	})

	it("closes from its header", () => {
		withSession({ kind: "ready", model: MOCK_ASSISTANT_MODEL })
		useAssistantStore.setState({ panelOpen: true })
		render(<AssistantSidePanel />)

		fireEvent.click(screen.getByRole("button", { name: "Close assistant" }))

		expect(useAssistantStore.getState().panelOpen).toBe(false)
	})
})
