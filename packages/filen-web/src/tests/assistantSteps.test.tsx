// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react"
import "@/lib/i18n"
import { AssistantSteps } from "@/features/assistant/components/assistantSteps"
import { useAssistantStore } from "@/features/assistant/store/useAssistantStore"
import { INITIAL_TRANSCRIPT_STATE, type TranscriptEntry, type TranscriptStep } from "@/features/assistant/store/transcript"

type StepsEntry = Extract<TranscriptEntry, { kind: "steps" }>

const NOW = 1_000_000

function tool(patch: Partial<Extract<TranscriptStep, { kind: "tool" }>> = {}): TranscriptStep {
	return {
		id: "t1:call_1_0",
		kind: "tool",
		status: "done",
		progress: null,
		tool: "list_directory",
		subject: "Documents",
		resultCount: 12,
		resultMore: false,
		failure: null,
		startedAt: NOW,
		endedAt: NOW + 300,
		...patch
	}
}

function entry(steps: TranscriptStep[], patch: Partial<StepsEntry> = {}): StepsEntry {
	return {
		kind: "steps",
		id: "steps:t1",
		turnId: "t1",
		steps,
		startedAt: NOW,
		lastAt: NOW + 4200,
		outcome: null,
		expanded: null,
		...patch
	}
}

// Reads the entry from the store, as the transcript does, so the toggle's store write shows.
function FromStore() {
	const steps = useAssistantStore(state => state.entries.find(candidate => candidate.kind === "steps"))

	return steps?.kind === "steps" ? <AssistantSteps entry={steps} /> : null
}

beforeEach(() => {
	vi.useFakeTimers({ toFake: ["setInterval", "clearInterval", "setTimeout", "clearTimeout", "Date"] })
	vi.setSystemTime(NOW)
})

afterEach(() => {
	cleanup()
	vi.useRealTimers()
	useAssistantStore.setState({ ...INITIAL_TRANSCRIPT_STATE })
})

describe("AssistantSteps", () => {
	it("lists every step while the turn runs, with what each call acts on and found, and no toggle", () => {
		render(
			<AssistantSteps
				entry={entry([
					{
						id: "t1:routing",
						kind: "routing",
						status: "done",
						progress: null,
						route: "files",
						startedAt: NOW,
						endedAt: NOW + 50
					},
					tool(),
					tool({
						id: "t1:call_2_0",
						tool: "search_items",
						subject: "invoice",
						status: "active",
						resultCount: null,
						endedAt: null
					})
				])}
			/>
		)

		const rows = screen.getAllByRole("listitem").map(row => row.textContent)

		expect(rows).toEqual(["Working with your files", 'Listing "Documents"· 12 items', 'Searching for "invoice"'])
		expect(screen.getByRole("list", { name: "Steps" })).toBeDefined()
		expect(screen.queryByRole("button")).toBeNull()
	})

	it("ticks the elapsed time of the running step only, and stops ticking once it is done", () => {
		const running = tool({ status: "active", resultCount: null, endedAt: null })
		const { rerender, container } = render(<AssistantSteps entry={entry([running])} />)

		expect(container.textContent).not.toContain(" s")

		act(() => {
			vi.advanceTimersByTime(2000)
		})

		expect(screen.getByRole("listitem").textContent).toContain("2 s")
		expect(vi.getTimerCount()).toBe(1)

		rerender(<AssistantSteps entry={entry([tool({ endedAt: NOW + 2400 })])} />)

		expect(screen.getByRole("listitem").textContent).toContain("2.4 s")
		expect(vi.getTimerCount()).toBe(0)
	})

	it("shows a long search's progress and its candidates so far, grouped, in its row while it runs", () => {
		render(
			<AssistantSteps
				entry={entry([
					tool({
						tool: "search_items",
						subject: "invoice",
						status: "active",
						progress: 0.3,
						found: 1234,
						resultCount: null,
						endedAt: null
					})
				])}
			/>
		)

		expect(screen.getByRole("listitem").textContent).toBe('Searching for "invoice"· 1,234 candidates')
		expect(document.querySelector('[style*="scaleX(0.3)"]')).not.toBeNull()
	})

	it("shows the prompt-reading bar only while part of the prompt is left", () => {
		const { container, rerender } = render(
			<AssistantSteps
				entry={entry([{ id: "t1:thinking2", kind: "thinking", status: "active", progress: 0.4, startedAt: NOW, endedAt: null }])}
			/>
		)

		expect(container.querySelector('[style*="scaleX(0.4)"]')).not.toBeNull()

		rerender(
			<AssistantSteps
				entry={entry([{ id: "t1:thinking2", kind: "thinking", status: "active", progress: 1, startedAt: NOW, endedAt: null }])}
			/>
		)

		expect(container.querySelector('[style*="scaleX"]')).toBeNull()
	})

	it("collapses a completed turn to its summary and opens it again on click", () => {
		useAssistantStore.setState({
			entries: [
				entry([tool(), tool({ id: "t1:call_2_0", tool: "read_note", subject: "Plan", resultCount: null })], { outcome: "complete" })
			]
		})
		render(<FromStore />)

		const toggle = screen.getByRole("button", { name: "Used 2 tools · 4.2 s" })

		expect(toggle.getAttribute("aria-expanded")).toBe("false")
		expect(screen.queryByRole("list")).toBeNull()

		fireEvent.click(toggle)

		expect(toggle.getAttribute("aria-expanded")).toBe("true")
		expect(toggle.getAttribute("aria-controls")).toBe(screen.getByRole("list").id)
		expect(screen.getAllByRole("listitem")).toHaveLength(2)

		fireEvent.click(toggle)

		expect(screen.queryByRole("list")).toBeNull()
	})

	it("sums up a turn without calls by its time", () => {
		render(
			<AssistantSteps
				entry={entry(
					[
						{
							id: "t1:routing",
							kind: "routing",
							status: "done",
							progress: null,
							route: "answer",
							startedAt: NOW,
							endedAt: NOW + 9
						}
					],
					{
						outcome: "complete",
						lastAt: NOW + 12_400
					}
				)}
			/>
		)

		expect(screen.getByRole("button").textContent).toBe("Worked for 12 s")
	})

	it("says why a call failed, waits or stopped, and keeps a stopped turn open", () => {
		render(
			<AssistantSteps
				entry={entry(
					[
						tool({ status: "failed", failure: "blocked", resultCount: null }),
						tool({
							id: "b",
							tool: "draft_chat_message",
							subject: "Maria",
							status: "failed",
							failure: "declined",
							resultCount: null
						}),
						tool({ id: "c", tool: "search_items", subject: "x", status: "failed", failure: "error", resultCount: null }),
						tool({ id: "d", tool: "read_note", subject: "Plan", status: "stopped", resultCount: null })
					],
					{ outcome: "cancelled" }
				)}
			/>
		)

		expect(screen.getAllByRole("listitem").map(row => row.textContent)).toEqual([
			'Listing "Documents"· Not allowed',
			'Drafting a message to "Maria"· Declined',
			'Searching for "x"· Failed',
			'Reading the note "Plan"· Stopped'
		])
		expect(screen.getByRole("button").getAttribute("aria-expanded")).toBe("true")
	})

	it("shows a call waiting for its card, and the generic line for a call without a subject", () => {
		render(
			<AssistantSteps
				entry={entry([
					tool({ tool: "draft_chat_message", subject: "Maria", status: "waiting", resultCount: null, endedAt: null }),
					tool({ id: "b", tool: "move_items", subject: null, resultCount: null }),
					tool({ id: "c", tool: "made_up", subject: "x", resultCount: null })
				])}
			/>
		)

		expect(screen.getAllByRole("listitem").map(row => row.textContent)).toEqual([
			'Drafting a message to "Maria"· Waiting for your confirmation',
			"Moving items",
			"Working"
		])
	})

	it("shows a queued call without a spinner or clock, and only the running call ticks", () => {
		const { container } = render(
			<AssistantSteps
				entry={entry([
					tool({ status: "active", resultCount: null, endedAt: null }),
					tool({ id: "b", tool: "search_items", subject: "rent", status: "queued", resultCount: null, endedAt: null })
				])}
			/>
		)

		expect(screen.getAllByRole("listitem").map(row => row.textContent)).toEqual([
			'Listing "Documents"',
			'Searching for "rent"· Up next'
		])
		expect(container.querySelectorAll(".animate-spin, .motion-safe\\:animate-spin")).toHaveLength(1)
		expect(vi.getTimerCount()).toBe(1)
	})

	it("keeps the summary neutral for a declined call or a repeat, and says so in words when a step failed", () => {
		const neutral = entry(
			[
				tool(),
				tool({ id: "b", tool: "draft_chat_message", subject: "Maria", status: "failed", failure: "declined", resultCount: null }),
				tool({ id: "c", status: "failed", failure: "repeated", resultCount: null })
			],
			{ outcome: "complete", expanded: true }
		)
		const { container, rerender } = render(<AssistantSteps entry={neutral} />)

		expect(container.querySelector(".text-destructive")).toBeNull()
		expect(screen.getByRole("button").textContent).toBe("Used 2 tools · 4.2 s")
		expect(screen.getAllByRole("listitem")[2]?.textContent).toBe('Listing "Documents"· Already done')

		rerender(
			<AssistantSteps entry={entry([tool({ status: "failed", failure: "error", resultCount: null })], { outcome: "complete" })} />
		)

		expect(screen.getByRole("button").textContent).toBe("A step failedUsed 1 tool · 4.2 s")
		expect(container.querySelector(".text-destructive")).not.toBeNull()
	})

	it("renders a subject as plain text, never as markup or a link", () => {
		const { container } = render(
			<AssistantSteps entry={entry([tool({ subject: "[login](https://evil.example) <b>x</b>", resultCount: null })])} />
		)

		expect(container.querySelector("a, b")).toBeNull()
		expect(screen.getByRole("listitem").textContent).toBe('Listing "[login](https://evil.example) <b>x</b>"')
	})
})
