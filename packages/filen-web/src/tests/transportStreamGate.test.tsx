// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, render, screen, waitFor } from "@testing-library/react"
import { createElement } from "react"
import "@/lib/i18n"

// A `.ts` that may be a transport stream renders nothing until one probe says which it is; the verdict then
// picks the viewer through previewType. Anything else never waits, and nothing is probed twice.

const probe = vi.hoisted(() => ({ probeTransportStream: vi.fn<() => Promise<boolean | null>>() }))

vi.mock("@/features/preview/lib/transportStreamProbe", () => probe)

const { TransportStreamGate } = await import("@/features/preview/components/transportStreamGate")
const { previewType } = await import("@/features/drive/lib/preview.logic")
const { linkedFileItem } = await import("@/tests/fixtures/sdk")

function renderGate(name: string, size: bigint, uuid: `${string}-${string}-${string}-${string}-${string}`) {
	const item = linkedFileItem(name, { size, uuid })

	render(createElement(TransportStreamGate, { item, render: () => createElement("p", null, `viewer:${previewType(item)}`) }))
}

afterEach(() => {
	cleanup()
	vi.clearAllMocks()
})

describe("TransportStreamGate", () => {
	it("renders a file that cannot be a transport stream at once, probing nothing", () => {
		renderGate("index.ts", 4_000n, "a1000000-0000-0000-0000-000000000000")

		expect(screen.getByText("viewer:code")).toBeTruthy()
		expect(probe.probeTransportStream).not.toHaveBeenCalled()
	})

	it("probes a large .ts once and opens it as video when it is a transport stream", async () => {
		probe.probeTransportStream.mockResolvedValue(true)
		renderGate("clip.ts", 4_000_000n, "a2000000-0000-0000-0000-000000000000")

		expect(screen.queryByText(/viewer:/)).toBeNull()

		await waitFor(() => {
			expect(screen.getByText("viewer:video")).toBeTruthy()
		})
		expect(probe.probeTransportStream).toHaveBeenCalledOnce()
	})

	it("opens TypeScript as code, and offers Download when it is too large to show as text", async () => {
		probe.probeTransportStream.mockResolvedValue(false)
		renderGate("types.ts", 4_000_000n, "a3000000-0000-0000-0000-000000000000")

		await waitFor(() => {
			expect(screen.getByText("viewer:code")).toBeTruthy()
		})

		cleanup()
		renderGate("bundle.ts", 40_000_000n, "a4000000-0000-0000-0000-000000000000")

		await waitFor(() => {
			expect(screen.getByText("This file is too large to show as text.")).toBeTruthy()
		})
	})

	it("opens a large .mts the probe finds to be TypeScript as code", async () => {
		probe.probeTransportStream.mockResolvedValue(false)
		renderGate("module.mts", 4_000_000n, "a6000000-0000-0000-0000-000000000000")

		await waitFor(() => {
			expect(screen.getByText("viewer:code")).toBeTruthy()
		})
	})

	it("offers Retry when the probe could not be read, and probes again", async () => {
		probe.probeTransportStream.mockResolvedValueOnce(null).mockResolvedValueOnce(true)
		renderGate("clip.ts", 4_000_000n, "a5000000-0000-0000-0000-000000000000")

		const retry = await screen.findByRole("button", { name: "Try again" })

		await act(async () => {
			retry.click()
			await Promise.resolve()
		})

		await waitFor(() => {
			expect(screen.getByText("viewer:video")).toBeTruthy()
		})
		expect(probe.probeTransportStream).toHaveBeenCalledTimes(2)
	})
})
