// @vitest-environment happy-dom

import { vi, describe, it, expect, beforeEach, afterEach } from "vitest"
import { createElement } from "react"
import { render, waitFor, act } from "@testing-library/react"

const { pinVerdict } = vi.hoisted(() => ({
	pinVerdict: vi.fn()
}))

// The gate's sniff is injected; the real reader is never reached.
vi.mock("expo/fetch", () => ({
	fetch: vi.fn()
}))

vi.mock("@shopify/flash-list", async () => {
	const { useState } = await import("react")

	return {
		useRecyclingState: (init: unknown) => useState(typeof init === "function" ? (init as () => unknown)() : init)
	}
})

vi.mock("@/stores/useDrivePreview.store", () => ({
	default: {
		getState: () => ({ pinVerdict })
	}
}))

vi.mock("@/components/drivePreview/previewStatus", async () => {
	const { createElement: h } = await import("react")

	return {
		PreviewSpinner: () => h("div", { "data-testid": "spinner" })
	}
})

vi.mock("@/components/drivePreview/unavailableOfflineNotice", async () => {
	const { createElement: h } = await import("react")

	return {
		default: () => h("div", { "data-testid": "offline" })
	}
})

vi.mock("@/components/drivePreview/previewLoadFailedNotice", async () => {
	const { createElement: h } = await import("react")

	return {
		default: (props: { onRetry: () => void }) => h("button", { "data-testid": "retry", onClick: props.onRetry })
	}
})

import TransportStreamGate from "@/components/drivePreview/transportStreamGate"
import fileKindVerdicts from "@/lib/fileKindVerdicts"
import type { DriveItemFileExtracted } from "@/types"

const SIZE = 4 * 1024 * 1024

const TS_188 = { stride: 188, prefix: 0 }
const BDAV = { stride: 192, prefix: 4 }

type Sniff = (readRange: unknown, size: number) => Promise<{ stride: number; prefix: number } | false | null>

function file(uuid: string, name = "capture.ts"): DriveItemFileExtracted {
	return {
		type: "file",
		data: { uuid, size: BigInt(SIZE), decryptedMeta: { name, mime: "video/mp2t" }, undecryptable: false }
	} as unknown as DriveItemFileExtracted
}

function urlQuery(status: "pending" | "success" | "error", data?: string | null) {
	return { status, data, refetch: vi.fn() } as never
}

function renderGate(uuid: string, sniff: Sniff, query = urlQuery("success", "http://127.0.0.1:1/f"), name?: string) {
	return render(
		createElement(TransportStreamGate, {
			item: file(uuid, name),
			fileUrlQuery: query,
			sniff
		})
	)
}

describe("TransportStreamGate", () => {
	beforeEach(() => {
		pinVerdict.mockReset()
	})

	it("shows a loading state, then records and pins what the read found", async () => {
		const sniff = vi.fn<Sniff>(async () => TS_188)
		const { container } = renderGate("stream", sniff)

		expect(container.querySelector("[data-testid='spinner']")).not.toBeNull()

		await waitFor(() => expect(pinVerdict).toHaveBeenCalledWith("stream", "transportStream"))

		expect(sniff).toHaveBeenCalledWith(expect.any(Function), SIZE)
		expect(fileKindVerdicts.get("stream")).toBe("transportStream")
	})

	it("records TypeScript as notTransportStream", async () => {
		renderGate("source", async () => false)

		await waitFor(() => expect(pinVerdict).toHaveBeenCalledWith("source", "notTransportStream"))

		expect(fileKindVerdicts.get("source")).toBe("notTransportStream")
	})

	it("keeps any stream on iOS, which previews none, a BDAV one included", async () => {
		renderGate("bdav-ios", async () => BDAV)

		await waitFor(() => expect(pinVerdict).toHaveBeenCalledWith("bdav-ios", "transportStream"))

		expect(fileKindVerdicts.get("bdav-ios")).toBe("transportStream")
	})

	it("pins a verdict already known without reading", () => {
		fileKindVerdicts.record("known", "notTransportStream")

		const sniff = vi.fn<Sniff>(async () => TS_188)

		renderGate("known", sniff)

		expect(pinVerdict).toHaveBeenCalledWith("known", "notTransportStream")
		expect(sniff).not.toHaveBeenCalled()
	})

	it("offers a retry when the read fails, and reads again on it", async () => {
		const sniff = vi.fn<Sniff>(async () => null)
		const { container } = renderGate("flaky", sniff)

		await waitFor(() => expect(container.querySelector("[data-testid='retry']")).not.toBeNull())

		sniff.mockResolvedValueOnce(TS_188)

		await act(async () => {
			;(container.querySelector("[data-testid='retry']") as HTMLButtonElement).click()
		})

		await waitFor(() => expect(pinVerdict).toHaveBeenCalledWith("flaky", "transportStream"))

		expect(sniff).toHaveBeenCalledTimes(2)
	})

	it("waits for the URL, and shows the offline notice when there is none to read", () => {
		const sniff = vi.fn<Sniff>(async () => TS_188)

		expect(renderGate("pending", sniff, urlQuery("pending")).container.querySelector("[data-testid='spinner']")).not.toBeNull()
		expect(renderGate("offline", sniff, urlQuery("success", null)).container.querySelector("[data-testid='offline']")).not.toBeNull()
		expect(renderGate("failed", sniff, urlQuery("error")).container.querySelector("[data-testid='retry']")).not.toBeNull()
		expect(sniff).not.toHaveBeenCalled()
	})
})

describe("TransportStreamGate on Android", () => {
	beforeEach(() => {
		pinVerdict.mockReset()
	})

	async function androidGate() {
		vi.resetModules()

		const { Platform } = await import("react-native")

		Platform.OS = "android"

		const gate = (await import("@/components/drivePreview/transportStreamGate")).default
		const verdicts = (await import("@/lib/fileKindVerdicts")).default

		return {
			verdicts,
			renderAndroid: (uuid: string, sniff: Sniff, name?: string) =>
				render(
					createElement(gate, {
						item: file(uuid, name),
						fileUrlQuery: urlQuery("success", "http://127.0.0.1:1/f"),
						sniff
					})
				)
		}
	}

	afterEach(async () => {
		const { Platform } = await import("react-native")

		Platform.OS = "ios"
		vi.resetModules()
	})

	it("keeps a 188-byte stream, which Media3 plays", async () => {
		const { verdicts, renderAndroid } = await androidGate()

		renderAndroid("a188", async () => TS_188)

		await waitFor(() => expect(pinVerdict).toHaveBeenCalledWith("a188", "transportStream"))

		expect(verdicts.get("a188")).toBe("transportStream")
	})

	it("keeps no verdict for a BDAV stream Media3 cannot read, and pins its page unplayable", async () => {
		const { verdicts, renderAndroid } = await androidGate()

		renderAndroid("bdav", async () => BDAV, "clip.m2ts")

		await waitFor(() => expect(pinVerdict).toHaveBeenCalledWith("bdav", "unplayable"))

		expect(verdicts.get("bdav")).toBeUndefined()
	})

	it("keeps a playable verdict for an .m2ts, whose name alone opened no player", async () => {
		const { verdicts, renderAndroid } = await androidGate()

		renderAndroid("named", async () => TS_188, "clip.m2ts")

		await waitFor(() => expect(pinVerdict).toHaveBeenCalledWith("named", "transportStream"))

		expect(verdicts.get("named")).toBe("transportStream")
	})
})
