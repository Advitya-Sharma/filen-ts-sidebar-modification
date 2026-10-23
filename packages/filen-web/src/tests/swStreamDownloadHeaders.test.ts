import { describe, expect, it } from "vitest"
import { streamDownloadHeaders } from "@/lib/sw/streamDownloadHeaders"

// The headers of a download the page streams through the service worker.

describe("streamDownloadHeaders", () => {
	it("states the length of bytes whose size is known", () => {
		expect(streamDownloadHeaders("entry.txt", 2048)).toMatchObject({
			"Content-Type": "application/octet-stream",
			"Content-Length": "2048",
			"X-Content-Type-Options": "nosniff"
		})
	})

	it("states no length for bytes made as they stream (an archive the page writes as it goes)", () => {
		const headers = streamDownloadHeaders("Filen.zip", null)

		expect(headers).not.toHaveProperty("Content-Length")
		expect(headers["Content-Disposition"]).toContain("Filen.zip")
	})
})
