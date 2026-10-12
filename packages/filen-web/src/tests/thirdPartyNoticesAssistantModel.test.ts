import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { thirdPartyLicenseTexts, thirdPartyNoticesPayload } from "@/features/settings/lib/thirdPartyNotices"
import { QWEN35_4B } from "@/features/assistant/runtime/models"

const { notices } = thirdPartyNoticesPayload()

describe("third-party notices: the assistant's model", () => {
	// Filen serves these weights itself, so their license has to reach the user. Pinning another model, or
	// another revision of this one, without re-running the generator fails here.
	const notice = notices.find(entry => entry.ecosystem === "vendored" && entry.name === QWEN35_4B.label)
	const terms = notice === undefined ? [] : thirdPartyLicenseTexts(notice)
	const license = readFileSync(new URL(`../../scripts/licenses/${QWEN35_4B.id}-LICENSE`, import.meta.url), "utf8")

	it("lists the pinned model with its license and source", () => {
		expect(notice).toMatchObject({
			version: QWEN35_4B.source.revision.slice(0, 7),
			license: QWEN35_4B.license.name,
			repository: `https://huggingface.co/${QWEN35_4B.source.repo}`,
			copyright: ["Copyright 2026 Alibaba Cloud"]
		})
	})

	it("ships the license text verbatim", () => {
		expect(terms).toContain(license)
	})

	it("names the source revision, the license URL and every file served, by hash", () => {
		const origin = terms.find(text => text.includes(QWEN35_4B.source.revision)) ?? ""

		expect(origin).toContain(QWEN35_4B.license.url)

		for (const file of [...(QWEN35_4B.whole === undefined ? [] : [QWEN35_4B.whole]), ...QWEN35_4B.shards]) {
			expect(origin).toContain(file.sha256)
		}
	})
})
