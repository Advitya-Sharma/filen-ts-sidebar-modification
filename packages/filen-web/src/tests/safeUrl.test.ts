import { describe, expect, it } from "vitest"
import { isSafeAbsoluteHref, webUrl } from "@/lib/safeUrl"

describe("isSafeAbsoluteHref", () => {
	it("allows https", () => {
		expect(isSafeAbsoluteHref("https://example.com")).toBe(true)
	})

	it("allows http", () => {
		expect(isSafeAbsoluteHref("http://example.com")).toBe(true)
	})

	it("allows mailto", () => {
		expect(isSafeAbsoluteHref("mailto:a@example.com")).toBe(true)
	})

	it("rejects javascript:", () => {
		expect(isSafeAbsoluteHref("javascript:alert(1)")).toBe(false)
	})

	it("rejects a javascript: scheme disguised with case and embedded whitespace", () => {
		expect(isSafeAbsoluteHref("   Java\tScript:alert(1)")).toBe(false)
	})

	it("rejects data:", () => {
		expect(isSafeAbsoluteHref("data:text/html,<script>alert(1)</script>")).toBe(false)
	})

	it("rejects vbscript:", () => {
		expect(isSafeAbsoluteHref("vbscript:msgbox(1)")).toBe(false)
	})

	it("rejects file:", () => {
		expect(isSafeAbsoluteHref("file:///etc/passwd")).toBe(false)
	})

	it("rejects an empty href", () => {
		expect(isSafeAbsoluteHref("")).toBe(false)
	})

	it("rejects relative and fragment-only hrefs, which would resolve against the app's origin", () => {
		expect(isSafeAbsoluteHref("#section")).toBe(false)
		expect(isSafeAbsoluteHref("docs/setup.md")).toBe(false)
		expect(isSafeAbsoluteHref("//example.com/x")).toBe(false)
	})

	it("rejects a scheme-relative href, which parses alone but opens against the app's origin", () => {
		// new URL("https:/reset/x") is https://reset/x, yet the page resolves it to its own origin's /reset/x.
		expect(isSafeAbsoluteHref("https:/reset/x")).toBe(false)
		expect(isSafeAbsoluteHref("https:d/uuid#key")).toBe(false)
		expect(isSafeAbsoluteHref("http:example.com")).toBe(false)
		expect(isSafeAbsoluteHref("https://example.com/a?b#c")).toBe(true)
	})

	it("fails closed on a string the URL parser can't resolve", () => {
		expect(isSafeAbsoluteHref("http://")).toBe(false)
		expect(isSafeAbsoluteHref("https://exa%mple.com")).toBe(false)
	})

	it("allows the shared policy's communication schemes", () => {
		for (const href of ["tel:+15551234", "sms:+15551234", "whatsapp://send?text=hi", "geo:37.78,-122.41", "maps:?q=Berlin"]) {
			expect(isSafeAbsoluteHref(href)).toBe(true)
		}
	})

	it("rejects interior whitespace and control characters the URL parser would strip", () => {
		expect(isSafeAbsoluteHref("https://exa\nmple.com")).toBe(false)
		expect(isSafeAbsoluteHref("\x01https://example.com")).toBe(false)
	})
})

describe("webUrl", () => {
	it("parses http(s) URLs", () => {
		expect(webUrl("https://example.com")?.href).toBe("https://example.com/")
		expect(webUrl("http://example.com/a")?.href).toBe("http://example.com/a")
	})

	it("rejects every other scheme, mailto: included", () => {
		expect(webUrl("mailto:a@b.com")).toBeNull()
		expect(webUrl("javascript:alert(1)")).toBeNull()
		expect(webUrl("ftp://host/file")).toBeNull()
	})

	it("rejects relative and unparseable input", () => {
		expect(webUrl("owner/repo")).toBeNull()
		expect(webUrl("not a url")).toBeNull()
	})
})
