import { classifyLink } from "@filen/shared"

// The URL-safety verdict for hrefs taken from untrusted content (documents, manifests, chat text): the
// shared link policy's allowlist (http, https, mailto, tel, sms, whatsapp, geo, maps; classifyLink), plus
// a full parse so a url the browser could not navigate to is refused too. A relative or fragment-only
// reference is refused as well: resolved against the app's own origin it would open a copy of the app.
export function isSafeAbsoluteHref(href: string): boolean {
	const verdict = classifyLink(href)

	if (verdict.kind !== "external") {
		return false
	}

	try {
		new URL(verdict.url)

		return true
	} catch {
		return false
	}
}

// The stricter http(s)-only parse, for links that must open a web page (no mailto:).
export function webUrl(raw: string): URL | null {
	try {
		const url = new URL(raw)

		return url.protocol === "http:" || url.protocol === "https:" ? url : null
	} catch {
		return null
	}
}
