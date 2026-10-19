// The single URL-safety verdict for hrefs taken from untrusted content (documents, manifests, chat
// text). The URL constructor rather than a regex, because it normalizes what it parses (lowercases the
// scheme, strips leading/embedded whitespace inside it) before `.protocol` is read, so an obfuscated
// scheme like " javascript:..." or "java\tscript:..." can't slip past a naive string match. Parsed
// with no base: a relative or fragment-only reference would resolve against the app's own origin and
// open a copy of the app, so it is rejected along with every unlisted scheme.

const SAFE_LINK_PROTOCOLS = new Set(["http:", "https:", "mailto:"])
// A special-scheme href without "//" ("https:/reset/x", "http:host") parses as absolute alone, but a page of
// the same scheme resolves it against its own origin: it must read the same against either base.
const PROBE_BASES = ["https://probe.invalid/a/", "http://probe.invalid/a/"]

export function isSafeAbsoluteHref(href: string): boolean {
	try {
		const url = new URL(href)

		return SAFE_LINK_PROTOCOLS.has(url.protocol) && PROBE_BASES.every(base => new URL(href, base).href === url.href)
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
