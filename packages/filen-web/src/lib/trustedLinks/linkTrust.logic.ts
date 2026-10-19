import { classifyLink, isWebLinkScheme, parseFilenPublicLink } from "@filen/shared"

// The domain a link's one-time trust confirmation is asked for, or null when none is warranted: a web
// (http/https) link only, since every other allowed scheme hands off to the OS's own compose, dial or map
// UI, which is the confirmation there; never a Filen public link, which is the app's own domain; and never
// a link the shared policy blocks or one too malformed to have a stable domain identity. Callers that
// render links from untrusted content judge the scheme first (isSafeAbsoluteHref, hardenLinkHref).
export function externalLinkDomain(url: string): string | null {
	const verdict = classifyLink(url)

	if (verdict.kind !== "external" || !isWebLinkScheme(verdict.scheme) || parseFilenPublicLink(verdict.url) !== null) {
		return null
	}

	try {
		const host = new URL(verdict.url).hostname.toLowerCase()

		return host === "" ? null : host
	} catch {
		return null
	}
}

// True means: prevent the anchor's own navigation and ask first. A null domain never intercepts, same as
// an already-trusted one.
export function shouldInterceptLinkClick(domain: string | null, trustedDomains: ReadonlySet<string>): boolean {
	return domain !== null && !trustedDomains.has(domain)
}
