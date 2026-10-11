import type { DocxLinkTarget } from "../../dom/docx/protocol"

// External link activation under the caller's policy: links are judged when clicked, never when the
// document is painted, and anything the policy refuses is ignored.

export interface DocxLinkPolicy {
	// Whether the href may be opened at all (scheme allowlist and the like).
	readonly allow: (href: string) => boolean
	readonly open: (href: string) => void
	// Called instead of open for refused links (e.g. to tell the user).
	readonly refused?: (href: string) => void
}

export function linkActivator(policy: DocxLinkPolicy): (link: Extract<DocxLinkTarget, { kind: "external" }>) => void {
	return link => {
		const href = link.href.trim()

		if (policy.allow(href)) policy.open(href)
		else policy.refused?.(href)
	}
}
