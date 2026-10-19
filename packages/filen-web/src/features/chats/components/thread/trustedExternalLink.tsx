import { externalLinkDomain } from "@/features/chats/lib/linkTrust.logic"
import { useTrustedLinkGate } from "@/features/chats/hooks/useTrustedLinkGate"
import { ExternalAnchor } from "@/features/chats/components/thread/externalAnchor"

// A message-text link anchor, gated by a one-time-per-domain trust confirmation before it's ever opened
// as a bare external navigation. Scoped to EXTERNAL (non-Filen) links only: a Filen public link
// resolves to null here (externalLinkDomain, linkTrust.logic.ts) and the caller (messageContent.tsx)
// renders a plain anchor for those instead — this component only ever mounts for a genuine external
// href. A domain already in the persisted trust set opens immediately on click, same as a plain anchor
// would; an unconfirmed domain intercepts the click (or middle-click) and shows the confirmation (useTrustedLinkGate),
// which persists the domain only once the user actually confirms — declining leaves the domain
// untrusted for next time too.
export function TrustedExternalLink({ href, className }: { href: string; className?: string }) {
	const gate = useTrustedLinkGate()

	// No stable domain to gate on (an unparseable href — see externalLinkDomain's own doc comment: this
	// stays reachable in principle since segments no longer arrive pre-hardened, though the one current
	// caller, messageContent.tsx, always hardens first) — render a plain, ungated anchor rather than
	// block navigation on a confirmation that could never resolve.
	if (externalLinkDomain(href) === null) {
		return (
			<ExternalAnchor
				href={href}
				className={className}
			/>
		)
	}

	return (
		<>
			<ExternalAnchor
				href={href}
				className={className}
				onClick={event => {
					if (gate.intercept(href)) {
						event.preventDefault()
					}
				}}
				onAuxClick={event => {
					// A middle-click opens the link too (in a new tab, without a click event).
					if (event.button === 1 && gate.intercept(href)) {
						event.preventDefault()
					}
				}}
			/>
			{gate.dialog}
		</>
	)
}
