import { useEffect } from "react"
import { useTrustedLinkGate, type TrustedLinkGate } from "@/lib/trustedLinks/useTrustedLinkGate"

// The app's one-time-per-domain confirmation before a link opens, mounted only once the sheet shows a link (its
// trusted domains are read then, not for every sheet opened): `onGate` gets its intercept and its dialog renders here.
export function LinkTrustGate({ onGate }: { onGate: (gate: TrustedLinkGate | null) => void }) {
	const gate = useTrustedLinkGate()

	useEffect(() => {
		onGate(gate)
	})

	useEffect(
		() => () => {
			onGate(null)
		},
		[onGate]
	)

	return gate.dialog
}
