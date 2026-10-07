import type { SyntheticEvent } from "react"

// Confirmation cards act only on a click the browser itself dispatched (a pointer or a keyboard activation
// of the focused button), never on one synthesized by script. Its own module so tests can stand in for
// the browser, which is the only thing that can produce a trusted event.
export function isTrustedActivation(event: SyntheticEvent): boolean {
	return event.nativeEvent.isTrusted
}
