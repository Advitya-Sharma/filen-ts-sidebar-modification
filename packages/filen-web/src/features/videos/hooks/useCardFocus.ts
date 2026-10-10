import { useRef } from "react"

// How long a cursor move waits for its card to mount (its row or page may still be on the way).
const FOCUS_WAIT_MS = 2000

// The card elements of a roving cursor, by key, and focusing one that may mount later: the request completes when its
// card registers, unless a newer one replaced it or it waited too long. Kept in a hook so the components holding the
// cursor never touch a ref while rendering.
export function useCardFocus<K>() {
	const elements = useRef(new Map<K, HTMLAnchorElement>())
	const request = useRef<{ key: K; scroll: boolean; until: number } | null>(null)

	function apply(element: HTMLAnchorElement, scroll: boolean): void {
		element.focus({ preventScroll: true })

		if (scroll) {
			element.scrollIntoView({ inline: "nearest", block: "nearest" })
		}
	}

	function register(key: K, element: HTMLAnchorElement | null): void {
		if (element === null) {
			elements.current.delete(key)

			return
		}

		elements.current.set(key, element)

		const pending = request.current

		if (pending?.key === key) {
			request.current = null

			if (performance.now() <= pending.until) {
				apply(element, pending.scroll)
			}
		}
	}

	function focus(key: K, options: { scroll: boolean }): void {
		const element = elements.current.get(key)

		if (element !== undefined) {
			request.current = null
			apply(element, options.scroll)

			return
		}

		request.current = { key, scroll: options.scroll, until: performance.now() + FOCUS_WAIT_MS }
	}

	return { register, focus }
}
