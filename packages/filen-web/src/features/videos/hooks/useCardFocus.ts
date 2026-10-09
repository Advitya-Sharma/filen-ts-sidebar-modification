import { useRef } from "react"

// Bounds the frames spent waiting for a card that a cursor move just mounted (a row or grid row out of range).
const FOCUS_RETRY_FRAMES = 10

// The card elements of a roving cursor, by key, and focusing one that may mount a few frames later. Kept in a hook so
// the components holding the cursor never touch a ref while rendering.
export function useCardFocus<K>() {
	const elements = useRef(new Map<K, HTMLAnchorElement>())
	const request = useRef<{ key: K } | null>(null)

	function register(key: K, element: HTMLAnchorElement | null): void {
		if (element === null) {
			elements.current.delete(key)
		} else {
			elements.current.set(key, element)
		}
	}

	function focus(key: K, options: { scroll: boolean }): void {
		const current = { key }

		request.current = current

		const attempt = (framesLeft: number): void => {
			if (request.current !== current) {
				return
			}

			const element = elements.current.get(key)

			if (element !== undefined) {
				element.focus({ preventScroll: true })

				if (options.scroll) {
					element.scrollIntoView({ inline: "nearest", block: "nearest" })
				}

				return
			}

			if (framesLeft > 0) {
				requestAnimationFrame(() => {
					attempt(framesLeft - 1)
				})
			}
		}

		attempt(FOCUS_RETRY_FRAMES)
	}

	return { register, focus }
}
